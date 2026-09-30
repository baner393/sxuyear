from __future__ import annotations

import os
import subprocess
import time
from pathlib import Path

import pytest

from backend.process_utils import hidden_process_kwargs


@pytest.mark.skipif(os.name != "nt", reason="Windows-only process flags")
def test_hidden_process_kwargs_suppresses_windows_console() -> None:
    options = hidden_process_kwargs()

    assert options["creationflags"] & subprocess.CREATE_NO_WINDOW
    startupinfo = options["startupinfo"]
    assert startupinfo.dwFlags & subprocess.STARTF_USESHOWWINDOW
    assert startupinfo.wShowWindow == subprocess.SW_HIDE


@pytest.mark.skipif(os.name != "nt", reason="Windows-only window inspection")
def test_hidden_powershell_process_has_no_visible_window() -> None:
    import ctypes
    from ctypes import wintypes

    process = subprocess.Popen(
        ["powershell.exe", "-NoProfile", "-Command", "Start-Sleep -Seconds 3"],
        **hidden_process_kwargs(),
    )
    try:
        time.sleep(0.4)
        visible_windows = []
        callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        @callback_type
        def inspect_window(hwnd, _lparam):
            window_pid = wintypes.DWORD()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
            if window_pid.value == process.pid and ctypes.windll.user32.IsWindowVisible(hwnd):
                visible_windows.append(hwnd)
            return True

        assert process.poll() is None
        # EnumWindows may return zero in a non-interactive test desktop even
        # when enumeration itself found no windows, so the observable contract
        # here is the absence of a visible HWND for the still-running process.
        ctypes.windll.user32.EnumWindows(inspect_window, 0)
        assert visible_windows == []
    finally:
        process.terminate()
        process.wait(timeout=5)


def test_all_office_process_call_sites_use_hidden_options() -> None:
    backend = Path(__file__).resolve().parents[1]
    sources = (
        backend / "api" / "word_preview_worker.py",
        backend / "api" / "routes.py",
        backend / "thesis_builder" / "converter.py",
    )

    for source in sources:
        text = source.read_text(encoding="utf-8")
        assert "hidden_process_kwargs" in text, source
        assert "**hidden_process_kwargs()" in text, source
