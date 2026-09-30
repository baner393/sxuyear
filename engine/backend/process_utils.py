"""Windows child-process helpers used by the packaged desktop runtime."""

from __future__ import annotations

import os
import subprocess
from typing import Any, Dict


def hidden_process_kwargs() -> Dict[str, Any]:
    """Return subprocess options that suppress child console windows on Windows.

    The desktop application is packaged as a windowed process.  PowerShell and
    command-line office converters must remain invisible because their lifetime
    is tied to the preview/export operation; closing such a console would abort
    the user's task.
    """
    if os.name != "nt":
        return {}

    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    return {
        "creationflags": subprocess.CREATE_NO_WINDOW,
        "startupinfo": startupinfo,
    }
