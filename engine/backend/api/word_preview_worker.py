"""A single long-lived Word process for safe, non-structural document passes."""

from __future__ import annotations

import json
import queue
import subprocess
import threading
from pathlib import Path
from typing import Optional

from ..process_utils import hidden_process_kwargs


_SCRIPT = Path(__file__).resolve().parent.parent / "scripts" / "word_preview_worker.ps1"
_START_TIMEOUT = 30
_REQUEST_TIMEOUT = 120


class WordPreviewWorker:
    """Serialize in-memory preview exports through one Word automation host.

    This deliberately does not perform the historic failed "hot edit" path:
    DOCX pagination edits still happen outside Word and are saved before any
    request reaches here.  The process only opens a finished source document,
    updates fields in memory, exports a screen PDF and closes the document.
    """

    def __init__(self) -> None:
        self._lock = threading.RLock()
        self._process: Optional[subprocess.Popen[bytes]] = None

    def export(self, input_docx: Path, output_pdf: Path,
               *, persist_fields: bool = False) -> None:
        with self._lock:
            self._ensure_started()
            try:
                self._send({"op": "preview", "input": str(input_docx),
                            "output": str(output_pdf),
                            "persist_fields": persist_fields})
            except Exception:
                self._stop_locked()
                raise
            if not output_pdf.exists():
                self._stop_locked()
                raise RuntimeError("Word preview worker returned without a PDF")

    def warm(self) -> None:
        """Start Word ahead of the first render; safe to call repeatedly."""
        with self._lock:
            self._ensure_started()

    def export_final_pdf(self, input_docx: Path, output_pdf: Path) -> None:
        """Export an already-finalized DOCX with Word's original PDF quality."""
        with self._lock:
            self._ensure_started()
            try:
                self._send({"op": "export_full", "input": str(input_docx),
                            "output": str(output_pdf)})
            except Exception:
                self._stop_locked()
                raise
            if not output_pdf.exists():
                self._stop_locked()
                raise RuntimeError("Word final-PDF worker returned without a PDF")

    def pass_docx(self, input_docx: Path, *, update_fields: bool,
                  update_page_numbers: bool, measure: bool) -> str:
        """Run Word's field/page-number update and optional page measurement."""
        with self._lock:
            self._ensure_started()
            try:
                response = self._send({
                    "op": "pass",
                    "input": str(input_docx),
                    "update_fields": update_fields,
                    "update_page_numbers": update_page_numbers,
                    "measure": measure,
                })
            except Exception:
                self._stop_locked()
                raise
        sections = response.get("sections", {})
        if not isinstance(sections, dict):
            raise RuntimeError(f"Word pass worker returned invalid sections: {sections!r}")
        lines = [f"{int(number)}:{int(page)}"
                 for number, page in sorted(sections.items(), key=lambda item: int(item[0]))]
        if measure and response.get("pages") is not None:
            lines.append(f"PAGES:{int(response['pages'])}")
        lines.append("OK")
        return "\n".join(lines)

    def close(self) -> None:
        with self._lock:
            self._stop_locked()

    def _ensure_started(self) -> None:
        if self._process is not None and self._process.poll() is None:
            return
        self._stop_locked()
        if not _SCRIPT.exists():
            raise FileNotFoundError(_SCRIPT)
        self._process = subprocess.Popen(
            ["powershell.exe", "-NoProfile", "-Sta", "-ExecutionPolicy", "Bypass",
             "-File", str(_SCRIPT)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            **hidden_process_kwargs(),
        )
        ready = self._read_response(_START_TIMEOUT)
        if ready != {"ready": True}:
            self._stop_locked()
            raise RuntimeError(f"Word preview worker did not become ready: {ready!r}")

    def _send(self, request: dict) -> dict:
        assert self._process is not None and self._process.stdin is not None
        payload = (json.dumps(request, ensure_ascii=False, separators=(",", ":")) + "\n").encode("utf-8")
        self._process.stdin.write(payload)
        self._process.stdin.flush()
        response = self._read_response(_REQUEST_TIMEOUT)
        if response.get("ok") is not True:
            raise RuntimeError(f"Word preview worker failed: {response.get('error', response)!r}")
        return response

    def _read_response(self, timeout: int) -> dict:
        assert self._process is not None and self._process.stdout is not None
        stdout = self._process.stdout
        lines: queue.Queue[bytes] = queue.Queue(maxsize=1)
        threading.Thread(target=lambda: lines.put(stdout.readline()),
                         daemon=True).start()
        try:
            line = lines.get(timeout=timeout)
        except queue.Empty as exc:
            raise subprocess.TimeoutExpired("word_preview_worker", timeout) from exc
        if not line:
            code = self._process.poll()
            raise RuntimeError(f"Word preview worker closed its pipe (exit={code})")
        try:
            return json.loads(line.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise RuntimeError(f"invalid Word preview worker response: {line!r}") from exc

    def _stop_locked(self) -> None:
        process = self._process
        self._process = None
        if process is None:
            return
        if process.stdin is not None:
            try:
                process.stdin.close()
            except OSError:
                pass
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)


WORD_PREVIEW_WORKER = WordPreviewWorker()
