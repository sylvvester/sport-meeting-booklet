from __future__ import annotations

import csv
import hashlib
import multiprocessing as mp
import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from queue import Empty


PDF_FORMAT = 17
OFFICE_PROCESS_NAMES = {"WINWORD.EXE", "WPS.EXE", "ET.EXE", "WPP.EXE"}


class OfficeExportError(RuntimeError):
    pass


def _office_pids() -> set[int]:
    """只读取 PID 和进程名，不读取命令行或打开文档。"""
    completed = subprocess.run(
        ["tasklist", "/FO", "CSV", "/NH"],
        check=False,
        capture_output=True,
        text=True,
        encoding="mbcs",
        errors="replace",
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    pids: set[int] = set()
    for row in csv.reader(completed.stdout.splitlines()):
        if len(row) >= 2 and row[0].upper() in OFFICE_PROCESS_NAMES:
            try:
                pids.add(int(row[1]))
            except ValueError:
                pass
    return pids


def _pid_from_application(application) -> int | None:
    try:
        import win32process

        hwnd = int(application.Hwnd)
        return int(win32process.GetWindowThreadProcessId(hwnd)[1])
    except Exception:
        return None


def _update_word_fields(document) -> None:
    try:
        for table_of_contents in document.TablesOfContents:
            table_of_contents.Update()
    except Exception:
        pass
    try:
        for story in document.StoryRanges:
            story.Fields.Update()
    except Exception:
        pass


def _export_worker(
    adapter: str,
    source: str,
    target: str,
    messages,
    baseline_pids: tuple[int, ...] = (),
) -> None:
    application = None
    document = None
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        prog_id = "Word.Application" if adapter == "word" else "KWPS.Application"
        application = win32com.client.DispatchEx(prog_id)
        application.Visible = False
        try:
            application.DisplayAlerts = 0
        except Exception:
            pass
        pid = _pid_from_application(application)
        if pid is None:
            new_pids = _office_pids() - set(baseline_pids)
            if len(new_pids) == 1:
                pid = new_pids.pop()
        messages.put(("pid", pid))
        document = application.Documents.Open(source, ReadOnly=True, AddToRecentFiles=False)
        _update_word_fields(document)
        try:
            document.ExportAsFixedFormat(target, PDF_FORMAT)
        except Exception:
            document.SaveAs(target, PDF_FORMAT)
        messages.put(("ok", None))
    except Exception as exc:
        messages.put(("error", f"{type(exc).__name__}: {exc}"))
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if application is not None:
            try:
                application.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _terminate_controlled_pid(pid: int, baseline: set[int]) -> None:
    """仅终止由 DispatchEx 返回、且启动前不存在的 PID。"""
    if pid in baseline or pid not in _office_pids():
        return
    subprocess.run(
        ["taskkill", "/PID", str(pid), "/T", "/F"],
        check=False,
        capture_output=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


@dataclass(frozen=True, slots=True)
class ExportOutcome:
    output_path: Path
    adapter: str
    output_name_hash: str


class ExportManager:
    def __init__(self, logger=None):
        self.logger = logger
        self._has_exported_in_session = False

    def export_pdf(self, source_docx: str | Path, output_pdf: str | Path) -> ExportOutcome:
        source = Path(source_docx).resolve()
        target = Path(output_pdf).resolve()
        if not source.is_file():
            raise FileNotFoundError(source)
        target.parent.mkdir(parents=True, exist_ok=True)
        timeout = 120 if self._has_exported_in_session else 240
        errors: list[str] = []
        for adapter in ("word", "wps"):
            try:
                self._run_adapter(adapter, source, target, timeout)
                self._has_exported_in_session = True
                name_hash = hashlib.sha256(target.name.encode("utf-8")).hexdigest()[:8]
                if self.logger:
                    self.logger.info(
                        "pdf exported",
                        fields={"adapter": adapter, "timeout_seconds": timeout, "output_name_hash": name_hash},
                    )
                return ExportOutcome(target, adapter, name_hash)
            except Exception as exc:
                errors.append(f"{adapter}: {exc}")
        raise OfficeExportError("；".join(errors))

    def _run_adapter(self, adapter: str, source: Path, target: Path, timeout: int) -> None:
        baseline = _office_pids()
        context = mp.get_context("spawn")
        messages = context.Queue()
        controlled_pid: int | None = None
        with tempfile.TemporaryDirectory(prefix="sport-meeting-export-") as temp_directory:
            temporary_pdf = Path(temp_directory) / "output.pdf"
            process = context.Process(
                target=_export_worker,
                args=(adapter, str(source), str(temporary_pdf), messages, tuple(baseline)),
                daemon=True,
            )
            process.start()
            process.join(timeout)
            while True:
                try:
                    kind, value = messages.get_nowait()
                except Empty:
                    break
                if kind == "pid" and value is not None:
                    controlled_pid = int(value)
                elif kind == "error":
                    raise OfficeExportError(str(value))
            if process.is_alive():
                process.terminate()
                process.join(5)
                if controlled_pid is not None:
                    _terminate_controlled_pid(controlled_pid, baseline)
                raise TimeoutError(f"{adapter} 导出超过 {timeout} 秒")
            if process.exitcode != 0:
                raise OfficeExportError(f"{adapter} 导出进程异常退出：{process.exitcode}")
            if not temporary_pdf.is_file() or temporary_pdf.stat().st_size == 0:
                raise OfficeExportError(f"{adapter} 未生成有效 PDF")
            shutil.copy2(temporary_pdf, target)
