"""백그라운드 작업 스레드

Excel COM 은 스레드마다 CoInitialize 가 필요하고, 매크로 실행은 InputBox 때문에
길게 블로킹되므로 UI 스레드에서 직접 돌리면 창이 얼어붙는다.
그래서 모든 Excel 작업을 QThread 로 분리하고 결과만 시그널로 넘긴다.
"""
from __future__ import annotations

import threading
from pathlib import Path
from typing import List, Optional

from PyQt6.QtCore import QThread, pyqtSignal

import excel_macro_bot as bot


def classify(message: str) -> str:
    """엔진 로그 문자열에서 콘솔 색상용 레벨을 추론한다."""
    stripped = message.lstrip()
    if stripped.startswith("❌"):
        return "error"
    if stripped.startswith("⚠"):
        return "warning"
    if stripped.startswith("✓"):
        return "success"
    return "info"


class ScanWorker(QThread):
    """대상 파일 목록만 수집한다. 네트워크 드라이브에서 느릴 수 있어 분리."""

    found = pyqtSignal(list)
    failed = pyqtSignal(str)

    def __init__(self, root: Path, prefix: str, pattern: str, limit: int = 0, parent=None):
        super().__init__(parent)
        self.root, self.prefix, self.pattern, self.limit = root, prefix, pattern, limit

    def run(self) -> None:
        try:
            files = bot.collect_files(self.root, self.prefix, self.pattern)
            if self.limit:
                files = files[: self.limit]
            self.found.emit(files)
        except Exception as exc:
            self.failed.emit(str(exc))


class MacroListWorker(QThread):
    """PERSONAL.XLSB 등에서 매크로 이름 목록을 읽어온다."""

    listed = pyqtSignal(list, list)      # names, warnings
    failed = pyqtSignal(str)

    def __init__(self, personal: Optional[Path], parent=None):
        super().__init__(parent)
        self.personal = personal

    def run(self) -> None:
        import pythoncom
        from pywintypes import com_error

        pythoncom.CoInitialize()
        xl = None
        try:
            messages: List[str] = []
            log = bot.Logger(sink=messages.append)
            xl = bot.start_excel(visible=False)
            bot.open_macro_host(xl, self.personal, log)
            names, warnings = bot.collect_macro_names(xl)
            self.listed.emit(names, warnings + [m for m in messages if m.startswith("⚠")])
        except Exception as exc:
            self.failed.emit(str(exc))
        finally:
            if xl is not None:
                try:
                    xl.Quit()
                except com_error:
                    pass
            pythoncom.CoUninitialize()


class PipelineWorker(QThread):
    """11단계 파이프라인 실행."""

    message = pyqtSignal(str, str)               # text, level
    file_started = pyqtSignal(int, int, object)  # index, total, Path
    file_done = pyqtSignal(int, int, object)     # index, total, FileResult
    completed = pyqtSignal(list)                 # List[FileResult]
    failed = pyqtSignal(str)

    def __init__(self, cfg, inputs, dialog, files: List[Path],
                 personal: Optional[Path], stop_after: int, keep_open: bool, parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self.inputs = inputs
        self.dialog = dialog
        self.files = files
        self.personal = personal
        self.stop_after = stop_after
        self.keep_open = keep_open
        self.cancel = threading.Event()

    def stop(self) -> None:
        """다음 파일로 넘어가기 전에 멈춘다 (진행 중인 파일은 끝까지 처리)."""
        self.cancel.set()

    def run(self) -> None:
        import fund_pipeline as fp

        log = bot.Logger(sink=lambda text: self.message.emit(text, classify(text)))
        try:
            results = fp.run_pipeline(
                root=self.files[0].parent if self.files else None,
                files=self.files,
                cfg=self.cfg,
                inputs=self.inputs,
                dialog=self.dialog,
                log=log,
                personal=self.personal,
                stop_after=self.stop_after,
                keep_open=self.keep_open,
                cancel=self.cancel,
                on_file=lambda i, n, p: self.file_started.emit(i, n, p),
                on_result=lambda i, n, p, r: self.file_done.emit(i, n, r),
            )
            self.completed.emit(results)
        except (fp.StepError, bot.MacroError) as exc:
            self.failed.emit(str(exc))
        except Exception as exc:
            self.failed.emit(f"{type(exc).__name__}: {exc}")


class BotWorker(QThread):
    """실제 매크로 일괄 실행. dry-run / probe 모드도 이 워커가 처리한다."""

    message = pyqtSignal(str, str)               # text, level
    file_started = pyqtSignal(int, int, object)  # index, total, Path
    file_done = pyqtSignal(int, int, object)     # index, total, Result
    completed = pyqtSignal(list)                 # List[Result]
    failed = pyqtSignal(str)

    def __init__(self, cfg: bot.Settings, personal: Optional[Path], keep_open: bool,
                 files: Optional[List[Path]] = None, parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self.personal = personal
        self.keep_open = keep_open
        self.files = files
        self.cancel = threading.Event()

    def stop(self) -> None:
        """다음 파일로 넘어가기 전에 멈춘다 (진행 중인 파일은 끝까지 처리)."""
        self.cancel.set()

    def run(self) -> None:
        log = bot.Logger(sink=lambda text: self.message.emit(text, classify(text)))
        try:
            results = bot.run(
                self.cfg, self.personal, self.keep_open, log,
                files=self.files,
                cancel=self.cancel,
                on_file=lambda i, n, p: self.file_started.emit(i, n, p),
                on_result=lambda i, n, p, r: self.file_done.emit(i, n, r),
            )
            self.completed.emit(results)
        except bot.MacroError as exc:
            self.failed.emit(str(exc))
        except Exception as exc:
            self.failed.emit(f"{type(exc).__name__}: {exc}")
