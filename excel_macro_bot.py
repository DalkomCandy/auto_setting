"""엑셀 매크로 일괄 실행 봇 (Windows 전용)

지정한 폴더 아래의 엑셀 파일을 하나씩 열어서
  1. 대상 셀(기본 B2)을 선택하고
  2. 빠른 실행 도구 모음에 등록된 매크로를 실행하고
  3. 매크로가 띄우는 InputBox 에 셀 값을 입력한 뒤 [확인]을 누르고
  4. 파일을 덮어쓰기 저장
하는 작업을 반복한다.

InputBox 는 모달 창이라 Application.Run 이 반환되지 않는다.
그래서 매크로 실행 직전에 감시 스레드를 띄워 창이 뜨는 즉시
win32 메시지로 값을 채우고 확인을 누르는 구조다.

사용 예:
    python excel_macro_bot.py --root "D:\\all" --macro "PERSONAL.XLSB!수익_개요" --dry-run
    python excel_macro_bot.py --root "D:\\all" --macro "PERSONAL.XLSB!수익_개요" --probe
    python excel_macro_bot.py --root "D:\\all" --macro "PERSONAL.XLSB!수익_개요"
"""
from __future__ import annotations

import argparse
import csv
import os
import re
import sys
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import win_dialog as wd
from win_dialog import IS_WINDOWS

if IS_WINDOWS:
    import pythoncom
    import win32com.client
    import win32process
    from pywintypes import com_error
else:  # 문법 검사/편집 목적으로만 임포트 가능하게
    com_error = Exception

#: Sub/Function 선언에서 프로시저 이름을 뽑는 정규식 (--list-macros)
PROC_RE = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function)\s+([^\s(]+)",
    re.MULTILINE | re.IGNORECASE,
)

MSO_AUTOMATION_SECURITY_LOW = 1


class MacroError(RuntimeError):
    """매크로 이름 해석/실행 실패"""


# ── 로깅 ────────────────────────────────────────────
class Logger:
    """콘솔 출력용. `sink` 를 주면 GUI 등 다른 곳으로 메시지를 넘긴다."""

    def __init__(self, verbose: bool = True, sink=None):
        self.verbose = verbose
        self.sink = sink
        self._lock = threading.Lock()

    def __call__(self, message: str, indent: int = 0) -> None:
        text = f"{'  ' * indent}{message}"
        with self._lock:
            if self.sink is not None:
                self.sink(text)
            else:
                print(text, flush=True)

    def debug(self, message: str, indent: int = 0) -> None:
        if self.verbose:
            self(message, indent)


# ── InputBox 감시 스레드 ────────────────────────────
@dataclass
class DialogOptions:
    value: str
    title_contains: str = ""
    poll_interval: float = 0.12
    appear_timeout: float = 30.0
    confirm_timeout: float = 10.0
    input_method: str = "settext"      # settext | chars
    dismiss_followup: bool = True
    probe: bool = False
    #: 값을 넣을 입력칸 순번. -1 이면 자동(비어 있는 첫 칸)
    field_index: int = -1


class DialogHandler(threading.Thread):
    """매크로가 띄우는 대화상자를 감시하다 값을 넣고 확인을 누른다."""

    def __init__(self, pid: int, opts: DialogOptions, log: Logger):
        super().__init__(daemon=True, name="dialog-handler")
        self.pid = pid
        self.opts = opts
        self.log = log
        # Thread 내부에 이미 _stop() 메서드가 있으므로 이름이 겹치면 join() 이 깨진다
        self._stop_event = threading.Event()
        self.filled = False
        self.error: Optional[str] = None
        self.probe_report: List[str] = []
        self._handled: set = set()

    def stop(self) -> None:
        self._stop_event.set()

    def run(self) -> None:
        started = time.monotonic()
        while not self._stop_event.is_set():
            hwnd = self._next_dialog()
            if hwnd is None:
                if not self.filled and time.monotonic() - started > self.opts.appear_timeout:
                    self.error = self.error or (
                        f"InputBox 창이 {self.opts.appear_timeout:.0f}초 안에 나타나지 않았습니다"
                    )
                    # 매크로가 계속 돌 수도 있으므로 감시는 유지한다
                    started = time.monotonic()
                time.sleep(self.opts.poll_interval)
                continue
            try:
                self._handle(hwnd)
            except Exception as exc:                      # 감시 스레드는 절대 죽지 않게
                self.error = self.error or f"대화상자 처리 실패: {exc}"
                self._handled.add(hwnd)

    def _next_dialog(self) -> Optional[int]:
        for hwnd in wd.find_dialogs(self.pid, self.opts.title_contains):
            if hwnd not in self._handled:
                return hwnd
        return None

    def _handle(self, hwnd: int) -> None:
        controls = self._settled_controls(hwnd)
        if not wd.is_alive(hwnd):        # 처리 전에 스스로 닫힌 창
            self._handled.add(hwnd)
            return

        if self.opts.probe:
            self.probe_report.append(wd.format_dialog_report(hwnd, controls))
            wd.cancel_dialog(hwnd, controls)
            self._handled.add(hwnd)
            return

        edits = [c for c in controls if c.is_edit]
        target = wd.pick_edit(controls, self.opts.field_index) if edits else None

        if target is not None and not self.filled:
            self._fill(hwnd, target, controls)
        elif not edits and self.filled and self.opts.dismiss_followup:
            method = wd.confirm_dialog(hwnd, controls, self.opts.confirm_timeout)
            self.log.debug(
                f"후속 안내창 확인 ({method or '실패'}): {wd.window_title(hwnd)!r}", indent=2
            )
        else:
            title = wd.window_title(hwnd)
            self.error = self.error or f"예상하지 못한 대화상자입니다: {title!r}"
            self.log(f"⚠ 예상 밖의 창을 취소합니다: {title!r}", indent=2)
            wd.cancel_dialog(hwnd, controls)

        self._handled.add(hwnd)

    def _settled_controls(self, hwnd: int):
        """창이 막 뜬 직후엔 컨트롤이 아직 없을 수 있어 잠깐 재시도한다."""
        controls = []
        for _ in range(6):
            controls = wd.list_controls(hwnd)
            if any(c.is_edit for c in controls) or any(c.is_button for c in controls):
                break
            time.sleep(0.08)
        return controls

    def _fill(self, hwnd: int, edit, controls) -> None:
        if self.opts.input_method == "chars":
            wd.type_text(edit.hwnd, self.opts.value)
        else:
            wd.set_text(edit.hwnd, self.opts.value)

        written = wd.get_text(edit.hwnd)
        if written != self.opts.value:
            # WM_SETTEXT 를 무시하는 컨트롤이면 타이핑 방식으로 한 번 더
            wd.type_text(edit.hwnd, self.opts.value)
            written = wd.get_text(edit.hwnd)

        if written != self.opts.value:
            self.error = (
                f"입력값을 넣지 못했습니다 (기대 {self.opts.value!r}, 실제 {written!r})"
            )
            wd.cancel_dialog(hwnd, controls)
            return

        method = wd.confirm_dialog(hwnd, controls, self.opts.confirm_timeout)
        if not method:
            self.error = "확인 버튼을 누르지 못했습니다"
            wd.cancel_dialog(hwnd, controls)
            return

        self.filled = True
        self.log.debug(f"InputBox 입력 {self.opts.value!r} → {method}", indent=2)


# ── 매크로 실행기 ───────────────────────────────────
class MacroRunner:
    """입력된 이름 그대로 Application.Run 을 호출한다.

    자동으로 이름을 추측하지 않는다 — VBA 프로시저 이름에는 공백을 쓸 수
    없어 빠른 실행 버튼 이름("수익 개요")과 실제 매크로 이름("수익_개요")이
    다른 경우가 흔하고, 후보를 여러 개 시도하면 의도와 다른 매크로가
    조용히 실행될 위험이 있다. 그래서 사용자가 [목록 불러오기]로 정확한
    이름을 확인해 직접 입력한 값을 그대로 사용한다.
    """

    def __init__(self, name: str):
        self.name = name.strip()

    def run(self, xl) -> str:
        if not self.name:
            raise MacroError("매크로 이름이 비어 있습니다.")
        try:
            xl.Run(self.name)
        except com_error as exc:
            detail = describe_com_error(exc)
            raise MacroError(
                f"매크로 {self.name!r} 실행 실패: {detail}\n"
                "[목록 불러오기]로 정확한 이름을 확인해 그대로 입력하세요 "
                "(예: PERSONAL.XLSB!수익_개요)."
            ) from exc
        return self.name


def describe_com_error(exc: BaseException) -> str:
    info = getattr(exc, "excepinfo", None)
    if info and len(info) > 2 and info[2]:
        return str(info[2]).strip()
    return str(exc)


# ── Excel 헬퍼 ──────────────────────────────────────
def default_personal_path() -> Optional[Path]:
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return None
    return Path(appdata) / "Microsoft" / "Excel" / "XLSTART" / "PERSONAL.XLSB"


def start_excel(visible: bool = True):
    """자동화 전용 Excel 인스턴스를 새로 띄운다."""
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = visible
    xl.DisplayAlerts = True          # InputBox 가 정상적으로 뜨도록 유지
    try:
        xl.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
    except com_error:
        pass                          # 정책상 변경 불가한 환경도 있다
    return xl


def excel_pid(xl) -> int:
    _, pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
    return pid


def open_macro_host(xl, path: Optional[Path], log: Logger):
    """매크로가 들어 있는 개인용 매크로 통합 문서/추가 기능을 연다.

    자동화로 띄운 Excel 은 XLSTART 폴더를 자동으로 읽지 않으므로
    PERSONAL.XLSB 를 직접 열어줘야 매크로를 호출할 수 있다.
    """
    if path is None:
        return None
    if not path.exists():
        log(f"⚠ 매크로 파일을 찾을 수 없습니다: {path}")
        return None

    for book in xl.Workbooks:
        if book.Name.casefold() == path.name.casefold():
            log(f"매크로 파일 이미 열림: {book.Name}")
            return book

    book = xl.Workbooks.Open(str(path), UpdateLinks=0)
    log(f"매크로 파일 열기: {book.Name}")
    return book


VBA_ACCESS_HINT = (
    "VBA 프로젝트에 접근할 수 없습니다. Excel > 파일 > 옵션 > 보안 센터 > "
    "보안 센터 설정 > 매크로 설정에서 'VBA 프로젝트 개체 모델에 안전하게 접근할 수 있음'을 "
    "켜거나, Alt+F11 로 직접 이름을 확인하세요."
)


def collect_macro_names(xl) -> tuple:
    """열려 있는 통합 문서의 Sub/Function 이름을 모은다.

    Returns: (["PERSONAL.XLSB!수익_개요", ...], ["경고 메시지", ...])
    """
    names: List[str] = []
    warnings: List[str] = []

    for book in xl.Workbooks:
        try:
            components = list(book.VBProject.VBComponents)
        except com_error as exc:
            warnings.append(f"{book.Name}: {VBA_ACCESS_HINT} ({describe_com_error(exc)})")
            continue

        for comp in components:
            module = comp.CodeModule
            count = module.CountOfLines
            code = module.Lines(1, count) if count else ""
            for match in PROC_RE.finditer(code):
                names.append(f"{book.Name}!{match.group(2)}")
    return names, warnings


def list_macros(xl, log: Logger) -> None:
    """CLI 용 출력 래퍼."""
    names, warnings = collect_macro_names(xl)
    for warning in warnings:
        log(f"⚠ {warning}")
    if not names:
        log("사용 가능한 매크로를 찾지 못했습니다.")
        return
    log("")
    for name in names:
        log(f"  {name}")


def pick_sheet(book, name: Optional[str]):
    if name:
        return book.Worksheets(name)
    return book.ActiveSheet


def read_cell(sheet, address: str, source: str) -> str:
    """셀 값을 InputBox 에 넣을 문자열로 변환한다.

    기본값은 화면에 보이는 그대로(`Text`)다. 종목아이디처럼 `005930` 형태의
    앞자리 0이 있는 값은 `Value` 로 읽으면 5930 이 되어버리기 때문이다.
    다만 열 너비가 좁으면 `Text` 가 `#####` 이 되므로 그 경우엔 `Value` 를 쓴다.
    """
    cell = sheet.Range(address)
    if source == "text":
        text = str(cell.Text or "").strip()
        if text and set(text) != {"#"}:
            return text
    return _stringify(cell.Value)


def _stringify(raw) -> str:
    if raw is None:
        return ""
    if isinstance(raw, bool):
        return str(raw)
    if isinstance(raw, float) and raw.is_integer():
        return str(int(raw))
    return str(raw).strip()


# ── 파일 수집 ───────────────────────────────────────
def collect_files(root: Path, prefix: str, pattern: str) -> List[Path]:
    """all폴더 > 개별 폴더 > 26.3Q*.xlsx 구조를 재귀적으로 훑는다."""
    files = []
    for path in root.rglob(pattern):
        if not path.is_file():
            continue
        if path.name.startswith("~$"):       # Excel 임시 잠금 파일
            continue
        if prefix and not path.name.casefold().startswith(prefix.casefold()):
            continue
        files.append(path)
    return sorted(files)


# ── 처리 결과 ───────────────────────────────────────
@dataclass
class Result:
    path: Path
    status: str                  # ok | skipped | failed
    value: str = ""
    detail: str = ""


@dataclass
class Settings:
    root: Path
    macro: str
    prefix: str = "26.3Q"
    pattern: str = "*.xlsx"
    target_cell: str = "B2"
    value_cell: str = "B2"
    value_source: str = "text"
    sheet: Optional[str] = None
    dry_run: bool = False
    probe: bool = False
    limit: int = 0
    dialog: DialogOptions = field(default_factory=lambda: DialogOptions(value=""))


# ── 파일 1건 처리 ───────────────────────────────────
def process_file(xl, path: Path, cfg: Settings, runner: MacroRunner,
                 pid: int, log: Logger) -> Result:
    book = None
    try:
        xl.DisplayAlerts = False
        book = xl.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
        xl.DisplayAlerts = True

        sheet = pick_sheet(book, cfg.sheet)
        book.Activate()
        sheet.Activate()

        value = read_cell(sheet, cfg.value_cell, cfg.value_source)
        if not value:
            return Result(path, "skipped", detail=f"{cfg.value_cell} 셀이 비어 있음")

        sheet.Range(cfg.target_cell).Select()
        log(f"{cfg.target_cell} 선택, 입력값 {value!r}", indent=1)

        if cfg.dry_run:
            return Result(path, "skipped", value=value, detail="dry-run (매크로 미실행)")

        opts = DialogOptions(**{**cfg.dialog.__dict__, "value": value})
        watcher = DialogHandler(pid, opts, log)
        watcher.start()
        try:
            resolved = runner.run(xl)
        finally:
            watcher.stop()
            watcher.join(timeout=5)

        if cfg.probe:
            report = "\n".join(watcher.probe_report) or "  (대화상자를 보지 못했습니다)"
            log(f"매크로 {resolved!r} 실행 중 나타난 대화상자:", indent=1)
            log(report)
            return Result(path, "skipped", value=value, detail="probe (저장하지 않음)")

        if watcher.error:
            return Result(path, "failed", value=value, detail=watcher.error)
        if not watcher.filled:
            return Result(path, "failed", value=value,
                          detail="InputBox 창을 찾지 못했습니다")

        book.Save()
        return Result(path, "ok", value=value, detail=f"매크로 {resolved}")

    except MacroError:
        raise                                  # 이름 해석 실패는 전체 중단 사유
    except com_error as exc:
        return Result(path, "failed", detail=describe_com_error(exc))
    except Exception as exc:
        return Result(path, "failed", detail=str(exc))
    finally:
        if book is not None:
            try:
                xl.DisplayAlerts = False
                book.Close(SaveChanges=False)   # 성공 시엔 위에서 이미 저장했다
            except com_error:
                pass
            finally:
                xl.DisplayAlerts = True


# ── 실행 ────────────────────────────────────────────
def run(cfg: Settings, personal: Optional[Path], keep_open: bool, log: Logger,
        files: Optional[List[Path]] = None, cancel=None,
        on_file=None, on_result=None) -> List[Result]:
    """대상 파일을 순회하며 처리한다.

    files     : 미리 스캔해 둔 목록 (없으면 직접 스캔)
    cancel    : threading.Event — 파일 사이에서 중단 여부를 확인
    on_file   : (index, total, path) 처리 시작 알림
    on_result : (index, total, path, result) 처리 완료 알림
    """
    if files is None:
        files = collect_files(cfg.root, cfg.prefix, cfg.pattern)
        if cfg.limit:
            files = files[: cfg.limit]

    if not files:
        log(f"대상 파일이 없습니다: {cfg.root} 아래 {cfg.prefix!r} 로 시작하는 {cfg.pattern}")
        return []

    total = len(files)
    log(f"대상 파일 {total}개")
    for path in files:
        log.debug(f"  {path.relative_to(cfg.root)}")
    log("")

    pythoncom.CoInitialize()
    xl = None
    results: List[Result] = []
    try:
        xl = start_excel()
        host = open_macro_host(xl, personal, log)
        runner = MacroRunner(cfg.macro)
        pid = excel_pid(xl)

        for index, path in enumerate(files, start=1):
            if cancel is not None and cancel.is_set():
                log(f"\n사용자 요청으로 중단했습니다 ({index - 1}/{total} 처리됨)")
                break

            log(f"[{index}/{total}] {path.name}")
            if on_file is not None:
                on_file(index, total, path)

            result = process_file(xl, path, cfg, runner, pid, log)
            results.append(result)

            icon = {"ok": "✓", "skipped": "-", "failed": "❌"}[result.status]
            log(f"{icon} {result.status}: {result.detail}", indent=1)
            if on_result is not None:
                on_result(index, total, path, result)
    finally:
        if xl is not None and not keep_open:
            try:
                xl.DisplayAlerts = False
                xl.Quit()
            except com_error:
                pass
        pythoncom.CoUninitialize()
    return results


def write_log(results: List[Result], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["파일", "경로", "상태", "입력값", "비고"])
        for item in results:
            writer.writerow([item.path.name, str(item.path), item.status,
                             item.value, item.detail])


def summarize(results: List[Result], log: Logger) -> int:
    counts = {"ok": 0, "skipped": 0, "failed": 0}
    for item in results:
        counts[item.status] += 1

    log("\n" + "─" * 50)
    log(f"완료 {counts['ok']}건 / 건너뜀 {counts['skipped']}건 / 실패 {counts['failed']}건")
    if counts["failed"]:
        log("\n실패 목록:")
        for item in results:
            if item.status == "failed":
                log(f"  ❌ {item.path.name}: {item.detail}")
    return 1 if counts["failed"] else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="폴더 안의 엑셀 파일을 열어 셀 선택 → 매크로 실행 → InputBox 입력을 자동화합니다.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--root", type=Path, help="대상 최상위 폴더 (all폴더)")
    parser.add_argument("--macro", default="",
                        help="실행할 매크로 이름 (정확한 전체 이름, 예: PERSONAL.XLSB!수익_개요). "
                             "자동으로 추측하지 않으므로 --list-macros 로 확인한 이름을 그대로 입력하세요.")
    parser.add_argument("--prefix", default="26.3Q",
                        help="파일명 접두사 필터 (기본: %(default)s, 빈 문자열이면 전체)")
    parser.add_argument("--pattern", default="*.xlsx",
                        help="파일 glob 패턴 (기본: %(default)s)")
    parser.add_argument("--cell", default="B2", help="선택할 셀 (기본: %(default)s)")
    parser.add_argument("--value-cell", default=None,
                        help="InputBox 에 넣을 값을 읽을 셀 (기본: --cell 과 동일)")
    parser.add_argument("--value-source", choices=["text", "value"], default="text",
                        help="text=화면 표시값(앞자리 0 유지), value=원본 값 (기본: %(default)s)")
    parser.add_argument("--sheet", default=None, help="시트 이름 (기본: 열었을 때 활성 시트)")

    parser.add_argument("--personal", type=Path, default=None,
                        help="매크로가 있는 파일 경로 (기본: XLSTART\\PERSONAL.XLSB 자동 탐색)")
    parser.add_argument("--no-personal", action="store_true",
                        help="PERSONAL.XLSB 를 열지 않음")

    parser.add_argument("--dry-run", action="store_true",
                        help="파일을 열어 셀 값만 확인하고 매크로는 실행하지 않음")
    parser.add_argument("--probe", action="store_true",
                        help="매크로를 실행해 대화상자 구조만 출력하고 취소 (저장 안 함)")
    parser.add_argument("--list-macros", action="store_true",
                        help="사용 가능한 매크로 이름을 출력하고 종료")
    parser.add_argument("--limit", type=int, default=0, help="앞에서 N개만 처리 (0=전체)")

    parser.add_argument("--dialog-title", default="",
                        help="이 문자열이 제목에 포함된 창만 대상으로 함")
    parser.add_argument("--dialog-timeout", type=float, default=30.0,
                        help="InputBox 가 뜨기를 기다리는 시간(초) (기본: %(default)s)")
    parser.add_argument("--confirm-timeout", type=float, default=10.0,
                        help="확인 클릭 후 창이 닫히기를 기다리는 시간(초) (기본: %(default)s)")
    parser.add_argument("--poll", type=float, default=0.12,
                        help="창 감시 주기(초) (기본: %(default)s)")
    parser.add_argument("--input-method", choices=["settext", "chars"], default="settext",
                        help="입력 방식. settext 가 안 먹으면 chars 로 (기본: %(default)s)")
    parser.add_argument("--no-dismiss-followup", action="store_true",
                        help="매크로가 띄우는 후속 안내창을 자동으로 확인하지 않음")

    parser.add_argument("--log", default="", help="결과 CSV 경로 (기본: 자동 생성, none=끄기)")
    parser.add_argument("--keep-open", action="store_true",
                        help="작업 후 Excel 을 종료하지 않음 (디버깅용)")
    parser.add_argument("--quiet", action="store_true", help="상세 로그 숨김")
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    log = Logger(verbose=not args.quiet)

    if not IS_WINDOWS:
        log("❌ 이 스크립트는 Excel COM 자동화를 사용하므로 Windows 에서만 동작합니다.")
        return 2

    personal = None if args.no_personal else (args.personal or default_personal_path())

    if args.list_macros:
        pythoncom.CoInitialize()
        xl = None
        try:
            xl = start_excel(visible=False)
            open_macro_host(xl, personal, log)
            list_macros(xl, log)
        finally:
            if xl is not None:
                try:
                    xl.Quit()
                except com_error:
                    pass
            pythoncom.CoUninitialize()
        return 0

    if args.root is None:
        log("❌ --root 로 대상 폴더를 지정하세요. (--list-macros 제외)")
        return 2
    root = args.root.expanduser()
    if not root.is_dir():
        log(f"❌ 폴더를 찾을 수 없습니다: {root}")
        return 2
    if not args.macro.strip() and not args.dry_run:
        log("❌ --macro 로 실행할 매크로의 정확한 이름을 지정하세요. "
            "(--list-macros 로 먼저 확인하세요. --dry-run 은 매크로 없이도 가능합니다.)")
        return 2

    cfg = Settings(
        root=root,
        macro=args.macro,
        prefix=args.prefix,
        pattern=args.pattern,
        target_cell=args.cell,
        value_cell=args.value_cell or args.cell,
        value_source=args.value_source,
        sheet=args.sheet,
        dry_run=args.dry_run,
        probe=args.probe,
        limit=max(0, args.limit),
        dialog=DialogOptions(
            value="",
            title_contains=args.dialog_title,
            poll_interval=args.poll,
            appear_timeout=args.dialog_timeout,
            confirm_timeout=args.confirm_timeout,
            input_method=args.input_method,
            dismiss_followup=not args.no_dismiss_followup,
            probe=args.probe,
        ),
    )

    if cfg.dry_run:
        log("※ dry-run 모드: 매크로를 실행하지 않고 셀 값만 확인합니다.\n")
    elif cfg.probe:
        log("※ probe 모드: 대화상자 구조만 확인하고 파일을 저장하지 않습니다.\n")
    else:
        log("※ 실행 중에는 Excel 창을 건드리지 마세요.\n")

    try:
        results = run(cfg, personal, args.keep_open, log)
    except MacroError as exc:
        log(f"\n❌ {exc}")
        return 1
    except KeyboardInterrupt:
        log("\n중단되었습니다.")
        return 130

    if not results:
        return 0

    if args.log.casefold() != "none":
        target = Path(args.log) if args.log else Path(
            f"macro_bot_log_{datetime.now():%Y%m%d_%H%M%S}.csv"
        )
        write_log(results, target)
        log(f"\n결과 기록: {target}")

    return summarize(results, log)


if __name__ == "__main__":
    sys.exit(main())
