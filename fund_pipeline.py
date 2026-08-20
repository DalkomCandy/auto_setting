"""수익 보고서 생성 파이프라인 (11단계)

엑셀 파일 하나에 대해 아래 순서로 작업한다.

  1. 이름 F_1 → 오른쪽 1 · 아래 3 칸에서 펀드아이디 읽기
  2. 이름 {펀드아이디}_nav → 오른쪽 1 · 아래 1 칸에서 장부기준일 읽기
  3. 이름 {펀드아이디}_inv → 아래 5번째부터 마지막까지 종목아이디 수집
  4. F_1 이 있는 시트 내용 전체 삭제 + 이름 관리자의 모든 이름 삭제
  5. 숨겨진 시트 ALT_ITEMS 삭제
  6. B2 선택 → '수익 개요' 매크로 → 펀드아이디 입력 → [생성]
  7. E5:E9 에 평가기준일·변수일·장부기준일·평가자·리뷰어 입력
  8. B16 선택 → '수익 NAV' 매크로 → 펀드아이디 입력
  9. H16 선택 → '수익 투자자산' 매크로 → 펀드아이디 입력
 10. H21 부터 아래로 종목아이디 붙여넣기
 11. 마지막 종목아이디 아래 3칸 → '장부가 미수이자' 매크로 → 펀드아이디 입력

각 단계는 이름과 함께 결과가 기록되므로, 실패하면 몇 단계에서 왜 멈췄는지 알 수 있다.
1~3 단계는 파일을 전혀 건드리지 않는 읽기 전용이라 미리보기로 쓸 수 있다.
"""
from __future__ import annotations

import re
import time
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Callable, List, Optional, Tuple

import excel_macro_bot as bot
from excel_macro_bot import DialogHandler, DialogOptions, Logger, MacroRunner

if bot.IS_WINDOWS:
    import pythoncom
    from pywintypes import com_error
else:
    com_error = Exception

XL_DOWN = -4121
XL_SHEET_VISIBLE = -1
XL_SHIFT_UP = -4162
COLOR_WHITE = 0xFFFFFF

#: 이 행 아래로는 데이터가 아니라 빈 시트로 본다 (End(xlDown) 폭주 방지)
MAX_DATA_ROW = 1_000_000

#: 읽기 전용으로 파일을 건드리지 않는 단계 수
READONLY_STEPS = 3


class StepError(Exception):
    """단계 실패. 메시지가 그대로 사용자에게 보고되고 이후 단계는 중단된다."""


class StepSkip(Exception):
    """할 일이 없어 건너뛴다. 실패가 아니므로 다음 단계로 계속 진행한다.

    예) 종목아이디가 0개인 펀드 → 붙여넣기(10)와 미수이자(11)는 건너뛰지만
        수익 개요·NAV·투자자산 보고서는 그대로 만들고 저장한다.
    """


@dataclass
class FundInputs:
    """프로그램 시작 시 한 번만 입력받는 값 (모든 파일 공통)"""
    eval_date: str = ""      # 평가기준일  → E5
    var_date: str = ""       # 변수일      → E6
    evaluator: str = ""      # 평가자      → E8
    reviewer: str = ""       # 리뷰어      → E9

    def missing(self) -> List[str]:
        labels = {"eval_date": "평가기준일", "var_date": "변수일",
                  "evaluator": "평가자", "reviewer": "리뷰어"}
        return [label for key, label in labels.items() if not getattr(self, key).strip()]


@dataclass
class PipelineConfig:
    """시트 구조와 매크로 이름. 파일 양식이 바뀌면 여기만 고치면 된다."""
    anchor_name: str = "F_1"
    fund_offset: Tuple[int, int] = (3, 1)       # F_1 기준 (아래, 오른쪽)
    fund_prefix: str = "F"                      # 빈 문자열이면 검사 안 함

    nav_suffix: str = "_nav"
    nav_offset: Tuple[int, int] = (1, 1)

    inv_suffix: str = "_inv"
    inv_first_row_offset: int = 5               # _inv 아래 5번째가 첫 종목

    alt_sheet: str = "ALT_ITEMS"
    #: delete = Ctrl+A → 우클릭 → 삭제 (셀 자체를 삭제)
    #: clear  = 내용과 서식만 지우고 셀은 그대로 둠
    clear_mode: str = "delete"

    overview_cell: str = "B2"
    header_cell: str = "E5"                     # E5:E9 로 5칸 사용
    nav_cell: str = "B16"
    inv_cell: str = "H16"
    items_cell: str = "H21"
    interest_row_offset: int = 3                # 마지막 종목 아래 3칸

    macro_overview: str = "수익 개요"
    macro_nav: str = "수익 NAV"
    macro_inv: str = "수익 투자자산"
    macro_interest: str = "장부가 미수이자"

    def macro_names(self) -> List[str]:
        return [self.macro_overview, self.macro_nav, self.macro_inv, self.macro_interest]


@dataclass
class StepResult:
    number: int
    name: str
    status: str                  # ok | failed | skipped
    detail: str = ""
    elapsed: float = 0.0

    @property
    def label(self) -> str:
        return f"{self.number}. {self.name}"


@dataclass
class FileResult:
    path: Path
    status: str = "failed"       # ok | failed | skipped
    fund_id: str = ""
    book_date: str = ""
    item_count: int = 0
    saved: bool = False
    steps: List[StepResult] = field(default_factory=list)

    @property
    def items_counted(self) -> bool:
        """3단계까지 갔는지. 0개와 '아직 못 읽음'을 구분하려고 쓴다."""
        return any(s.number == 3 and s.status == "ok" for s in self.steps)

    @property
    def failed_step(self) -> str:
        for step in self.steps:
            if step.status == "failed":
                return step.label
        return ""

    @property
    def detail(self) -> str:
        for step in self.steps:
            if step.status == "failed":
                return f"[{step.label}] {step.detail}"
        done = sum(1 for s in self.steps if s.status == "ok")
        skipped = sum(1 for s in self.steps if s.status == "skipped")
        parts = [f"{done}단계 완료"]
        if skipped:
            parts.append(f"{skipped}단계 건너뜀")
        parts.append("저장됨" if self.saved else "저장 안 함")
        return " · ".join(parts)


# ── 셀 값 도우미 ────────────────────────────────────
def cell_text(cell) -> str:
    """화면에 보이는 값. 열이 좁아 ##### 이면 원본으로 대체."""
    text = str(cell.Text or "").strip()
    if text and set(text) != {"#"}:
        return text
    return stringify(cell.Value)


def stringify(raw) -> str:
    if raw is None:
        return ""
    if isinstance(raw, bool):
        return str(raw)
    if isinstance(raw, (datetime, date)):
        return raw.strftime("%Y-%m-%d")
    if isinstance(raw, float) and raw.is_integer():
        return str(int(raw))
    return str(raw).strip()


def as_date_text(cell) -> str:
    """장부기준일을 yyyy-mm-dd 문자열로."""
    raw = cell.Value
    if isinstance(raw, (datetime, date)):
        return raw.strftime("%Y-%m-%d")
    text = cell_text(cell)
    match = re.search(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", text)
    if match:
        year, month, day = (int(g) for g in match.groups())
        return f"{year:04d}-{month:02d}-{day:02d}"
    return text


def is_blank(cell) -> bool:
    return stringify(cell.Value) == ""


def is_file_locked(path: Path) -> bool:
    """다른 프로그램이 이 파일을 쓰기 모드로 잡고 있어 우리가 쓰기로 열 수 없는지 확인.

    파일 하나를 쓰기로 여는 순간 Excel 이 "사용 중인 파일" 대화상자를 띄울 수
    있는데, 그건 우리 감시 스레드가 기다리는 InputBox 가 아니라서 그대로
    멈춰버린다. 미리 확인해 명확한 오류로 끝내는 편이 낫다.

    UNC 경로(네트워크 드라이브)에서는 rename 트릭이 안 먹을 수 있어 open() 방식을 쓴다.
    """
    if not path.exists():
        return False
    path_str = str(path)
    if path_str.startswith(("\\\\", "//")):
        try:
            with open(path, "r+b"):
                return False
        except OSError:
            return True
    try:
        path.rename(path)
        return False
    except OSError:
        return True


def find_name(book, target: str):
    """워크북/시트 범위를 통틀어 이름을 찾는다 (대소문자 무시)."""
    wanted = target.casefold()
    for item in book.Names:
        try:
            local = str(item.Name).split("!")[-1]
        except com_error:
            continue
        if local.casefold() == wanted:
            return item
    return None


def name_anchor(book, target: str):
    """이름이 가리키는 범위의 좌측 상단 셀."""
    item = find_name(book, target)
    if item is None:
        raise StepError(f"이름 관리자에서 '{target}' 을(를) 찾지 못했습니다")
    try:
        return item.RefersToRange.Cells(1, 1)
    except com_error as exc:
        raise StepError(
            f"이름 '{target}' 이 셀 범위를 가리키지 않습니다 "
            f"(RefersTo={_safe(item)}, {bot.describe_com_error(exc)})"
        ) from exc


def _safe(item) -> str:
    try:
        return str(item.RefersTo)
    except com_error:
        return "?"


def address(cell) -> str:
    """셀 주소를 A1 형식으로.

    pywin32 지연 바인딩에서 Range.Address 는 메서드가 아니라 이미 평가된
    문자열('$B$4')로 돌아온다. 거기에 인자를 붙여 호출하면
    TypeError: 'str' object is not callable 이 난다.
    조기 바인딩(gencache)에서는 반대로 메서드로 잡히므로 둘 다 받아준다.
    """
    try:
        raw = cell.Address
        if callable(raw):
            raw = raw(False, False)
        return str(raw).replace("$", "")
    except Exception:
        return "?"


def offset(cell, down: int = 0, right: int = 0):
    """cell 에서 아래로 down, 오른쪽으로 right 만큼 이동한 셀.

    Range.Offset 도 Address 와 같은 부류라 바인딩 방식에 따라 메서드가 아닐
    수 있는데, 그 경우 예외 없이 엉뚱한 셀을 돌려준다. 그래서 시트 좌표로
    직접 계산해 어떤 환경에서도 같은 결과가 나오게 한다.
    """
    return cell.Worksheet.Cells(cell.Row + down, cell.Column + right)


# ── 파이프라인 ──────────────────────────────────────
class WorkbookPipeline:
    """엑셀 파일 한 개를 11단계로 처리한다."""

    def __init__(self, xl, path: Path, cfg: PipelineConfig, inputs: FundInputs,
                 runners: dict, dialog: DialogOptions, pid: int, log: Logger,
                 stop_after: int = 0):
        self.xl = xl
        self.path = path
        self.cfg = cfg
        self.inputs = inputs
        self.runners = runners
        self.dialog = dialog
        self.pid = pid
        self.log = log
        self.stop_after = stop_after or 11

        self.book = None
        self.sheet = None
        self.fund_id = ""
        self.book_date = ""
        self.items: List[object] = []
        self.result = FileResult(path=path)

    # ── 실행 ─────────────────────────────────────
    def steps(self) -> List[Tuple[str, Callable]]:
        return [
            (f"{self.cfg.anchor_name} 에서 펀드아이디 읽기", self.step_fund_id),
            ("장부기준일 읽기", self.step_book_date),
            ("종목아이디 수집", self.step_items),
            ("시트 내용·이름 전체 삭제", self.step_clear),
            (f"{self.cfg.alt_sheet} 시트 삭제", self.step_delete_alt),
            (f"'{self.cfg.macro_overview}' 매크로", self.step_overview),
            ("E5:E9 기본 정보 입력", self.step_header),
            (f"'{self.cfg.macro_nav}' 매크로", self.step_nav),
            (f"'{self.cfg.macro_inv}' 매크로", self.step_inv),
            ("종목아이디 붙여넣기", self.step_paste_items),
            (f"'{self.cfg.macro_interest}' 매크로", self.step_interest),
        ]

    def run(self) -> FileResult:
        try:
            self._open()
        except StepError as exc:
            self.result.steps.append(StepResult(0, "파일 열기", "failed", str(exc)))
            return self.result
        except com_error as exc:
            self.result.steps.append(StepResult(
                0, "파일 열기", "failed", bot.describe_com_error(exc)))
            return self.result

        try:
            self._run_steps()
            self._finish()
        finally:
            self._close()
        return self.result

    def _run_steps(self) -> None:
        stopped = False
        for number, (name, func) in enumerate(self.steps(), start=1):
            if stopped:
                self.result.steps.append(StepResult(number, name, "skipped", "이전 단계 실패로 중단"))
                continue
            if number > self.stop_after:
                self.result.steps.append(StepResult(number, name, "skipped", "실행 범위 밖"))
                continue

            started = time.monotonic()
            try:
                detail = func() or ""
                elapsed = time.monotonic() - started
                self.result.steps.append(StepResult(number, name, "ok", detail, elapsed))
                self.log.debug(f"✓ {number}. {name}" + (f" — {detail}" if detail else ""), indent=1)
            except StepSkip as exc:
                elapsed = time.monotonic() - started
                self.result.steps.append(StepResult(number, name, "skipped", str(exc), elapsed))
                self.log(f"- {number}. {name} — {exc}", indent=1)
            except StepError as exc:
                elapsed = time.monotonic() - started
                self.result.steps.append(StepResult(number, name, "failed", str(exc), elapsed))
                self.log(f"❌ {number}. {name} — {exc}", indent=1)
                stopped = True
            except com_error as exc:
                elapsed = time.monotonic() - started
                detail = f"Excel 오류: {bot.describe_com_error(exc)}"
                self.result.steps.append(StepResult(number, name, "failed", detail, elapsed))
                self.log(f"❌ {number}. {name} — {detail}", indent=1)
                stopped = True
            except Exception as exc:                       # noqa: BLE001 - 단계 격리
                elapsed = time.monotonic() - started
                detail = f"{type(exc).__name__}: {exc}"
                self.result.steps.append(StepResult(number, name, "failed", detail, elapsed))
                self.log(f"❌ {number}. {name} — {detail}", indent=1)
                stopped = True

    def _finish(self) -> None:
        self.result.fund_id = self.fund_id
        self.result.book_date = self.book_date
        self.result.item_count = len(self.items)

        failed = any(s.status == "failed" for s in self.result.steps)
        partial = self.stop_after < 11

        if failed:
            self.result.status = "failed"
        elif partial:
            self.result.status = "skipped"          # 확인용 실행 — 저장하지 않음
        else:
            self.book.Save()
            self.result.saved = True
            self.result.status = "ok"

    # ── 파일 열고 닫기 ───────────────────────────
    def _open(self) -> None:
        # 저장할 예정이 없는 실행(미리보기·N단계까지 테스트)은 쓰기 권한이
        # 필요 없다. 읽기 전용으로 열면 그 파일이 다른 곳에서 이미 열려
        # 있어도(예: 사용자가 확인차 켜 둔 경우) 충돌 없이 열 수 있다.
        read_only = self.stop_after < 11

        if not read_only and is_file_locked(self.path):
            raise StepError(
                f"파일이 다른 프로그램에서 사용 중이라 쓰기 모드로 열 수 없습니다: {self.path}\n"
                "파일을 닫고 다시 실행하거나, [읽기만]/부분 실행으로 먼저 확인해 보세요."
            )

        self.xl.DisplayAlerts = False
        self.book = self.xl.Workbooks.Open(str(self.path), UpdateLinks=0, ReadOnly=read_only)
        self.xl.DisplayAlerts = True
        self.book.Activate()

    def _close(self) -> None:
        if self.book is None:
            return
        try:
            self.xl.DisplayAlerts = False
            self.book.Close(SaveChanges=False)      # 성공 시엔 _finish 에서 이미 저장
        except com_error:
            pass
        finally:
            self.xl.DisplayAlerts = True

    # ── 1~3 단계: 읽기 전용 ──────────────────────
    def step_fund_id(self) -> str:
        anchor = name_anchor(self.book, self.cfg.anchor_name)
        self.sheet = anchor.Worksheet
        self.sheet.Activate()

        down, right = self.cfg.fund_offset
        cell = offset(anchor, down, right)
        value = cell_text(cell)
        where = f"{self.sheet.Name}!{address(cell)}"

        if not value:
            raise StepError(
                f"펀드아이디 칸이 비어 있습니다 "
                f"({self.cfg.anchor_name}={address(anchor)} 기준 아래{down}·오른쪽{right} → {where})"
            )
        prefix = self.cfg.fund_prefix
        if prefix and not value.upper().startswith(prefix.upper()):
            raise StepError(
                f"펀드아이디가 '{prefix}' 로 시작하지 않습니다 (읽은 값 {value!r} @ {where}). "
                f"F_1 기준 위치가 맞는지 확인하세요."
            )
        self.fund_id = value
        return f"{value} @ {where}"

    def step_book_date(self) -> str:
        target = f"{self.fund_id}{self.cfg.nav_suffix}"
        anchor = name_anchor(self.book, target)
        down, right = self.cfg.nav_offset
        cell = offset(anchor, down, right)
        value = as_date_text(cell)
        where = f"{anchor.Worksheet.Name}!{address(cell)}"

        if not value:
            raise StepError(f"장부기준일이 비어 있습니다 ({target} 기준 → {where})")
        self.book_date = value
        return f"{value} @ {where}"

    def step_items(self) -> str:
        target = f"{self.fund_id}{self.cfg.inv_suffix}"
        anchor = name_anchor(self.book, target)
        sheet = anchor.Worksheet
        first = offset(anchor, self.cfg.inv_first_row_offset, 0)

        where = f"{target}={address(anchor)} 기준 아래{self.cfg.inv_first_row_offset} " \
                f"→ {sheet.Name}!{address(first)}"

        # 종목이 하나도 없는 펀드도 있다. 실패가 아니라 0개로 보고,
        # 붙여넣기(10)와 미수이자(11)만 건너뛴다.
        if is_blank(first):
            self.items = []
            self.log(f"⚠ 종목아이디 0개 — 10·11단계는 건너뜁니다 ({where})", indent=1)
            return f"0개 ({where} 비어 있음)"

        # Ctrl+↓ 와 같은 동작. 바로 아래가 비어 있으면 항목이 하나뿐이다.
        if is_blank(offset(first, 1, 0)):
            last_row = first.Row
        else:
            last_row = first.End(XL_DOWN).Row
            if last_row > MAX_DATA_ROW:
                raise StepError(
                    f"종목 목록의 끝을 찾지 못했습니다 (Ctrl+↓ 가 {last_row}행까지 내려감). "
                    f"목록 아래에 빈 칸이 있는지 확인하세요."
                )

        column = first.Column
        rng = sheet.Range(sheet.Cells(first.Row, column), sheet.Cells(last_row, column))
        self.items = _flatten(rng.Value)
        if not self.items:
            raise StepError(f"종목아이디를 하나도 읽지 못했습니다 ({sheet.Name}!{address(rng)})")

        preview = ", ".join(stringify(v) for v in self.items[:3])
        more = f" 외 {len(self.items) - 3}개" if len(self.items) > 3 else ""
        return f"{len(self.items)}개 ({preview}{more}) @ {sheet.Name}!{address(rng)}"

    # ── 4~5 단계: 정리 ───────────────────────────
    def step_clear(self) -> str:
        if self.cfg.clear_mode == "clear":
            self.sheet.Cells.Clear()
            how = "내용·서식 지움"
        else:
            # Ctrl+A → 마우스 오른쪽 → 삭제 와 같은 동작.
            # Clear 와 달리 셀 자체를 없애므로 병합·조건부서식·유효성검사까지 사라진다.
            self.sheet.Cells.Delete(XL_SHIFT_UP)
            how = "셀 전체 삭제"

        # F_1 이 있던 시트 전체를 흰색으로 채운다. Delete/Clear 만으로는 배경이
        # '채우기 없음'(테마에 따라 회색조로 보일 수 있음)이 되므로, 격자선까지
        # 덮이는 확실한 흰 배경을 원하면 이 칠하기가 필요하다.
        self.sheet.Cells.Interior.Color = COLOR_WHITE
        how += " · 시트 전체 흰색으로 채움"

        deleted, kept = 0, []
        for item in list(self.book.Names):
            try:
                label = str(item.Name)
            except com_error:
                label = "?"
            try:
                item.Delete()
                deleted += 1
            except com_error:
                kept.append(label)      # Print_Area 같은 내장 이름은 지워지지 않을 수 있다

        detail = f"시트 '{self.sheet.Name}' {how} · 이름 {deleted}개 삭제"
        if kept:
            detail += f" (삭제 안 됨 {len(kept)}개: {', '.join(kept[:3])})"
        return detail

    def step_delete_alt(self) -> str:
        target = self.cfg.alt_sheet.casefold()
        for sheet in self.book.Worksheets:
            if str(sheet.Name).casefold() != target:
                continue
            if self.book.Worksheets.Count <= 1:
                raise StepError(f"'{sheet.Name}' 이 마지막 시트라 삭제할 수 없습니다")
            self.xl.DisplayAlerts = False
            try:
                sheet.Delete()
            finally:
                self.xl.DisplayAlerts = True
            return f"'{self.cfg.alt_sheet}' 삭제"
        raise StepSkip(f"'{self.cfg.alt_sheet}' 시트가 없습니다")

    # ── 6~11 단계: 생성 ──────────────────────────
    def step_overview(self) -> str:
        return self._run_macro(self.cfg.macro_overview, self.cfg.overview_cell)

    def step_header(self) -> str:
        values = [
            ("평가기준일", self.inputs.eval_date),
            ("변수일", self.inputs.var_date),
            ("장부기준일", self.book_date),
            ("평가자", self.inputs.evaluator),
            ("리뷰어", self.inputs.reviewer),
        ]
        start = self.sheet.Range(self.cfg.header_cell)
        for index, (label, value) in enumerate(values):
            cell = offset(start, index, 0)
            if not value:
                raise StepError(f"{label} 값이 비어 있습니다 ({address(cell)})")
            cell.Value = value
        first, last = address(start), address(offset(start, 4, 0))
        return f"{first}:{last} = " + " / ".join(v for _, v in values)

    def step_nav(self) -> str:
        return self._run_macro(self.cfg.macro_nav, self.cfg.nav_cell)

    def step_inv(self) -> str:
        return self._run_macro(self.cfg.macro_inv, self.cfg.inv_cell)

    def step_paste_items(self) -> str:
        if not self.items:
            raise StepSkip("종목아이디가 0개라 붙여넣을 것이 없습니다")

        start = self.sheet.Range(self.cfg.items_cell)
        count = len(self.items)
        dest = self.sheet.Range(start, offset(start, count - 1, 0))

        # '005930' 처럼 앞자리 0이 있는 아이디가 숫자로 바뀌지 않게
        if any(isinstance(v, str) for v in self.items):
            dest.NumberFormat = "@"
        dest.Value = tuple((v,) for v in self.items)
        return f"{count}개 → {address(dest)}"

    def step_interest(self) -> str:
        if not self.items:
            raise StepSkip(
                "종목아이디가 0개라 실행 위치(마지막 종목 아래 "
                f"{self.cfg.interest_row_offset}칸)를 정할 수 없습니다"
            )

        start = self.sheet.Range(self.cfg.items_cell)
        last = offset(start, len(self.items) - 1, 0)
        target = offset(last, self.cfg.interest_row_offset, 0)
        return self._run_macro(self.cfg.macro_interest, address(target))

    # ── 매크로 실행 + 대화상자 처리 ──────────────
    def _run_macro(self, macro: str, cell_address: str) -> str:
        runner = self.runners[macro]
        cell = self.sheet.Range(cell_address)
        self.sheet.Activate()
        cell.Select()

        opts = DialogOptions(**{**self.dialog.__dict__, "value": self.fund_id})
        watcher = DialogHandler(self.pid, opts, self.log)
        watcher.start()
        try:
            resolved = runner.run(self.xl)
        finally:
            watcher.stop()
            watcher.join(timeout=5)

        if watcher.error:
            raise StepError(f"{cell_address} 선택 후 실행 — {watcher.error}")
        if not watcher.filled:
            raise StepError(
                f"{cell_address} 선택 후 실행했지만 펀드아이디 입력창을 찾지 못했습니다"
            )
        return f"{cell_address} → {resolved} (펀드아이디 {self.fund_id})"


def _flatten(raw) -> List[object]:
    """Range.Value 결과를 1차원 목록으로. 빈 칸은 버린다."""
    if raw is None:
        return []
    if not isinstance(raw, tuple):
        return [raw] if stringify(raw) else []
    values = []
    for row in raw:
        cell = row[0] if isinstance(row, tuple) else row
        if stringify(cell):
            values.append(cell)
    return values


# ── 여러 파일 실행 ──────────────────────────────────
def run_pipeline(root: Path, files: List[Path], cfg: PipelineConfig, inputs: FundInputs,
                 dialog: DialogOptions, log: Logger, personal: Optional[Path] = None,
                 stop_after: int = 0, keep_open: bool = False, cancel=None,
                 on_file=None, on_result=None) -> List[FileResult]:
    """대상 파일을 순서대로 처리한다.

    stop_after : 이 단계까지만 실행하고 저장하지 않는다 (0 = 전체 실행 후 저장)
    """
    if not files:
        log("대상 파일이 없습니다")
        return []

    if stop_after == 0 or stop_after > READONLY_STEPS:
        blanks = inputs.missing()
        if blanks:
            raise StepError(f"실행 정보가 비어 있습니다: {', '.join(blanks)}")

    total = len(files)
    mode = "전체 실행" if not stop_after else f"{stop_after}단계까지만 (저장 안 함)"
    log(f"대상 파일 {total}개 · {mode}")
    log("")

    pythoncom.CoInitialize()
    xl = None
    results: List[FileResult] = []
    try:
        xl = bot.start_excel()
        bot.open_macro_host(xl, personal, log)
        runners = {name: MacroRunner(name) for name in cfg.macro_names()}
        pid = bot.excel_pid(xl)

        for index, path in enumerate(files, start=1):
            if cancel is not None and cancel.is_set():
                log(f"\n사용자 요청으로 중단했습니다 ({index - 1}/{total} 처리됨)")
                break

            log(f"[{index}/{total}] {path.name}")
            if on_file is not None:
                on_file(index, total, path)

            pipeline = WorkbookPipeline(xl, path, cfg, inputs, runners, dialog,
                                        pid, log, stop_after)
            result = pipeline.run()
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
