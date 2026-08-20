"""Win32 창/컨트롤 조작 헬퍼

VBA `InputBox` 가 띄우는 모달 대화상자는 표준 다이얼로그(`#32770`)이므로
win32 메시지만으로 텍스트 입력과 버튼 클릭이 가능하다.
COM(win32com)은 매크로 실행 중 블로킹되므로 이 모듈의 함수들은
전부 별도 스레드에서 호출되는 것을 전제로 한다.
"""
from __future__ import annotations

import ctypes
import sys
import time
from ctypes import wintypes
from dataclasses import dataclass
from typing import List, Optional

IS_WINDOWS = sys.platform.startswith("win")

if IS_WINDOWS:
    import win32gui
    import win32process
    import pywintypes

# ── Win32 상수 ──────────────────────────────────────
WM_GETTEXT       = 0x000D
WM_GETTEXTLENGTH = 0x000E
WM_SETTEXT       = 0x000C
WM_CLOSE         = 0x0010
WM_COMMAND       = 0x0111
WM_KEYDOWN       = 0x0100
WM_KEYUP         = 0x0101
WM_CHAR          = 0x0102
BM_CLICK         = 0x00F5
EM_SETSEL        = 0x00B1
VK_RETURN        = 0x0D
VK_ESCAPE        = 0x1B
IDOK             = 1
IDCANCEL         = 2
SMTO_ABORTIFHUNG = 0x0002

EXCEL_MAIN_CLASS = "XLMAIN"

#: 확인 버튼으로 인정할 캡션 (& 가속키와 공백은 제거 후 비교)
CONFIRM_CAPTIONS = {"확인", "ok", "예", "yes", "적용", "실행", "생성", "생성하기", "만들기"}
#: 취소 버튼으로 인정할 캡션
CANCEL_CAPTIONS = {"취소", "cancel", "아니오", "no", "닫기", "close"}


def _setup_user32():
    """64비트에서 포인터가 잘리지 않도록 프로토타입을 명시한다."""
    u = ctypes.WinDLL("user32", use_last_error=True)
    LRESULT = ctypes.c_ssize_t

    u.SendMessageW.restype = LRESULT
    u.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, ctypes.c_size_t, LRESULT]

    u.PostMessageW.restype = wintypes.BOOL
    u.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, ctypes.c_size_t, LRESULT]

    u.SendMessageTimeoutW.restype = LRESULT
    u.SendMessageTimeoutW.argtypes = [
        wintypes.HWND, wintypes.UINT, ctypes.c_size_t, LRESULT,
        wintypes.UINT, wintypes.UINT, ctypes.POINTER(ctypes.c_size_t),
    ]

    u.IsWindow.restype = wintypes.BOOL
    u.IsWindow.argtypes = [wintypes.HWND]

    u.GetDlgCtrlID.restype = ctypes.c_int
    u.GetDlgCtrlID.argtypes = [wintypes.HWND]
    return u


user32 = _setup_user32() if IS_WINDOWS else None


# ── 저수준 래퍼 ─────────────────────────────────────
def send_timeout(hwnd: int, msg: int, wparam: int = 0, lparam: int = 0,
                 timeout_ms: int = 2000) -> Optional[int]:
    """응답 없는 창에서 멈추지 않는 SendMessage. 실패 시 None."""
    out = ctypes.c_size_t(0)
    ok = user32.SendMessageTimeoutW(
        hwnd, msg, wparam, lparam, SMTO_ABORTIFHUNG, timeout_ms, ctypes.byref(out)
    )
    return out.value if ok else None


def get_text(hwnd: int) -> str:
    """다른 프로세스의 컨트롤에서도 동작하는 캡션 조회."""
    length = send_timeout(hwnd, WM_GETTEXTLENGTH)
    if not length:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    send_timeout(hwnd, WM_GETTEXT, length + 1, ctypes.cast(buf, ctypes.c_void_p).value)
    return buf.value


def set_text(hwnd: int, text: str) -> None:
    buf = ctypes.create_unicode_buffer(text)
    user32.SendMessageW(hwnd, WM_SETTEXT, 0, ctypes.cast(buf, ctypes.c_void_p).value)


def type_text(hwnd: int, text: str, delay: float = 0.01) -> None:
    """WM_CHAR 를 한 글자씩 보내 실제 타이핑처럼 변경 알림까지 발생시킨다.

    WM_SETTEXT 를 무시하는 컨트롤(RefEdit 등)에 대한 대안.
    """
    set_text(hwnd, "")
    user32.SendMessageW(hwnd, EM_SETSEL, 0, -1)
    for ch in text:
        user32.PostMessageW(hwnd, WM_CHAR, ord(ch), 0)
        if delay:
            time.sleep(delay)


def press_key(hwnd: int, vk: int) -> None:
    user32.PostMessageW(hwnd, WM_KEYDOWN, vk, 0)
    user32.PostMessageW(hwnd, WM_KEYUP, vk, 0)


def is_alive(hwnd: int) -> bool:
    return bool(user32.IsWindow(hwnd)) and win32gui.IsWindowVisible(hwnd)


def wait_gone(hwnd: int, timeout: float, poll: float = 0.05) -> bool:
    """창이 닫힐 때까지 대기. 닫혔으면 True."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not is_alive(hwnd):
            return True
        time.sleep(poll)
    return not is_alive(hwnd)


# ── 창 탐색 ─────────────────────────────────────────
@dataclass
class Control:
    hwnd: int
    cls: str
    text: str
    ctrl_id: int
    rect: tuple

    @property
    def is_edit(self) -> bool:
        # Edit, RichEdit20W, RefEdit20 등을 모두 포괄
        return "edit" in self.cls.lower()

    @property
    def is_ref_edit(self) -> bool:
        """셀 범위를 고르는 RefEdit 칸인지.

        클래스명이 보통 'RefEdit20WndClass' 처럼 'edit' 를 포함하고 있어
        is_edit 로도 True 가 나온다. 그래서 '비어 있는 칸을 고른다'는
        pick_edit() 의 판별만으로는, RefEdit(SELECT RANGE) 가 아직
        '$B$2' 로 채워지기 전 찰나를 잡으면 거기에 펀드아이디를 넣어버릴
        수 있다 — 문자열이 유효한 셀 범위가 아니니 "이 이름에 대한
        구문이 잘못되었습니다" 오류로 이어진다. 그래서 이름으로도 한 번
        더 걸러낸다.
        """
        return "refedit" in self.cls.lower()

    @property
    def is_button(self) -> bool:
        return self.cls == "Button"

    @property
    def caption(self) -> str:
        """가속키(&)와 공백을 제거한 비교용 캡션."""
        return self.text.replace("&", "").strip().casefold()


def _describe(hwnd: int) -> Control:
    try:
        cls = win32gui.GetClassName(hwnd)
    except pywintypes.error:
        cls = "?"
    try:
        rect = win32gui.GetWindowRect(hwnd)
    except pywintypes.error:
        rect = (0, 0, 0, 0)
    return Control(hwnd, cls, get_text(hwnd), user32.GetDlgCtrlID(hwnd), rect)


def list_controls(hwnd: int) -> List[Control]:
    """대화상자의 모든 하위 컨트롤(자손 포함)을 나열한다."""
    handles: List[int] = []

    def collect(child, _):
        handles.append(child)
        return True

    try:
        win32gui.EnumChildWindows(hwnd, collect, None)
    except pywintypes.error:
        pass
    return [_describe(h) for h in handles]


def find_dialogs(pid: int, title_contains: str = "") -> List[int]:
    """Excel 프로세스가 띄운 최상위 대화상자 핸들 목록.

    메인 창(XLMAIN)과 크기가 없는 유령 창은 제외한다.
    """
    found: List[int] = []

    def visit(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        try:
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
        except pywintypes.error:
            return True
        if wpid != pid:
            return True
        try:
            if win32gui.GetClassName(hwnd) == EXCEL_MAIN_CLASS:
                return True
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        except pywintypes.error:
            return True
        if right - left <= 0 or bottom - top <= 0:
            return True
        if title_contains and title_contains not in win32gui.GetWindowText(hwnd):
            return True
        found.append(hwnd)
        return True

    win32gui.EnumWindows(visit, None)
    return found


def window_title(hwnd: int) -> str:
    try:
        return win32gui.GetWindowText(hwnd)
    except pywintypes.error:
        return ""


# ── 대화상자 조작 ───────────────────────────────────
def pick_button(controls: List[Control], captions: set) -> Optional[Control]:
    for ctrl in controls:
        if ctrl.is_button and ctrl.caption in captions:
            return ctrl
    return None


def pick_edit(controls: List[Control], index: int = -1) -> Optional[Control]:
    """값을 넣을 입력칸을 고른다.

    '수익 개요' 대화상자처럼 입력칸이 둘(SELECT RANGE + 펀드아이디)일 때,
    SELECT RANGE 는 RefEdit(셀 범위 선택) 컨트롤이라 절대 건드리면 안 된다 —
    거기에 펀드아이디 같은 문자열을 넣으면 유효한 셀 범위가 아니라서
    "이 이름에 대한 구문이 잘못되었습니다" 오류로 이어진다. RefEdit 는
    클래스명에 'edit' 가 포함돼 있어 '비어 있는 칸' 판별만으로는 부족하다
    (특히 Excel 이 아직 $B$2 를 채우기 전 찰나를 잡으면 빈 칸으로 보인다).
    그래서 RefEdit 가 아닌 칸을 먼저 찾고, 그 안에서 '비어 있는 첫 칸'을 고른다.

    index >= 0 이면 자동 판별 대신 그 순번의 칸(RefEdit 포함 전체 목록 기준)을
    그대로 쓴다 — 자동 판별이 틀렸을 때 사용자가 직접 지정하는 용도다.
    """
    edits = [c for c in controls if c.is_edit]
    if not edits:
        return None
    if index >= 0:
        return edits[index] if index < len(edits) else None

    candidates = [c for c in edits if not c.is_ref_edit] or edits
    if len(candidates) > 1:
        empty = [c for c in candidates if not c.text.strip()]
        if empty:
            return empty[0]
    return candidates[0]


def confirm_dialog(hwnd: int, controls: List[Control], timeout: float) -> str:
    """확인 버튼 클릭 → WM_COMMAND(IDOK) → Enter 순으로 시도.

    Returns: 성공한 방법 이름. 실패 시 빈 문자열.
    """
    button = pick_button(controls, CONFIRM_CAPTIONS)
    if button is None:
        # 캡션을 못 읽는 경우 ID가 IDOK 인 버튼을 사용
        button = next((c for c in controls if c.is_button and c.ctrl_id == IDOK), None)

    if button is not None:
        user32.PostMessageW(button.hwnd, BM_CLICK, 0, 0)
        if wait_gone(hwnd, timeout):
            return f"버튼 클릭({button.text or 'IDOK'})"

    user32.PostMessageW(hwnd, WM_COMMAND, IDOK, 0)
    if wait_gone(hwnd, timeout):
        return "WM_COMMAND(IDOK)"

    press_key(hwnd, VK_RETURN)
    if wait_gone(hwnd, timeout):
        return "Enter"
    return ""


def cancel_dialog(hwnd: int, controls: List[Control], timeout: float = 5.0) -> bool:
    """취소/ESC/닫기 순으로 시도해 대화상자를 걷어낸다."""
    button = pick_button(controls, CANCEL_CAPTIONS)
    if button is not None:
        user32.PostMessageW(button.hwnd, BM_CLICK, 0, 0)
        if wait_gone(hwnd, timeout):
            return True

    user32.PostMessageW(hwnd, WM_COMMAND, IDCANCEL, 0)
    if wait_gone(hwnd, timeout):
        return True

    press_key(hwnd, VK_ESCAPE)
    if wait_gone(hwnd, timeout):
        return True

    user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
    return wait_gone(hwnd, timeout)


def format_dialog_report(hwnd: int, controls: List[Control]) -> str:
    """--probe 모드용 대화상자 구조 덤프."""
    lines = [
        f"  창 제목 : {window_title(hwnd)!r}",
        f"  클래스  : {win32gui.GetClassName(hwnd)}",
        f"  컨트롤  : {len(controls)}개",
    ]
    for ctrl in controls:
        if ctrl.is_ref_edit:
            tag = "REFEDIT"       # SELECT RANGE — 절대 값을 넣지 않는 칸
        elif ctrl.is_edit:
            tag = "EDIT   "
        elif ctrl.is_button:
            tag = "BUTTON "
        else:
            tag = "       "
        lines.append(
            f"    {tag} cls={ctrl.cls:<16} id={ctrl.ctrl_id:<6} text={ctrl.text!r}"
        )
    return "\n".join(lines)
