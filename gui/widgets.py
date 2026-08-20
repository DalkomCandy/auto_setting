"""재사용 위젯 — 콘솔, 진행률, 폴더 입력"""
from __future__ import annotations

import json
import os
import time
from pathlib import Path
from typing import Optional

from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from PyQt6.QtWidgets import (
    QComboBox, QFileDialog, QHBoxLayout, QLabel, QProgressBar, QPushButton,
    QSizePolicy, QTextEdit, QVBoxLayout, QWidget,
)

ORG_NAME = "ExcelMacroBot"
APP_NAME = "ExcelMacroBot"

LEVEL_COLORS = {
    "info": "#1d1d1f",
    "success": "#2e7d32",
    "warning": "#ff8f00",
    "error": "#c62828",
}

STATUS_COLORS = {
    "ok": "#2e7d32",
    "skipped": "#86868b",
    "failed": "#c62828",
    "running": "#0071e3",
    "pending": "#b0b0b5",
}

STATUS_LABELS = {
    "ok": "완료",
    "skipped": "건너뜀",
    "failed": "실패",
    "running": "처리 중",
    "pending": "대기",
}

_HISTORY_MAX = 8


def _config_dir() -> Path:
    """설정 파일이 들어갈 사용자별 폴더. 저장소 밖이라 git pull 과 무관하다."""
    appdata = os.environ.get("APPDATA")
    base = Path(appdata) if appdata else Path.home() / ".config"
    return base / ORG_NAME


def config_path() -> Path:
    """설정 JSON 파일의 전체 경로. GUI 에서 '이 위치 열기' 버튼에도 쓴다."""
    return _config_dir() / "config.json"


class JsonSettings:
    """QSettings 를 대신하는 JSON 파일 저장소.

    Windows 레지스트리는 사람이 열어보거나 백업·공유하기 번거로워서,
    같은 인터페이스(value/setValue/contains/clear/sync)를 유지한 채
    사람이 읽고 고칠 수 있는 JSON 파일로 바꿨다. 호출부(window.py 의
    _load_settings/_save_settings, 이 파일의 FolderRow)는 수정할 필요가 없다.
    """

    def __init__(self):
        self.path = config_path()

    def _read(self) -> dict:
        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            return {}

    def _write(self, data: dict) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True),
            encoding="utf-8",
        )

    def contains(self, key: str) -> bool:
        return key in self._read()

    def value(self, key: str, default=None, type_=None):
        data = self._read()
        if key not in data:
            return default
        raw = data[key]
        if type_ is None:
            return raw
        if type_ is bool:
            return bool(raw)
        if type_ is int:
            return int(raw)
        if type_ is float:
            return float(raw)
        if type_ is str:
            return str(raw)
        return raw

    def setValue(self, key: str, value) -> None:
        data = self._read()
        data[key] = value
        self._write(data)

    def clear(self) -> None:
        self._write({})

    def sync(self) -> None:
        pass  # setValue 마다 이미 파일에 쓰므로 할 일 없음


def settings() -> JsonSettings:
    return JsonSettings()


class ConsoleView(QTextEdit):
    """타임스탬프가 붙는 읽기 전용 로그 뷰"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

    def append_message(self, message: str, level: str = "info") -> None:
        if not message.strip():
            self.append("")
            return
        color = LEVEL_COLORS.get(level, LEVEL_COLORS["info"])
        stamp = time.strftime("[%H:%M:%S]")
        # 들여쓰기(선행 공백)를 HTML 에서도 유지
        indent = len(message) - len(message.lstrip())
        body = "&nbsp;" * indent + _escape(message.strip())
        self.append(
            f'<span style="color:#b0b0b5">{stamp}</span> '
            f'<span style="color:{color}">{body}</span>'
        )
        bar = self.verticalScrollBar()
        bar.setValue(bar.maximum())


def _escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


class ProgressPanel(QWidget):
    """진행률 바 + 상태 / 속도·경과·예상 시간"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._start: Optional[float] = None
        self._build()

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)

        row = QHBoxLayout()
        row.setSpacing(10)
        self.bar = QProgressBar()
        self.bar.setTextVisible(False)
        self.percent = QLabel("0%")
        self.percent.setProperty("hint", True)
        self.percent.setMinimumWidth(38)
        self.percent.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        row.addWidget(self.bar, 1)
        row.addWidget(self.percent)
        layout.addLayout(row)

        self.status = QLabel("대기 중")
        self.status.setProperty("hint", True)
        layout.addWidget(self.status)

        self.detail = QLabel("")
        self.detail.setProperty("hint", True)
        layout.addWidget(self.detail)

    def start(self, total: int) -> None:
        self._start = time.time()
        self.bar.setMaximum(max(total, 1))
        self.bar.setValue(0)
        self.percent.setText("0%")
        self.status.setText(f"시작 — 총 {total}개")
        self.detail.setText("")

    def advance(self, done: int, total: int, current: str = "") -> None:
        self.bar.setMaximum(max(total, 1))
        self.bar.setValue(done)
        self.percent.setText(f"{int(done / total * 100)}%" if total else "0%")
        self.status.setText(f"{done} / {total}   {current}".strip())

        if self._start and done > 0:
            elapsed = time.time() - self._start
            speed = done / elapsed if elapsed else 0
            eta = (total - done) / speed if speed else 0
            self.detail.setText(
                f"경과 {self._fmt(elapsed)} · 남은 시간 약 {self._fmt(eta)} "
                f"· 파일당 {elapsed / done:.1f}초"
            )

    def finish(self, text: str) -> None:
        self.bar.setValue(self.bar.maximum())
        self.percent.setText("100%")
        self.status.setText(text)

    def reset(self) -> None:
        self._start = None
        self.bar.setValue(0)
        self.percent.setText("0%")
        self.status.setText("대기 중")
        self.detail.setText("")

    @staticmethod
    def _fmt(seconds: float) -> str:
        if seconds < 60:
            return f"{int(seconds)}초"
        if seconds < 3600:
            return f"{int(seconds // 60)}분 {int(seconds % 60)}초"
        return f"{int(seconds // 3600)}시간 {int((seconds % 3600) // 60)}분"


class FolderRow(QWidget):
    """폴더 경로 입력 — 히스토리 드롭다운 + 찾아보기 + 드래그앤드롭"""

    changed = pyqtSignal(str)

    def __init__(self, setting_key: str, placeholder: str = "", parent=None):
        super().__init__(parent)
        self.setting_key = setting_key
        self.setAcceptDrops(True)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.combo = QComboBox()
        self.combo.setEditable(True)
        self.combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.combo.lineEdit().setPlaceholderText(placeholder)
        self.combo.currentTextChanged.connect(self.changed.emit)

        self.browse = QPushButton("찾아보기")
        self.browse.setProperty("secondary", True)
        self.browse.setFixedWidth(74)
        self.browse.clicked.connect(self._browse)

        layout.addWidget(self.combo, 1)
        layout.addWidget(self.browse)
        self._load_history()

    def path(self) -> str:
        return self.combo.currentText().strip()

    def set_path(self, value: str) -> None:
        self.combo.setCurrentText(value)

    def _browse(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "대상 폴더 선택", self.path())
        if folder:
            self.combo.setCurrentText(str(Path(folder)))
            self.remember()

    def remember(self) -> None:
        """현재 경로를 히스토리 맨 앞에 저장한다."""
        value = self.path()
        if not value:
            return
        store = settings()
        history = store.value(f"{self.setting_key}/history", []) or []
        if isinstance(history, str):
            history = [history]
        history = [value] + [p for p in history if p != value]
        store.setValue(f"{self.setting_key}/history", history[:_HISTORY_MAX])
        self._load_history(keep=value)

    def _load_history(self, keep: str = "") -> None:
        store = settings()
        history = store.value(f"{self.setting_key}/history", []) or []
        if isinstance(history, str):
            history = [history]
        current = keep or self.path()
        self.combo.blockSignals(True)
        self.combo.clear()
        self.combo.addItems(history)
        self.combo.setCurrentText(current)
        self.combo.blockSignals(False)

    # ── 드래그앤드롭 ─────────────────────────────
    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:
        urls = event.mimeData().urls()
        if not urls:
            return
        dropped = Path(urls[0].toLocalFile())
        folder = dropped if dropped.is_dir() else dropped.parent
        self.combo.setCurrentText(str(folder))
        self.remember()
