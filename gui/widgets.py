"""재사용 위젯 — 콘솔, 진행률, 폴더 입력"""
from __future__ import annotations

import time
from pathlib import Path
from typing import Optional

from PyQt6.QtCore import Qt, QSettings, pyqtSignal
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


def settings() -> QSettings:
    return QSettings(ORG_NAME, APP_NAME)


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
