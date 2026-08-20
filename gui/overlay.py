"""처리 중 화면을 흰색으로 덮는 오버레이

Excel COM 자동화가 셀을 선택하고 매크로 대화상자를 열고 닫는 동안 화면에
그대로 노출된다. 다른 사람이 볼 수 있는 화면일 수도 있고, 사용자가 실수로
그 위를 클릭하면 자동화(Win32 창 탐색)를 방해할 수도 있다. 그래서 작업
중에는 모든 모니터를 흰 화면으로 덮는다.

진행 상황과 [중지] 버튼은 오버레이 위에도 그대로 둔다 — 화면을 가렸다고
사용자가 진행 상황을 못 보거나 멈추지 못하면 안 되기 때문이다. Esc 를
누르면 가리기만 풀리고(자동화는 계속 진행) 급하게 화면을 봐야 할 때 쓸 수
있다.
"""
from __future__ import annotations

from typing import Callable, List

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QKeyEvent
from PyQt6.QtWidgets import QApplication, QLabel, QPushButton, QVBoxLayout, QWidget


class _ScreenPanel(QWidget):
    """모니터 한 대를 덮는 흰 화면 한 장."""

    def __init__(self, geometry, on_escape: Callable[[], None]):
        super().__init__()
        self._on_escape = on_escape
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
        )
        self.setStyleSheet("background-color: #ffffff;")
        self.setGeometry(geometry)

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.key() == Qt.Key.Key_Escape:
            self._on_escape()
            return
        super().keyPressEvent(event)

    def mousePressEvent(self, event) -> None:
        # 클릭이 뒤쪽 창(Excel 등)으로 새지 않도록 여기서 흡수한다.
        event.accept()


class ProcessingOverlay:
    """모니터마다 하나씩, 화면 전체를 덮는 흰 오버레이 묶음."""

    def __init__(self, stop_callback: Callable[[], None]):
        self._stop_callback = stop_callback
        self._panels: List[_ScreenPanel] = []
        self._status_labels: List[QLabel] = []

    @property
    def is_visible(self) -> bool:
        return bool(self._panels)

    def show_all(self, status: str = "처리를 준비하고 있습니다…") -> None:
        self.hide_all()
        screens = QApplication.screens()
        if not screens:
            return
        for screen in screens:
            panel = _ScreenPanel(screen.geometry(), self.hide_all)
            self._status_labels.append(self._build_content(panel, status))
            panel.show()
            self._panels.append(panel)
        self._panels[0].setFocus()
        self._panels[0].activateWindow()

    def _build_content(self, panel: _ScreenPanel, status: str) -> QLabel:
        layout = QVBoxLayout(panel)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(14)

        title = QLabel("처리 중입니다")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 22px; font-weight: 700; color: #1d1d1f;")
        layout.addWidget(title)

        status_label = QLabel(status)
        status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        status_label.setWordWrap(True)
        status_label.setStyleSheet("font-size: 14px; color: #555555;")
        status_label.setFixedWidth(480)
        layout.addWidget(status_label)

        stop_btn = QPushButton("중지")
        stop_btn.setFixedWidth(120)
        stop_btn.setStyleSheet(
            "padding: 8px; border-radius: 6px; font-weight: 600;"
            "background-color: #ff3b30; color: white; border: none;"
        )
        stop_btn.clicked.connect(self._stop_callback)
        layout.addWidget(stop_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        hint = QLabel("Esc — 화면 가리기만 해제합니다 (작업은 계속 진행됩니다)")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint.setStyleSheet("font-size: 11px; color: #999999;")
        layout.addWidget(hint)

        return status_label

    def update_status(self, text: str) -> None:
        for label in self._status_labels:
            label.setText(text)

    def hide_all(self) -> None:
        for panel in self._panels:
            panel.close()
            panel.deleteLater()
        self._panels = []
        self._status_labels = []
