"""엑셀 매크로 봇 — GUI 진입점

    python main.py

CLI 로 쓰려면 excel_macro_bot.py 를 직접 실행하세요.
"""
from __future__ import annotations

import sys
from pathlib import Path

from PyQt6.QtWidgets import QApplication, QMessageBox

BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).parent)).resolve()
STYLE_PATH = BASE_DIR / "gui" / "style.qss"

# PyInstaller 로 묶었을 때도 gui / excel_macro_bot 을 찾을 수 있게
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("엑셀 매크로 봇")

    if STYLE_PATH.exists():
        app.setStyleSheet(STYLE_PATH.read_text(encoding="utf-8"))

    from gui import MainWindow

    window = MainWindow()
    window.show()

    if not sys.platform.startswith("win"):
        QMessageBox.warning(
            window, "Windows 전용",
            "이 프로그램은 Excel COM 자동화를 사용하므로 Windows 에서만 동작합니다.\n"
            "화면은 열리지만 실행 기능은 사용할 수 없습니다.",
        )

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
