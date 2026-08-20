"""메인 윈도우 — 11단계 수익 보고서 생성"""
from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QAbstractItemView, QCheckBox, QComboBox, QDoubleSpinBox, QFileDialog,
    QFormLayout, QGroupBox, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QMainWindow, QMessageBox, QPushButton, QScrollArea, QSpinBox, QSplitter,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

import excel_macro_bot as bot
import fund_pipeline as fp
from .widgets import (
    STATUS_COLORS, STATUS_LABELS, ConsoleView, FolderRow, ProgressPanel, settings,
)
from .worker import MacroListWorker, PipelineWorker, ScanWorker

COLUMNS = ["파일명", "폴더", "상태", "펀드아이디", "장부기준일", "종목수", "결과"]


def _relative(path: Path, root: Optional[Path]) -> str:
    """root 기준 상대경로. 관계가 없으면 전체 경로를 그대로 보여준다."""
    if root is None:
        return str(path)
    try:
        return str(path.relative_to(root)) or "."
    except ValueError:
        return str(path)


def _row(*widgets, stretch: int = 0) -> QWidget:
    """폼 한 줄에 위젯 여러 개를 나란히 놓는다."""
    box = QWidget()
    layout = QHBoxLayout(box)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(6)
    for index, widget in enumerate(widgets):
        layout.addWidget(widget, 1 if index == stretch else 0)
    return box


def _label(text: str) -> QLabel:
    tag = QLabel(text)
    tag.setProperty("hint", True)
    return tag


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("수익 보고서 생성 봇")
        self.resize(1280, 820)

        self.files: List[Path] = []
        self.scan_root: Optional[Path] = None
        self.rows: Dict[str, int] = {}
        self.row_paths: List[Path] = []
        self.results: List[fp.FileResult] = []
        self.by_path: Dict[str, fp.FileResult] = {}
        self.worker: Optional[PipelineWorker] = None
        self.helper: Optional[object] = None

        self._build_ui()
        self._load_settings()
        self._set_running(False)
        self.console.append_message(
            "폴더를 고르고 [파일 찾기] → [읽기만 (1~3단계)] 순으로 먼저 확인해 보세요."
        )
        self.console.append_message(
            "표에서 파일을 더블클릭하면 그 파일의 단계별 결과를 볼 수 있습니다."
        )

    # ── UI ───────────────────────────────────────────
    def _build_ui(self) -> None:
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self._build_left())
        splitter.addWidget(self._build_right())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([510, 770])

        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.addWidget(splitter)
        self.setCentralWidget(container)
        self.statusBar().showMessage("준비됨")

        for widget in (self.prefix, self.pattern):
            widget.textChanged.connect(self._invalidate_files)
        self.folder.changed.connect(self._invalidate_files)

    def _build_left(self) -> QWidget:
        panel = QWidget()
        panel.setObjectName("panel")
        outer = QVBoxLayout(panel)
        outer.setContentsMargins(16, 14, 16, 14)
        outer.setSpacing(2)

        title = QLabel("설정")
        title.setProperty("headline", True)
        outer.addWidget(title)

        outer.addWidget(self._group_target())
        outer.addWidget(self._group_inputs())
        outer.addWidget(self._group_macros())
        outer.addWidget(self._group_advanced())
        outer.addStretch(1)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(panel)
        scroll.setMinimumWidth(490)
        return scroll

    def _group_target(self) -> QGroupBox:
        group = QGroupBox("대상 폴더")
        form = QFormLayout(group)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setSpacing(8)

        self.folder = FolderRow("root", "예: D:\\all폴더")
        form.addRow("폴더", self.folder)

        self.prefix = QLineEdit("26.3Q")
        self.prefix.setObjectName("prefix")
        self.pattern = QLineEdit("*.xlsx")
        self.pattern.setObjectName("pattern")
        self.pattern.setFixedWidth(90)
        form.addRow("파일명 시작", _row(self.prefix, self.pattern))
        form.addRow("", _label("하위 폴더까지 모두 훑습니다 (all폴더 › 개별 폴더 › 파일)"))
        return group

    def _group_inputs(self) -> QGroupBox:
        group = QGroupBox("실행 정보 (모든 파일 공통)")
        form = QFormLayout(group)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setSpacing(8)

        self.eval_date = QLineEdit()
        self.eval_date.setObjectName("eval_date")
        self.eval_date.setPlaceholderText("예: 2026-03-31")
        form.addRow("평가기준일", self.eval_date)

        self.var_date = QLineEdit()
        self.var_date.setObjectName("var_date")
        self.var_date.setPlaceholderText("예: 2026-04-15")
        form.addRow("변수일", self.var_date)

        self.evaluator = QLineEdit()
        self.evaluator.setObjectName("evaluator")
        form.addRow("평가자", self.evaluator)

        self.reviewer = QLineEdit()
        self.reviewer.setObjectName("reviewer")
        form.addRow("리뷰어", self.reviewer)

        form.addRow("", _label("E5~E9 에 들어갑니다. 장부기준일(E7)은 파일에서 읽습니다."))
        return group

    def _group_macros(self) -> QGroupBox:
        group = QGroupBox("매크로 이름")
        form = QFormLayout(group)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setSpacing(8)

        defaults = fp.PipelineConfig()
        self.macro_overview = self._macro_box("macro_overview", defaults.macro_overview)
        self.macro_nav = self._macro_box("macro_nav", defaults.macro_nav)
        self.macro_inv = self._macro_box("macro_inv", defaults.macro_inv)
        self.macro_interest = self._macro_box("macro_interest", defaults.macro_interest)

        self.macro_btn = QPushButton("목록 불러오기")
        self.macro_btn.setProperty("secondary", True)
        self.macro_btn.setToolTip("Excel 을 잠깐 열어 사용 가능한 매크로 이름을 읽어옵니다.")
        self.macro_btn.clicked.connect(self._load_macro_names)

        form.addRow("6단계", self.macro_overview)
        form.addRow("8단계", self.macro_nav)
        form.addRow("9단계", self.macro_inv)
        form.addRow("11단계", self.macro_interest)
        form.addRow("", self.macro_btn)
        form.addRow("", _label(
            "이름을 자동으로 추측하지 않습니다 — 입력한 그대로 실행합니다.\n"
            "[목록 불러오기]로 정확한 이름을 확인해 그대로 넣거나(예: PERSONAL.XLSB!수익_개요),\n"
            "직접 확인했다면 손으로 입력하세요."
        ))
        return group

    def _macro_box(self, name: str, example: str) -> QComboBox:
        """빈 값으로 시작해 실제 이름을 직접 입력하도록 강제한다.

        예전에는 여기 실제 값처럼 보이는 기본값('수익 개요')을 채워뒀는데,
        VBA 이름에는 공백을 못 써서 그대로 실행하면 항상 실패했다.
        이제 예시는 플레이스홀더로만 보여주고 실제 값은 비워둔다.
        """
        box = QComboBox()
        box.setObjectName(name)
        box.setEditable(True)
        box.lineEdit().setPlaceholderText(f"예: PERSONAL.XLSB!{example.replace(' ', '_')}")
        return box

    def _group_advanced(self) -> QGroupBox:
        group = QGroupBox("고급")
        outer = QVBoxLayout(group)
        outer.setSpacing(8)

        self.advanced_toggle = QCheckBox("고급 옵션 표시")
        self.advanced_toggle.setObjectName("advanced_toggle")
        self.advanced_toggle.toggled.connect(lambda on: self.advanced_body.setVisible(on))
        outer.addWidget(self.advanced_toggle)

        self.advanced_body = QWidget()
        form = QFormLayout(self.advanced_body)
        form.setContentsMargins(0, 4, 0, 0)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        form.setSpacing(8)
        default = fp.PipelineConfig()

        # 실행 범위
        self.stop_after = QSpinBox()
        self.stop_after.setObjectName("stop_after")
        self.stop_after.setRange(0, 11)
        self.stop_after.setSpecialValueText("전체 (11단계, 저장)")
        self.stop_after.setToolTip("N 단계까지만 실행하고 저장하지 않습니다. 0 이면 전체 실행 후 저장.")
        form.addRow("실행 범위", self.stop_after)

        # 위치 — 읽기
        self.anchor_name = QLineEdit(default.anchor_name)
        self.anchor_name.setObjectName("anchor_name")
        self.fund_down = self._spin("fund_down", default.fund_offset[0], 0, 100)
        self.fund_right = self._spin("fund_right", default.fund_offset[1], 0, 100)
        form.addRow("펀드아이디 이름", self.anchor_name)
        form.addRow("↳ 위치", _row(_label("아래"), self.fund_down,
                                  _label("오른쪽"), self.fund_right, QWidget(), stretch=4))

        self.fund_prefix = QLineEdit(default.fund_prefix)
        self.fund_prefix.setObjectName("fund_prefix")
        self.fund_prefix.setToolTip("펀드아이디가 이 글자로 시작하지 않으면 위치 오류로 봅니다. 비우면 검사 안 함.")
        form.addRow("↳ 시작 글자", self.fund_prefix)

        self.nav_suffix = QLineEdit(default.nav_suffix)
        self.nav_suffix.setObjectName("nav_suffix")
        self.nav_down = self._spin("nav_down", default.nav_offset[0], 0, 100)
        self.nav_right = self._spin("nav_right", default.nav_offset[1], 0, 100)
        form.addRow("장부기준일 접미사", self.nav_suffix)
        form.addRow("↳ 위치", _row(_label("아래"), self.nav_down,
                                  _label("오른쪽"), self.nav_right, QWidget(), stretch=4))

        self.inv_suffix = QLineEdit(default.inv_suffix)
        self.inv_suffix.setObjectName("inv_suffix")
        self.inv_first = self._spin("inv_first", default.inv_first_row_offset, 0, 100)
        form.addRow("종목 목록 접미사", self.inv_suffix)
        form.addRow("↳ 첫 종목", _row(_label("아래"), self.inv_first, QWidget(), stretch=2))

        self.alt_sheet = QLineEdit(default.alt_sheet)
        self.alt_sheet.setObjectName("alt_sheet")
        form.addRow("삭제할 시트", self.alt_sheet)

        self.clear_mode = QComboBox()
        self.clear_mode.setObjectName("clear_mode")
        self.clear_mode.addItems(["셀 전체 삭제 (Ctrl+A → 삭제)", "내용·서식만 지우기"])
        self.clear_mode.setToolTip(
            "셀 전체 삭제는 병합·조건부서식·유효성검사까지 없앱니다. "
            "내용·서식만 지우기는 셀을 그대로 두고 값과 서식만 비웁니다."
        )
        form.addRow("4단계 방식", self.clear_mode)

        # 위치 — 쓰기
        self.overview_cell = self._cell("overview_cell", default.overview_cell)
        self.header_cell = self._cell("header_cell", default.header_cell)
        self.nav_cell = self._cell("nav_cell", default.nav_cell)
        self.inv_cell = self._cell("inv_cell", default.inv_cell)
        self.items_cell = self._cell("items_cell", default.items_cell)
        self.interest_down = self._spin("interest_down", default.interest_row_offset, 0, 100)
        form.addRow("셀 위치", _row(_label("개요"), self.overview_cell,
                                  _label("정보"), self.header_cell, QWidget(), stretch=4))
        form.addRow("", _row(_label("NAV"), self.nav_cell,
                             _label("투자자산"), self.inv_cell, QWidget(), stretch=4))
        form.addRow("", _row(_label("종목"), self.items_cell,
                             _label("미수이자 = 마지막 종목 아래"), self.interest_down,
                             QWidget(), stretch=4))

        # 대화상자
        self.input_method = QComboBox()
        self.input_method.setObjectName("input_method")
        self.input_method.addItems(["기본 (WM_SETTEXT)", "타이핑 (한 글자씩)"])
        form.addRow("입력 방식", self.input_method)

        self.field_index = QSpinBox()
        self.field_index.setObjectName("field_index")
        self.field_index.setRange(-1, 9)
        self.field_index.setValue(-1)
        self.field_index.setSpecialValueText("자동 (빈 칸)")
        self.field_index.setToolTip(
            "입력칸이 여러 개일 때 몇 번째에 펀드아이디를 넣을지. "
            "자동이면 비어 있는 첫 칸을 고릅니다 (SELECT RANGE 는 이미 채워져 있으므로)."
        )
        form.addRow("입력칸 선택", self.field_index)

        self.dialog_title = QLineEdit()
        self.dialog_title.setObjectName("dialog_title")
        self.dialog_title.setPlaceholderText("비우면 모든 창을 대상으로")
        form.addRow("창 제목 필터", self.dialog_title)

        self.appear = self._dspin("appear", 30.0, 1.0, 600.0)
        self.confirm = self._dspin("confirm", 10.0, 1.0, 120.0)
        form.addRow("입력창 대기", self.appear)
        form.addRow("확인 후 대기", self.confirm)

        self.dismiss = QCheckBox("매크로가 띄우는 후속 안내창 자동 확인")
        self.dismiss.setObjectName("dismiss")
        self.dismiss.setChecked(True)
        form.addRow("", self.dismiss)

        self.personal = QLineEdit()
        self.personal.setObjectName("personal")
        self.personal.setPlaceholderText("자동 (XLSTART\\PERSONAL.XLSB)")
        browse = QPushButton("찾기")
        browse.setProperty("secondary", True)
        browse.setFixedWidth(52)
        browse.clicked.connect(self._browse_personal)
        form.addRow("매크로 파일", _row(self.personal, browse))

        self.no_personal = QCheckBox("매크로 파일을 열지 않음")
        self.no_personal.setObjectName("no_personal")
        form.addRow("", self.no_personal)

        self.keep_open = QCheckBox("작업 후 Excel 을 닫지 않음 (디버깅용)")
        self.keep_open.setObjectName("keep_open")
        form.addRow("", self.keep_open)

        self.limit = QSpinBox()
        self.limit.setObjectName("limit")
        self.limit.setRange(0, 100000)
        self.limit.setSpecialValueText("전체")
        self.limit.valueChanged.connect(self._invalidate_files)
        form.addRow("처리 개수 제한", self.limit)

        outer.addWidget(self.advanced_body)
        self.advanced_body.setVisible(False)
        return group

    def _spin(self, name: str, value: int, low: int, high: int) -> QSpinBox:
        box = QSpinBox()
        box.setObjectName(name)
        box.setRange(low, high)
        box.setValue(value)
        box.setFixedWidth(58)
        return box

    def _dspin(self, name: str, value: float, low: float, high: float) -> QDoubleSpinBox:
        box = QDoubleSpinBox()
        box.setObjectName(name)
        box.setRange(low, high)
        box.setValue(value)
        box.setSuffix(" 초")
        return box

    def _cell(self, name: str, value: str) -> QLineEdit:
        box = QLineEdit(value)
        box.setObjectName(name)
        box.setFixedWidth(58)
        return box

    def _build_right(self) -> QWidget:
        panel = QWidget()
        panel.setObjectName("panel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        bar = QHBoxLayout()
        bar.setSpacing(8)
        self.scan_btn = QPushButton("파일 찾기")
        self.scan_btn.setProperty("secondary", True)
        self.scan_btn.clicked.connect(self._scan)

        self.preview_btn = QPushButton("읽기만 (1~3단계)")
        self.preview_btn.setProperty("secondary", True)
        self.preview_btn.setToolTip(
            "펀드아이디·장부기준일·종목아이디만 읽어봅니다. 파일을 전혀 수정하지 않습니다."
        )
        self.preview_btn.clicked.connect(lambda: self._start(preview=True))

        self.run_btn = QPushButton("실행")
        self.run_btn.setProperty("primary", True)
        self.run_btn.setMinimumWidth(96)
        self.run_btn.clicked.connect(lambda: self._start(preview=False))

        self.stop_btn = QPushButton("중지")
        self.stop_btn.setProperty("danger", True)
        self.stop_btn.clicked.connect(self._stop)

        bar.addWidget(self.scan_btn)
        bar.addWidget(self.preview_btn)
        bar.addStretch(1)
        bar.addWidget(self.run_btn)
        bar.addWidget(self.stop_btn)
        layout.addLayout(bar)

        self.progress = ProgressPanel()
        layout.addWidget(self.progress)

        inner = QSplitter(Qt.Orientation.Vertical)
        self.table = QTableWidget(0, len(COLUMNS))
        self.table.setHorizontalHeaderLabels(COLUMNS)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.itemDoubleClicked.connect(self._show_steps)
        header = self.table.horizontalHeader()
        for index in range(len(COLUMNS)):
            mode = (QHeaderView.ResizeMode.Stretch if index in (1, 6)
                    else QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(index, mode)
        inner.addWidget(self.table)

        self.console = ConsoleView()
        inner.addWidget(self.console)
        inner.setSizes([380, 300])
        layout.addWidget(inner, 1)

        bottom = QHBoxLayout()
        bottom.setSpacing(8)
        self.export_btn = QPushButton("단계별 결과 CSV 저장")
        self.export_btn.setProperty("secondary", True)
        self.export_btn.clicked.connect(self._export)
        clear_btn = QPushButton("로그 지우기")
        clear_btn.setProperty("secondary", True)
        clear_btn.clicked.connect(self.console.clear)
        bottom.addWidget(self.export_btn)
        bottom.addWidget(clear_btn)
        bottom.addStretch(1)
        self.summary = QLabel("")
        self.summary.setProperty("hint", True)
        bottom.addWidget(self.summary)
        layout.addLayout(bottom)
        return panel

    # ── 파일 스캔 ────────────────────────────────────
    def _invalidate_files(self, *_) -> None:
        if not self.files:
            return
        self.files = []
        self.scan_root = None
        self.table.setRowCount(0)
        self.rows.clear()
        self.progress.reset()
        self.statusBar().showMessage("검색 조건이 바뀌었습니다 — [파일 찾기]를 다시 눌러주세요")

    def _scan(self) -> None:
        root = self._root()
        if root is None:
            return
        self.folder.remember()
        self.scan_root = root
        self.scan_btn.setEnabled(False)
        self.statusBar().showMessage("파일을 찾는 중…")

        worker = ScanWorker(root, self.prefix.text().strip(),
                            self.pattern.text().strip() or "*.xlsx", self.limit.value())
        worker.found.connect(self._on_scanned)
        worker.failed.connect(lambda m: self.console.append_message(f"❌ 파일 검색 실패: {m}", "error"))
        worker.finished.connect(lambda: self.scan_btn.setEnabled(True))
        self.helper = worker
        worker.start()

    def _on_scanned(self, files: List[Path]) -> None:
        self.files = files
        self._fill_table(files)
        self.statusBar().showMessage(f"대상 파일 {len(files)}개")
        if files:
            self.console.append_message(f"대상 파일 {len(files)}개를 찾았습니다.", "success")
        else:
            self.console.append_message(
                f"조건에 맞는 파일이 없습니다. '파일명 시작'"
                f"({self.prefix.text().strip() or '전체'})과 확장자를 확인해 보세요.", "warning"
            )

    def _fill_table(self, files: List[Path]) -> None:
        # 표를 그리는 중에는 대화상자를 띄우지 않는다 (_root() 를 쓰지 않는 이유)
        self.table.setRowCount(len(files))
        self.rows.clear()
        self.row_paths = list(files)
        for row, path in enumerate(files):
            self.rows[str(path)] = row
            # 폴더마다 파일명이 같을 수 있으므로 행에 전체 경로를 함께 저장한다
            name_item = QTableWidgetItem(path.name)
            name_item.setData(Qt.ItemDataRole.UserRole, str(path))
            name_item.setToolTip(str(path))
            self.table.setItem(row, 0, name_item)
            self.table.setItem(row, 1, QTableWidgetItem(_relative(path.parent, self.scan_root)))
            self._set_status(row, "pending")
            for column in range(3, len(COLUMNS)):
                self.table.setItem(row, column, QTableWidgetItem(""))

    def _set_status(self, row: int, status: str) -> None:
        item = QTableWidgetItem(STATUS_LABELS.get(status, status))
        item.setForeground(QColor(STATUS_COLORS.get(status, "#1d1d1f")))
        self.table.setItem(row, 2, item)

    # ── 매크로 목록 ──────────────────────────────────
    def _load_macro_names(self) -> None:
        self.macro_btn.setEnabled(False)
        self.statusBar().showMessage("매크로 목록을 읽는 중… (Excel 이 잠깐 실행됩니다)")
        self.console.append_message("매크로 목록을 읽는 중…")

        worker = MacroListWorker(self._personal_path())
        worker.listed.connect(self._on_macros)
        worker.failed.connect(lambda m: self.console.append_message(f"❌ {m}", "error"))
        worker.finished.connect(lambda: self.macro_btn.setEnabled(True))
        self.helper = worker
        worker.start()

    def _on_macros(self, names: List[str], warnings: List[str]) -> None:
        for warning in warnings:
            self.console.append_message(f"⚠ {warning}", "warning")
        if not names:
            self.console.append_message(
                "매크로를 찾지 못했습니다. 매크로 파일 경로를 확인하거나 "
                "Alt+F11 에서 직접 이름을 확인하세요.", "warning"
            )
            return
        for box in self._macro_boxes():
            current = box.currentText().strip()
            box.clear()
            box.addItems(names)
            box.setCurrentText(current)
        self.console.append_message(
            f"매크로 {len(names)}개를 찾았습니다. 각 단계의 드롭다운에서 고르세요.", "success"
        )

    def _macro_boxes(self) -> List[QComboBox]:
        return [self.macro_overview, self.macro_nav, self.macro_inv, self.macro_interest]

    # ── 실행 ─────────────────────────────────────────
    def _start(self, preview: bool) -> None:
        root = self._root()
        if root is None:
            return
        self.folder.remember()
        self.scan_root = root

        files = self.files
        if not files:
            self.console.append_message("파일 목록이 없어 지금 검색합니다…")
            files = bot.collect_files(root, self.prefix.text().strip(),
                                      self.pattern.text().strip() or "*.xlsx")
            if self.limit.value():
                files = files[: self.limit.value()]
            self.files = files
        if not files:
            QMessageBox.information(
                self, "대상 없음",
                f"조건에 맞는 파일이 없습니다.\n\n폴더: {root}\n"
                f"파일명 시작: {self.prefix.text().strip() or '(전체)'}\n"
                f"확장자: {self.pattern.text().strip() or '*.xlsx'}",
            )
            return

        stop_after = fp.READONLY_STEPS if preview else self.stop_after.value()
        saves = stop_after == 0

        if not preview:
            missing = self._inputs().missing()
            if stop_after == 0 or stop_after > fp.READONLY_STEPS:
                if missing:
                    QMessageBox.warning(
                        self, "실행 정보 필요",
                        "아래 항목을 채워주세요 (E5~E9 에 들어갑니다):\n\n· "
                        + "\n· ".join(missing),
                    )
                    return
            empty_macros = [box for box in self._macro_boxes() if not box.currentText().strip()]
            if empty_macros:
                QMessageBox.warning(self, "매크로 이름 필요", "매크로 이름을 모두 입력하세요.")
                return

        if saves:
            answer = QMessageBox.question(
                self, "실행 확인",
                f"파일 {len(files)}개를 11단계로 처리하고 <b>덮어쓰기 저장</b>합니다.<br><br>"
                "각 파일에서 시트 내용과 이름 관리자를 <b>모두 지우고</b> 새로 생성합니다.<br><br>"
                "· Excel 을 모두 닫으셨나요?<br>"
                "· 폴더를 백업해 두셨나요?<br><br>"
                "계속할까요?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return

        self.results = []
        self.by_path.clear()
        self.summary.setText("")
        self._fill_table(files)
        self.progress.start(len(files))
        self._set_running(True)

        if preview:
            label = "읽기만 (1~3단계) — 파일을 수정하지 않습니다"
        elif stop_after:
            label = f"{stop_after}단계까지만 — 저장하지 않습니다"
        else:
            label = "전체 실행 (11단계) — 저장합니다"
        self.console.append_message(f"── {label} ──")
        if saves:
            self.console.append_message("실행 중에는 Excel 창을 건드리지 마세요.", "warning")

        self.worker = PipelineWorker(
            cfg=self._config(), inputs=self._inputs(), dialog=self._dialog_options(),
            files=files, personal=self._personal_path(), stop_after=stop_after,
            keep_open=self.keep_open.isChecked(),
        )
        self.worker.message.connect(self.console.append_message)
        self.worker.file_started.connect(self._on_file_started)
        self.worker.file_done.connect(self._on_file_done)
        self.worker.completed.connect(self._on_completed)
        self.worker.failed.connect(self._on_failed)
        self.worker.finished.connect(lambda: self._set_running(False))
        self.worker.start()

    def _stop(self) -> None:
        if self.worker is not None and self.worker.isRunning():
            self.worker.stop()
            self.stop_btn.setEnabled(False)
            self.console.append_message(
                "중지 요청 — 지금 처리 중인 파일이 끝나면 멈춥니다.", "warning"
            )

    def _on_file_started(self, index: int, total: int, path: Path) -> None:
        row = self.rows.get(str(path))
        if row is not None:
            self._set_status(row, "running")
            self.table.scrollToItem(self.table.item(row, 0))
        self.progress.advance(index - 1, total, path.name)

    def _on_file_done(self, index: int, total: int, result: fp.FileResult) -> None:
        self.results.append(result)
        self.by_path[str(result.path)] = result
        row = self.rows.get(str(result.path))
        if row is not None:
            self._set_status(row, result.status)
            summary = result.failed_step or result.detail
            for column, text in ((3, result.fund_id), (4, result.book_date),
                                 (5, str(result.item_count) if result.items_counted else ""),
                                 (6, summary)):
                item = QTableWidgetItem(text)
                if column == 6:
                    item.setToolTip(result.detail)      # 잘린 내용을 마우스로 확인
                    if result.status == "failed":
                        item.setForeground(QColor(STATUS_COLORS["failed"]))
                self.table.setItem(row, column, item)
        self.progress.advance(index, total, result.path.name)

    def _on_completed(self, results: List[fp.FileResult]) -> None:
        counts = {"ok": 0, "skipped": 0, "failed": 0}
        for item in results:
            counts[item.status] += 1
        text = f"완료 {counts['ok']} · 확인용 {counts['skipped']} · 실패 {counts['failed']}"
        self.summary.setText(text)
        self.progress.finish(text)
        self.statusBar().showMessage(text)
        self.console.append_message(f"── 작업 종료: {text} ──",
                                    "error" if counts["failed"] else "success")

        # 어느 단계에서 몇 건이 걸렸는지 한눈에
        by_step: Dict[str, int] = {}
        for item in results:
            if item.status == "failed":
                by_step[item.failed_step] = by_step.get(item.failed_step, 0) + 1
        if by_step:
            self.console.append_message("실패한 단계별 건수:", "error")
            for step, count in sorted(by_step.items()):
                self.console.append_message(f"  {step} — {count}건", "error")
            self.console.append_message(
                "표에서 파일을 더블클릭하면 그 파일의 단계별 상세를 볼 수 있습니다.", "warning"
            )

    def _on_failed(self, message: str) -> None:
        self.console.append_message(f"❌ {message}", "error")
        self.progress.reset()
        self.statusBar().showMessage("실행 실패")
        QMessageBox.critical(self, "실행 실패", message)

    def _set_running(self, running: bool) -> None:
        for button in (self.scan_btn, self.preview_btn, self.run_btn, self.macro_btn):
            button.setEnabled(not running)
        self.stop_btn.setEnabled(running)

    def _show_steps(self, item: QTableWidgetItem) -> None:
        """표를 더블클릭하면 그 파일의 11단계 결과를 콘솔에 펼친다."""
        name_item = self.table.item(item.row(), 0)
        path = name_item.data(Qt.ItemDataRole.UserRole) if name_item else None
        target = self.by_path.get(path) if path else None
        if target is None:
            self.console.append_message("아직 처리되지 않은 파일입니다.", "warning")
            return
        where = _relative(target.path.parent, self.scan_root)
        self.console.append_message(f"── {where} / {target.path.name} 단계별 결과 ──")
        for step in target.steps:
            icon = {"ok": "✓", "failed": "❌", "skipped": "-"}[step.status]
            level = {"ok": "success", "failed": "error", "skipped": "info"}[step.status]
            detail = f" — {step.detail}" if step.detail else ""
            self.console.append_message(f"  {icon} {step.label}{detail}", level)

    # ── 설정 조립 ────────────────────────────────────
    def _root(self) -> Optional[Path]:
        text = self.folder.path()
        if not text:
            QMessageBox.warning(self, "폴더 필요", "대상 폴더를 선택하세요.")
            return None
        path = Path(text).expanduser()
        if not path.is_dir():
            QMessageBox.warning(self, "폴더 없음", f"폴더를 찾을 수 없습니다:\n{path}")
            return None
        return path

    def _personal_path(self) -> Optional[Path]:
        if self.no_personal.isChecked():
            return None
        text = self.personal.text().strip()
        return Path(text) if text else bot.default_personal_path()

    def _inputs(self) -> fp.FundInputs:
        return fp.FundInputs(
            eval_date=self.eval_date.text().strip(),
            var_date=self.var_date.text().strip(),
            evaluator=self.evaluator.text().strip(),
            reviewer=self.reviewer.text().strip(),
        )

    def _config(self) -> fp.PipelineConfig:
        return fp.PipelineConfig(
            anchor_name=self.anchor_name.text().strip() or "F_1",
            fund_offset=(self.fund_down.value(), self.fund_right.value()),
            fund_prefix=self.fund_prefix.text().strip(),
            nav_suffix=self.nav_suffix.text().strip() or "_nav",
            nav_offset=(self.nav_down.value(), self.nav_right.value()),
            inv_suffix=self.inv_suffix.text().strip() or "_inv",
            inv_first_row_offset=self.inv_first.value(),
            alt_sheet=self.alt_sheet.text().strip(),
            clear_mode="delete" if self.clear_mode.currentIndex() == 0 else "clear",
            overview_cell=self.overview_cell.text().strip() or "B2",
            header_cell=self.header_cell.text().strip() or "E5",
            nav_cell=self.nav_cell.text().strip() or "B16",
            inv_cell=self.inv_cell.text().strip() or "H16",
            items_cell=self.items_cell.text().strip() or "H21",
            interest_row_offset=self.interest_down.value(),
            macro_overview=self.macro_overview.currentText().strip(),
            macro_nav=self.macro_nav.currentText().strip(),
            macro_inv=self.macro_inv.currentText().strip(),
            macro_interest=self.macro_interest.currentText().strip(),
        )

    def _dialog_options(self) -> bot.DialogOptions:
        return bot.DialogOptions(
            value="",
            title_contains=self.dialog_title.text().strip(),
            appear_timeout=self.appear.value(),
            confirm_timeout=self.confirm.value(),
            input_method="settext" if self.input_method.currentIndex() == 0 else "chars",
            dismiss_followup=self.dismiss.isChecked(),
            field_index=self.field_index.value(),
        )

    def _browse_personal(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "매크로 파일 선택", self.personal.text().strip(),
            "Excel 매크로 파일 (*.xlsb *.xlsm *.xlam);;모든 파일 (*.*)",
        )
        if path:
            self.personal.setText(path)

    # ── 결과 저장 ────────────────────────────────────
    def _export(self) -> None:
        if not self.results:
            QMessageBox.information(self, "결과 없음", "저장할 결과가 없습니다.")
            return
        default = f"수익보고서_결과_{datetime.now():%Y%m%d_%H%M%S}.csv"
        path, _ = QFileDialog.getSaveFileName(self, "결과 저장", default, "CSV (*.csv)")
        if not path:
            return
        with open(path, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(["파일명", "경로", "파일상태", "펀드아이디", "장부기준일", "종목수",
                             "단계", "단계상태", "상세", "소요(초)"])
            for result in self.results:
                for step in result.steps:
                    writer.writerow([
                        result.path.name, str(result.path), result.status,
                        result.fund_id, result.book_date, result.item_count,
                        step.label, step.status, step.detail, f"{step.elapsed:.2f}",
                    ])
        self.console.append_message(f"결과를 저장했습니다: {path}", "success")

    # ── 설정 저장/복원 ───────────────────────────────
    def _persistable(self):
        """objectName 을 지정해 둔 위젯만 저장 대상으로 삼는다."""
        for kind in (QLineEdit, QComboBox, QCheckBox, QSpinBox, QDoubleSpinBox):
            for widget in self.findChildren(kind):
                name = widget.objectName()
                if name and not name.startswith("qt_"):
                    yield name, widget

    def _load_settings(self) -> None:
        """저장된 값이 있는 항목만 덮어쓴다.

        QSettings.value(key, None, str) 는 키가 없어도 빈 문자열을 돌려주므로
        기본값이 지워진다. 반드시 contains() 로 존재 여부를 먼저 확인할 것.
        """
        store = settings()
        if store.contains("root/last"):
            self.folder.set_path(store.value("root/last", "", str))

        for name, widget in self._persistable():
            key = f"field/{name}"
            if not store.contains(key):
                continue
            if isinstance(widget, QCheckBox):
                widget.setChecked(store.value(key, False, bool))
            elif isinstance(widget, (QSpinBox, QDoubleSpinBox)):
                widget.setValue(type(widget.value())(store.value(key)))
            elif isinstance(widget, QComboBox):
                if widget.isEditable():
                    widget.setCurrentText(store.value(key, "", str))
                else:
                    widget.setCurrentIndex(store.value(key, 0, int))
            else:
                widget.setText(store.value(key, "", str))

    def _save_settings(self) -> None:
        store = settings()
        store.setValue("root/last", self.folder.path())
        for name, widget in self._persistable():
            key = f"field/{name}"
            if isinstance(widget, QCheckBox):
                store.setValue(key, widget.isChecked())
            elif isinstance(widget, (QSpinBox, QDoubleSpinBox)):
                store.setValue(key, widget.value())
            elif isinstance(widget, QComboBox):
                store.setValue(key, widget.currentText() if widget.isEditable()
                               else widget.currentIndex())
            else:
                store.setValue(key, widget.text())

    def closeEvent(self, event) -> None:
        if self.worker is not None and self.worker.isRunning():
            answer = QMessageBox.question(
                self, "작업 진행 중",
                "매크로를 실행하는 중입니다. 정말 종료할까요?\n"
                "Excel 이 열린 채로 남을 수 있습니다.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
            self.worker.stop()
            self.worker.wait(3000)
        self._save_settings()
        event.accept()
