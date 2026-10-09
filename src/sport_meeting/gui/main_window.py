from __future__ import annotations

import re
import tempfile
from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, QDate, QLocale, QModelIndex, QSettings, QTimer, Qt
from PySide6.QtGui import QAction, QColor, QFont, QKeySequence
from PySide6.QtWidgets import (
    QApplication,
    QCalendarWidget,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QSplitter,
    QStackedWidget,
    QStatusBar,
    QStyledItemDelegate,
    QTableView,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from sport_meeting.core.booklet_renderer import render_competition_groups, render_full_booklet
from sport_meeting.core.database import ProjectDatabase
from sport_meeting.core.finals_workflow import compute_track_final
from sport_meeting.core.heat_generation import generate_heat_assignments
from sport_meeting.core.legacy_booklet_content import (
    LATEST_SCHOOL_RECORDS,
    last_year_booklet_values,
)
from sport_meeting.core.importers import (
    RegistrationImport,
    ScheduleImport,
    import_registration_workbooks,
    import_schedule_workbook,
)
from sport_meeting.core.models import (
    ROUND_MODE_PRELIMINARY_FINAL,
    ROUND_MODE_TIMED_FINAL,
    TEAM_COUNT_MANUAL,
    TEAM_COUNT_ONE_PER_CLASS,
    TEAM_GROUP_MEN,
    TEAM_GROUP_MIXED,
    TEAM_GROUP_STAFF,
    TEAM_GROUP_WOMEN,
    TEAM_ROUND_MODE_KNOCKOUT,
    EventType,
    Participant,
    PerformanceKind,
    Result,
    ResultStatus,
    RoundType,
    Severity,
    TeamEventConfig,
    ProjectConfig,
)
from sport_meeting.core.performance import (
    PerformanceParseError,
    format_canonical,
    parse_result,
    recommended_input_unit,
)
from sport_meeting.core.qualification import rank_final_results
from sport_meeting.core.schedule_planner import (
    ScheduleCapacityError,
    build_standard_schedule,
    guess_event_round_mode,
)
from sport_meeting.core.reporting import (
    FinalMaterial,
    FinalRankingMaterial,
    render_final_assignments_batch,
    render_on_site_package,
    render_result_rankings_batch,
    render_validation_report,
)
from sport_meeting.core.validation import validate_project
from sport_meeting.core.workbook_templates import (
    create_school_level_registration_template,
    create_schedule_template,
)
from sport_meeting.export.office import ExportManager


class SimpleTableModel(QAbstractTableModel):
    def __init__(self, headers: list[str], rows: list[list[object]], parent=None):
        super().__init__(parent)
        self.headers = headers
        self.rows = rows

    def rowCount(self, _parent=QModelIndex()):
        return len(self.rows)

    def columnCount(self, _parent=QModelIndex()):
        return len(self.headers)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.ItemDataRole.DisplayRole:
            value = self.rows[index.row()][index.column()]
            return "" if value is None else str(value)
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        return self.headers[section] if orientation == Qt.Orientation.Horizontal else str(section + 1)


class ScoreTableWidget(QTableWidget):
    SCORE_COLUMN = 5

    def keyPressEvent(self, event) -> None:
        if event.matches(QKeySequence.StandardKey.Paste):
            self._paste_scores()
            return
        super().keyPressEvent(event)

    def _paste_scores(self) -> None:
        if self.currentColumn() != self.SCORE_COLUMN or self.currentRow() < 0:
            QMessageBox.information(self, "请选择成绩列", "请先点击成绩列中的起始单元格，再粘贴一列成绩。")
            return
        lines = QApplication.clipboard().text().replace("\r\n", "\n").replace("\r", "\n").split("\n")
        while lines and not lines[-1].strip():
            lines.pop()
        if not lines:
            return
        if any("\t" in line for line in lines):
            QMessageBox.warning(self, "粘贴列数不符", "成绩页一次只接受一列成绩，请在 Excel 中只复制成绩列。")
            return
        if self.currentRow() + len(lines) > self.rowCount():
            QMessageBox.warning(self, "粘贴行数过多", "粘贴内容超出当前项目的运动员行数，请检查多余行。")
            return
        for offset, value in enumerate(lines):
            self.setItem(self.currentRow() + offset, self.SCORE_COLUMN, QTableWidgetItem(value.strip()))


class ScoreInputDelegate(QStyledItemDelegate):
    def __init__(self, unit_provider, parent=None):
        super().__init__(parent)
        self.unit_provider = unit_provider

    def createEditor(self, parent, _option, _index):
        editor = QLineEdit(parent)
        unit = self.unit_provider()
        if unit == "SECOND":
            editor.setInputMask("99:99;_")
            editor.setPlaceholderText("秒:百分秒，例如 59:32")
        elif unit == "MINUTE":
            editor.setInputMask("9:99:99;_")
            editor.setPlaceholderText("分:秒:百分秒，例如 1:53:29")
        return editor

    def setEditorData(self, editor, index) -> None:
        value = str(index.data(Qt.ItemDataRole.EditRole) or "").strip()
        unit = self.unit_provider()
        if unit == "SECOND":
            match = re.fullmatch(r"(\d+)\.(\d{1,2})", value)
            if match:
                value = f"{int(match.group(1)):02d}:{match.group(2):0<2}"
            else:
                match = re.fullmatch(r"(\d+):(\d{1,2})", value)
                if match:
                    value = f"{int(match.group(1)):02d}:{match.group(2):0<2}"
        elif unit == "MINUTE":
            match = re.fullmatch(r"(\d+):(\d{1,2})\.(\d{1,2})", value)
            if match:
                value = f"{match.group(1)}:{int(match.group(2)):02d}:{match.group(3):0<2}"
        editor.setText(value)

    def setModelData(self, editor, model, index) -> None:
        value = editor.text().replace("_", "").replace(" ", "").strip()
        if self.unit_provider() == "SECOND" and ":" in value:
            seconds, fraction = value.split(":", 1)
            if seconds and fraction:
                value = f"{int(seconds)}:{fraction}"
        model.setData(index, value, Qt.ItemDataRole.EditRole)


class MainWindow(QMainWindow):
    def __init__(self, logger):
        super().__init__()
        self.logger = logger
        self.database: ProjectDatabase | None = None
        self.registration: RegistrationImport | None = None
        self.schedule: ScheduleImport | None = None
        self.latest_report = None
        self.export_manager = ExportManager(logger)
        self.setWindowTitle("学校运动会秩序册与决赛编排系统")
        self.resize(1180, 760)
        self._build_ui()
        QTimer.singleShot(0, self.open_last_project)

    def _build_ui(self) -> None:
        self.navigation = QListWidget()
        self.navigation.addItems([
            "1  项目首页", "2  项目与日程", "3  报名表设置", "4  批量报名",
            "5  核对修正", "6  抽签编排", "7  生成秩序册", "8  成绩录入",
            "9  决赛材料",
        ])
        self.navigation.setFixedWidth(190)
        self.pages = QStackedWidget()
        self.pages.addWidget(self._home_page())
        self.pages.addWidget(self._schedule_page())
        self.pages.addWidget(self._registration_template_settings_page())
        self.pages.addWidget(self._registration_page())
        self.pages.addWidget(self._validation_page())
        self.pages.addWidget(self._grouping_page())
        self.pages.addWidget(self._booklet_page())
        self.pages.addWidget(self._results_page())
        self.pages.addWidget(self._final_materials_page())
        self.navigation.currentRowChanged.connect(self.pages.setCurrentIndex)
        self.navigation.setCurrentRow(0)

        splitter = QSplitter()
        splitter.addWidget(self.navigation)
        splitter.addWidget(self.pages)
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(splitter)
        self.setCentralWidget(container)
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("请新建或打开一届运动会项目")

        menu = self.menuBar().addMenu("项目")
        new_action = QAction("新建项目", self)
        new_action.triggered.connect(self.new_project)
        open_action = QAction("打开项目", self)
        open_action.triggered.connect(self.open_project)
        menu.addActions([new_action, open_action])

    def _home_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        title = QLabel("学校运动会秩序册与决赛编排系统")
        title.setStyleSheet("font-size: 26px; font-weight: 700; padding: 18px 0;")
        layout.addWidget(title)
        self.project_label = QLabel("当前项目：未打开")
        self.project_label.setStyleSheet("font-size: 15px;")
        layout.addWidget(self.project_label)

        date_group = QGroupBox("运动会时间")
        date_layout = QHBoxLayout(date_group)
        self.meeting_start_date_edit = self._create_meeting_date_edit()
        self.meeting_end_date_edit = self._create_meeting_date_edit()
        today = QDate.currentDate()
        self.meeting_start_date_edit.setDate(today)
        self.meeting_end_date_edit.setDate(today.addDays(1))
        save_dates_button = QPushButton("保存运动会日期")
        save_dates_button.clicked.connect(self.save_meeting_dates)
        date_layout.addWidget(QLabel("开始日期"))
        date_layout.addWidget(self.meeting_start_date_edit)
        date_layout.addWidget(QLabel("结束日期"))
        date_layout.addWidget(self.meeting_end_date_edit)
        date_layout.addWidget(save_dates_button)
        date_layout.addStretch(1)
        layout.addWidget(date_group)
        self.meeting_dates_label = QLabel("尚未保存运动会日期")
        self.meeting_dates_label.setStyleSheet("color: #667085; padding-bottom: 6px;")
        layout.addWidget(self.meeting_dates_label)

        buttons = QHBoxLayout()
        for text, slot in [
            ("新建运动会项目", self.new_project),
            ("打开已有项目", self.open_project),
            ("批量导入报名表", self.import_registrations),
            ("导出核对报告", self.export_validation_report),
        ]:
            button = QPushButton(text)
            button.setMinimumHeight(42)
            button.clicked.connect(slot)
            buttons.addWidget(button)
        layout.addLayout(buttons)

        self.summary_label = QLabel("尚未导入报名数据")
        self.summary_label.setStyleSheet("padding: 12px; background: #f3f5f7; border-radius: 6px;")
        layout.addWidget(self.summary_label)
        self.athlete_table = QTableView()
        self.athlete_table.setAlternatingRowColors(True)
        layout.addWidget(self.athlete_table, 1)
        return page

    @staticmethod
    def _create_meeting_date_edit() -> QDateEdit:
        """日期控件固定使用支持中文和阿拉伯数字的 Windows 字体。

        部分 Windows 环境会为 QCalendarWidget 选中符号字体，导致 0-9
        显示成圆圈、竖线等异常字形。
        """
        locale = QLocale(QLocale.Language.Chinese, QLocale.Country.China)
        font = QFont("Microsoft YaHei UI", 11)
        editor = QDateEdit()
        editor.setLocale(locale)
        editor.setFont(font)
        editor.setCalendarPopup(True)
        editor.setDisplayFormat("yyyy年M月d日")
        editor.setMinimumWidth(205)
        calendar = QCalendarWidget()
        calendar.setLocale(locale)
        calendar.setFont(font)
        calendar.setStyleSheet(
            'QCalendarWidget QWidget { font-family: "Microsoft YaHei UI"; font-size: 11pt; }'
        )
        editor.setCalendarWidget(calendar)
        return editor

    def _registration_template_settings_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        title = QLabel("报名表设置")
        title.setStyleSheet("font-size: 22px; font-weight: 700; padding: 16px 0;")
        layout.addWidget(title)
        help_text = QLabel(
            "分别设置初中、高中的男生和女生可报项目。新增项目后系统会先预判赛制；"
            "如预判不符合本届安排，可直接用下拉框修改，保存后随当前运动会项目保留。"
        )
        help_text.setWordWrap(True)
        help_text.setStyleSheet("color: #667085; padding-bottom: 8px;")
        layout.addWidget(help_text)

        settings_tabs = QTabWidget()
        personal_page = QWidget()
        personal_layout = QVBoxLayout(personal_page)
        self.registration_event_editors = {}
        grid = QGridLayout()
        editor_specs = (
            ("junior_boys_events", "初中男生项目", 0, 0),
            ("junior_girls_events", "初中女生项目", 0, 1),
            ("senior_boys_events", "高中男生项目", 1, 0),
            ("senior_girls_events", "高中女生项目", 1, 1),
        )
        for field_name, label, row, column in editor_specs:
            group = QGroupBox(label)
            group_layout = QVBoxLayout(group)
            editor = QTableWidget(0, 2)
            editor.setHorizontalHeaderLabels(("项目名称", "赛制"))
            editor.setColumnWidth(0, 185)
            editor.setColumnWidth(1, 125)
            editor.setMinimumHeight(180)
            editor.itemChanged.connect(
                lambda item, table=editor: self._on_registration_event_name_changed(table, item)
            )
            group_layout.addWidget(editor)
            row_buttons = QHBoxLayout()
            add_button = QPushButton("新增项目")
            add_button.clicked.connect(
                lambda _checked=False, table=editor: self._add_registration_event_row(table)
            )
            batch_button = QPushButton("批量添加")
            batch_button.clicked.connect(
                lambda _checked=False, table=editor: self._batch_add_registration_events(table)
            )
            delete_button = QPushButton("删除选中")
            delete_button.clicked.connect(
                lambda _checked=False, table=editor: self._delete_selected_registration_events(table)
            )
            row_buttons.addWidget(add_button)
            row_buttons.addWidget(batch_button)
            row_buttons.addWidget(delete_button)
            group_layout.addLayout(row_buttons)
            grid.addWidget(group, row, column)
            self.registration_event_editors[field_name] = editor
        personal_layout.addLayout(grid)
        settings_tabs.addTab(personal_page, "个人项目")

        team_page = QWidget()
        team_layout = QVBoxLayout(team_page)
        team_help = QLabel(
            "每班1队会按照报名表中实际出现的班级自动计算；拔河和教工项目使用手动队数，"
            "可先留空，生成的日程模板中再补充。"
        )
        team_help.setWordWrap(True)
        team_help.setStyleSheet("color: #667085;")
        team_layout.addWidget(team_help)
        self.team_event_table = QTableWidget(0, 6)
        self.team_event_table.setHorizontalHeaderLabels(
            ("项目名称", "参赛组别", "参赛范围", "队数来源", "手动队数", "赛制")
        )
        for column, width in enumerate((180, 105, 115, 110, 85, 150)):
            self.team_event_table.setColumnWidth(column, width)
        self.team_event_table.setMinimumHeight(290)
        self.team_event_table.itemChanged.connect(self._on_team_event_name_changed)
        team_layout.addWidget(self.team_event_table)
        team_buttons = QHBoxLayout()
        add_team_button = QPushButton("新增集体项目")
        add_team_button.clicked.connect(
            lambda: self._add_team_event_row(self.team_event_table)
        )
        delete_team_button = QPushButton("删除选中")
        delete_team_button.clicked.connect(
            lambda: self._delete_selected_registration_events(self.team_event_table)
        )
        team_buttons.addWidget(add_team_button)
        team_buttons.addWidget(delete_team_button)
        team_buttons.addStretch(1)
        team_layout.addLayout(team_buttons)
        settings_tabs.addTab(team_page, "集体项目")
        layout.addWidget(settings_tabs, 1)

        buttons = QHBoxLayout()
        save_button = QPushButton("保存报名项目设置")
        save_button.setMinimumHeight(44)
        save_button.clicked.connect(self.save_registration_template_settings)
        junior_button = QPushButton("导出初中报名表模板")
        junior_button.setMinimumHeight(44)
        junior_button.clicked.connect(lambda: self.export_level_registration_template("初中"))
        senior_button = QPushButton("导出高中报名表模板")
        senior_button.setMinimumHeight(44)
        senior_button.clicked.connect(lambda: self.export_level_registration_template("高中"))
        buttons.addWidget(save_button)
        buttons.addWidget(junior_button)
        buttons.addWidget(senior_button)
        layout.addLayout(buttons)
        self._load_registration_template_settings(ProjectConfig())
        return page

    def _schedule_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        title = QLabel("项目与日程")
        title.setStyleSheet("font-size: 22px; font-weight: 700; padding: 16px 0;")
        layout.addWidget(title)
        help_text = QLabel("导入体育老师提交的 Excel 日程表；系统会识别新格式及老系统格式。")
        help_text.setStyleSheet("color: #667085; padding-bottom: 8px;")
        layout.addWidget(help_text)
        button_row = QHBoxLayout()
        import_button = QPushButton("导入比赛日程 Excel")
        import_button.setMinimumHeight(42)
        import_button.clicked.connect(self.import_schedule)
        template_button = QPushButton("生成标准日程模板")
        template_button.setMinimumHeight(42)
        template_button.clicked.connect(self.export_schedule_template)
        button_row.addWidget(import_button)
        button_row.addWidget(template_button)
        layout.addLayout(button_row)
        self.schedule_summary_label = QLabel("尚未导入比赛日程")
        self.schedule_summary_label.setStyleSheet("padding: 12px; background: #f3f5f7; border-radius: 6px;")
        layout.addWidget(self.schedule_summary_label)
        self.schedule_table = QTableView()
        self.schedule_table.setAlternatingRowColors(True)
        layout.addWidget(self.schedule_table, 1)
        return page

    def _placeholder_page(self, title_text: str) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        title = QLabel(title_text)
        title.setStyleSheet("font-size: 22px; font-weight: 700; padding: 16px;")
        layout.addWidget(title)
        text = QLabel("该步骤的核心数据结构已经建立，界面将在后续开发中接入。")
        text.setStyleSheet("color: #667085; padding: 16px;")
        layout.addWidget(text)
        layout.addStretch(1)
        return page

    def _registration_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        title = QLabel("批量报名")
        title.setStyleSheet("font-size: 22px; font-weight: 700; padding: 16px 0;")
        layout.addWidget(title)
        help_text = QLabel(
            "可一次选择多个班级报名表，或选择包含各班报名表的整个文件夹；导入前建议先下发标准模板。"
        )
        help_text.setWordWrap(True)
        help_text.setStyleSheet("color: #667085; padding-bottom: 8px;")
        layout.addWidget(help_text)
        buttons = QHBoxLayout()
        select_button = QPushButton("选择多个报名表")
        select_button.setMinimumHeight(44)
        select_button.clicked.connect(self.import_registrations)
        folder_button = QPushButton("导入文件夹内全部报名表")
        folder_button.setMinimumHeight(44)
        folder_button.clicked.connect(self.import_registration_folder)
        template_button = QPushButton("前往报名表设置与导出")
        template_button.setMinimumHeight(44)
        template_button.clicked.connect(lambda: self.navigation.setCurrentRow(2))
        buttons.addWidget(select_button)
        buttons.addWidget(folder_button)
        buttons.addWidget(template_button)
        layout.addLayout(buttons)
        self.registration_summary_label = QLabel("尚未导入报名数据")
        self.registration_summary_label.setStyleSheet(
            "padding: 12px; background: #f3f5f7; border-radius: 6px;"
        )
        layout.addWidget(self.registration_summary_label)
        self.registration_table = QTableView()
        self.registration_table.setAlternatingRowColors(True)
        layout.addWidget(self.registration_table, 1)
        return page

    def _validation_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        title = QLabel("核对修正")
        title.setStyleSheet("font-size: 22px; font-weight: 700; padding: 16px 0;")
        layout.addWidget(title)
        help_text = QLabel("统一核对报名人数、日程人数、预赛与决赛对应关系，以及单组跑道容量。")
        help_text.setStyleSheet("color: #667085; padding-bottom: 8px;")
        layout.addWidget(help_text)
        button_row = QHBoxLayout()
        validate_button = QPushButton("开始全项目校验")
        validate_button.setMinimumHeight(42)
        validate_button.clicked.connect(self.run_project_validation)
        export_button = QPushButton("导出当前核对报告")
        export_button.setMinimumHeight(42)
        export_button.clicked.connect(self.export_validation_report)
        button_row.addWidget(validate_button)
        button_row.addWidget(export_button)
        layout.addLayout(button_row)
        self.validation_summary_label = QLabel("尚未执行全项目校验")
        self.validation_summary_label.setStyleSheet("padding: 12px; background: #f3f5f7; border-radius: 6px;")
        layout.addWidget(self.validation_summary_label)
        self.validation_table = QTableView()
        self.validation_table.setAlternatingRowColors(True)
        layout.addWidget(self.validation_table, 1)
        return page

    def _grouping_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        title = QLabel("抽签编排")
        title.setStyleSheet("font-size: 22px; font-weight: 700; padding: 16px 0;")
        layout.addWidget(title)
        help_text = QLabel("按项目日程生成可复现的分组；径赛分配道次，田赛生成出场顺序。")
        help_text.setStyleSheet("color: #667085; padding-bottom: 8px;")
        layout.addWidget(help_text)
        generate_button = QPushButton("生成分组 / 重新抽签")
        generate_button.setMinimumHeight(46)
        generate_button.clicked.connect(self.generate_heats)
        layout.addWidget(generate_button)
        self.grouping_summary_label = QLabel("尚未生成分组")
        self.grouping_summary_label.setStyleSheet("padding: 12px; background: #f3f5f7; border-radius: 6px;")
        layout.addWidget(self.grouping_summary_label)
        self.grouping_table = QTableView()
        self.grouping_table.setAlternatingRowColors(True)
        layout.addWidget(self.grouping_table, 1)
        return page

    def _booklet_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        title = QLabel("生成秩序册")
        title.setStyleSheet("font-size: 22px; font-weight: 700; padding: 16px 0;")
        layout.addWidget(title)
        help_text = QLabel(
            "可生成包含封面、目录、竞赛日程、代表队名单和竞赛分组表的完整秩序册；"
            "封面、议程、组委会、裁判名单、竞赛规程和校纪录可在本页编辑并随项目保存。"
        )
        help_text.setWordWrap(True)
        help_text.setStyleSheet("color: #667085; padding-bottom: 8px;")
        layout.addWidget(help_text)
        buttons = QHBoxLayout()
        edit_button = QPushButton("编辑封面与静态章节")
        edit_button.setMinimumHeight(48)
        edit_button.clicked.connect(self.edit_booklet_content)
        buttons.addWidget(edit_button)
        full_button = QPushButton("生成完整秩序册 DOCX")
        full_button.setMinimumHeight(48)
        full_button.clicked.connect(self.generate_full_booklet_document)
        buttons.addWidget(full_button)
        grouping_button = QPushButton("仅生成竞赛分组表 DOCX")
        grouping_button.setMinimumHeight(48)
        grouping_button.clicked.connect(self.generate_competition_document)
        buttons.addWidget(grouping_button)
        layout.addLayout(buttons)
        self.booklet_summary_label = QLabel("生成前请先完成全项目校验和抽签编排。")
        self.booklet_summary_label.setStyleSheet("padding: 12px; background: #f3f5f7; border-radius: 6px;")
        layout.addWidget(self.booklet_summary_label)
        layout.addStretch(1)
        return page

    def edit_booklet_content(self) -> None:
        if not self.database:
            QMessageBox.information(self, "请先新建项目", "请先新建或打开运动会项目。")
            return
        config = self.database.load_config()
        dialog = QDialog(self)
        dialog.setWindowTitle("编辑秩序册封面与静态章节")
        dialog.resize(900, 680)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        title_editor = QLineEdit(config.booklet_title)
        form.addRow("封面名称：", title_editor)
        layout.addLayout(form)
        tabs = QTabWidget()
        fields = (
            ("一、开幕式议程", "opening_ceremony_text"),
            ("二、组委会名单", "organizing_committee_text"),
            ("三、仲裁与裁判", "officials_text"),
            ("四、竞赛规程", "competition_rules_text"),
            ("八、校纪录", "school_records_text"),
        )
        editors: dict[str, QPlainTextEdit] = {}
        for label, field_name in fields:
            editor = QPlainTextEdit()
            editor.setPlainText(getattr(config, field_name))
            editor.setPlaceholderText("可直接从往届 Word 复制文字并粘贴到这里。")
            tabs.addTab(editor, label)
            editors[field_name] = editor
        layout.addWidget(tabs, 1)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        load_last_year_button = buttons.addButton(
            "载入2025年内容", QDialogButtonBox.ButtonRole.ActionRole
        )
        load_latest_records_button = buttons.addButton(
            "载入最新校纪录", QDialogButtonBox.ButtonRole.ActionRole
        )

        def load_last_year_content() -> None:
            answer = QMessageBox.question(
                dialog,
                "载入去年内容",
                "这会替换当前窗口中的封面和五个静态章节；只有点击“保存到项目”后才会写入项目。是否继续？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
            values = last_year_booklet_values()
            title_editor.setText(values["booklet_title"])
            for field_name, editor in editors.items():
                editor.setPlainText(values[field_name])

        load_last_year_button.clicked.connect(load_last_year_content)

        def load_latest_records() -> None:
            answer = QMessageBox.question(
                dialog,
                "载入最新校纪录",
                "将用公开版虚构示例更新校纪录编辑框，正式比赛请填写本校真实校纪录。"
                "其他静态章节不变；点击“保存到项目”后才会替换当前项目的校纪录。是否继续？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer == QMessageBox.StandardButton.Yes:
                editors["school_records_text"].setPlainText(LATEST_SCHOOL_RECORDS)
                tabs.setCurrentWidget(editors["school_records_text"])

        load_latest_records_button.clicked.connect(load_latest_records)
        buttons.button(QDialogButtonBox.StandardButton.Save).setText("保存到项目")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("取消")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            config.booklet_title = title_editor.text().strip()
            for field_name, editor in editors.items():
                setattr(config, field_name, editor.toPlainText().strip())
            config.validate()
            self.database.save_config(config)
            self.statusBar().showMessage("秩序册静态内容已保存")
            QMessageBox.information(self, "保存成功", "封面和五个静态章节已保存到当前项目。")
        except Exception as exc:
            QMessageBox.warning(self, "内容无法保存", str(exc))

    def _results_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        title = QLabel("成绩录入")
        title.setStyleSheet("font-size: 22px; font-weight: 700; padding: 16px 0;")
        layout.addWidget(title)
        help_text = QLabel("选择预赛项目后逐行录入，或从 Excel 复制一列成绩并粘贴到成绩列。")
        help_text.setStyleSheet("color: #667085; padding-bottom: 8px;")
        layout.addWidget(help_text)
        controls = QHBoxLayout()
        self.result_event_combo = QComboBox()
        self.result_event_combo.currentIndexChanged.connect(self.load_result_event)
        controls.addWidget(self.result_event_combo, 1)
        controls.addWidget(QLabel("成绩单位："))
        self.result_unit_combo = QComboBox()
        for name, value in [("分", "MINUTE"), ("秒", "SECOND"), ("米", "METER"), ("个", "COUNT")]:
            self.result_unit_combo.addItem(name, value)
        controls.addWidget(self.result_unit_combo)
        self.manual_rank_button = QPushButton("手动决定名次")
        self.manual_rank_button.setCheckable(True)
        self.manual_rank_button.toggled.connect(self.toggle_manual_ranking)
        controls.addWidget(self.manual_rank_button)
        clear_rank_button = QPushButton("清除手动名次")
        clear_rank_button.clicked.connect(self.clear_manual_ranks)
        controls.addWidget(clear_rank_button)
        layout.addLayout(controls)
        self.result_table = ScoreTableWidget()
        self.result_table.setColumnCount(8)
        self.result_table.setHorizontalHeaderLabels(
            ["组", "道次", "号码", "姓名", "班级", "成绩", "状态", "组内名次"]
        )
        self.score_input_delegate = ScoreInputDelegate(
            lambda: self.result_unit_combo.currentData(), self.result_table
        )
        self.result_table.setItemDelegateForColumn(
            ScoreTableWidget.SCORE_COLUMN, self.score_input_delegate
        )
        self.result_table.setAlternatingRowColors(True)
        self.result_table.cellClicked.connect(self.assign_manual_rank)
        layout.addWidget(self.result_table, 2)
        save_button = QPushButton("保存成绩并生成决赛名单")
        save_button.setMinimumHeight(46)
        save_button.clicked.connect(self.save_results_and_generate_final)
        layout.addWidget(save_button)
        preview_label = QLabel("实时决赛名单预览")
        preview_label.setStyleSheet("font-size: 16px; font-weight: 700; padding-top: 8px;")
        layout.addWidget(preview_label)
        self.final_preview_table = QTableView()
        self.final_preview_table.setAlternatingRowColors(True)
        layout.addWidget(self.final_preview_table, 1)
        self.result_rows = []
        self.manual_rank_by_participant = {}
        return page

    def _final_materials_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        title = QLabel("决赛材料")
        title.setStyleSheet("font-size: 22px; font-weight: 700; padding: 16px 0;")
        layout.addWidget(title)
        help_text = QLabel(
            "预览已生成的决赛名单；径赛显示道次，田赛显示出场序，可导出当前项目或全部项目。"
        )
        help_text.setStyleSheet("color: #667085; padding-bottom: 8px;")
        layout.addWidget(help_text)

        self.final_material_event_combo = QComboBox()
        self.final_material_event_combo.currentIndexChanged.connect(
            self.load_final_material_event
        )
        selection_row = QHBoxLayout()
        selection_row.addWidget(self.final_material_event_combo, 1)
        selection_row.addWidget(QLabel("成绩单位："))
        self.final_result_unit_combo = QComboBox()
        for name, value in [("分", "MINUTE"), ("秒", "SECOND"), ("米", "METER"), ("个", "COUNT")]:
            self.final_result_unit_combo.addItem(name, value)
        selection_row.addWidget(self.final_result_unit_combo)
        layout.addLayout(selection_row)

        buttons = QHBoxLayout()
        current_button = QPushButton("生成当前项目 DOCX/PDF")
        current_button.setMinimumHeight(44)
        current_button.clicked.connect(self.export_current_final_material)
        all_button = QPushButton("批量生成全部决赛 DOCX/PDF")
        all_button.setMinimumHeight(44)
        all_button.clicked.connect(self.export_all_final_materials)
        buttons.addWidget(current_button)
        buttons.addWidget(all_button)
        layout.addLayout(buttons)

        package_buttons = QHBoxLayout()
        current_package_button = QPushButton("生成当前项目现场材料包")
        current_package_button.setMinimumHeight(44)
        current_package_button.clicked.connect(self.export_current_on_site_package)
        all_package_button = QPushButton("批量生成全部现场材料包")
        all_package_button.setMinimumHeight(44)
        all_package_button.clicked.connect(self.export_all_on_site_packages)
        package_buttons.addWidget(current_package_button)
        package_buttons.addWidget(all_package_button)
        layout.addLayout(package_buttons)

        self.final_material_summary_label = QLabel("尚未生成决赛名单")
        self.final_material_summary_label.setStyleSheet(
            "padding: 12px; background: #f3f5f7; border-radius: 6px;"
        )
        layout.addWidget(self.final_material_summary_label)
        self.final_material_table = ScoreTableWidget()
        self.final_material_table.setColumnCount(8)
        self.final_material_table.setHorizontalHeaderLabels(
            ["决赛组", "道次/出场序", "号码", "姓名", "单位", "成绩", "状态", "名次"]
        )
        self.final_score_input_delegate = ScoreInputDelegate(
            lambda: self.final_result_unit_combo.currentData(), self.final_material_table
        )
        self.final_material_table.setItemDelegateForColumn(
            ScoreTableWidget.SCORE_COLUMN, self.final_score_input_delegate
        )
        self.final_material_table.setAlternatingRowColors(True)
        layout.addWidget(self.final_material_table, 1)
        result_buttons = QHBoxLayout()
        save_results_button = QPushButton("保存决赛成绩并计算名次")
        save_results_button.setMinimumHeight(44)
        save_results_button.clicked.connect(self.save_final_results)
        export_ranking_button = QPushButton("生成当前成绩排名表 DOCX/PDF")
        export_ranking_button.setMinimumHeight(44)
        export_ranking_button.clicked.connect(self.export_current_final_ranking)
        export_all_rankings_button = QPushButton("批量生成全部成绩排名表 DOCX/PDF")
        export_all_rankings_button.setMinimumHeight(44)
        export_all_rankings_button.clicked.connect(self.export_all_final_rankings)
        result_buttons.addWidget(save_results_button)
        result_buttons.addWidget(export_ranking_button)
        result_buttons.addWidget(export_all_rankings_button)
        layout.addLayout(result_buttons)
        self.final_result_rows = []
        return page

    def _close_database(self) -> None:
        if self.database:
            self.database.close()
            self.database = None

    def new_project(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "新建运动会项目", "", "运动会项目 (*.sportsmeet)")
        if not path:
            return
        if not path.lower().endswith(".sportsmeet"):
            path += ".sportsmeet"
        try:
            self._close_database()
            self.database = ProjectDatabase.create(path)
            self._set_project(path)
        except Exception as exc:
            QMessageBox.critical(self, "无法新建项目", str(exc))

    def open_project(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "打开运动会项目", "", "运动会项目 (*.sportsmeet)")
        if not path:
            return
        try:
            self._close_database()
            self.database = ProjectDatabase.open(path)
            self._set_project(path)
        except Exception as exc:
            QMessageBox.critical(self, "无法打开项目", str(exc))

    def _set_project(self, path: str) -> None:
        self.project_label.setText(f"当前项目：{Path(path).name}")
        self.statusBar().showMessage("项目已打开")
        QSettings("LeishiSchool", "SportMeetingBooklet").setValue(
            "last_project_path", str(Path(path).resolve())
        )
        self.logger.info("project opened", fields={"schema_version": self.database.schema_version})
        self._load_registration_template_settings(self.database.load_config())
        self._load_meeting_dates(self.database.load_config())
        athletes, entries = self.database.load_registration()
        if athletes or entries:
            self._show_registration(athletes, entries, 0, 0, 0)
        rounds = self.database.load_schedule()
        if rounds:
            self._show_schedule(rounds, 0, 0, 0)
        assignments = self.database.load_heat_assignments()
        if assignments:
            self._show_assignments(assignments)
        self._refresh_result_events()
        self._refresh_final_material_events()

    def open_last_project(self) -> None:
        if self.database:
            return
        settings = QSettings("LeishiSchool", "SportMeetingBooklet")
        saved_path = settings.value("last_project_path", "", str)
        if not saved_path:
            return
        path = Path(saved_path)
        if not path.is_file():
            settings.remove("last_project_path")
            self.statusBar().showMessage("上次打开的项目已移动或删除，请重新选择项目")
            return
        try:
            self.database = ProjectDatabase.open(path)
            self._set_project(str(path))
            self.statusBar().showMessage("已自动打开上次使用的项目")
        except Exception as exc:
            self._close_database()
            settings.remove("last_project_path")
            QMessageBox.warning(self, "无法恢复上次项目", f"上次项目无法打开，已清除记录：\n{exc}")

    def import_registrations(self) -> None:
        if not self.database:
            QMessageBox.information(self, "请先新建项目", "导入报名表前，请先新建或打开运动会项目。")
            return
        paths, _ = QFileDialog.getOpenFileNames(self, "选择班级报名表", "", "Excel 工作簿 (*.xlsx *.xlsm)")
        if not paths:
            return
        self._import_registration_paths(paths)

    def import_registration_folder(self) -> None:
        if not self.database:
            QMessageBox.information(self, "请先新建项目", "导入报名表前，请先新建或打开运动会项目。")
            return
        directory = QFileDialog.getExistingDirectory(self, "选择班级报名表文件夹")
        if not directory:
            return
        folder = Path(directory)
        paths = sorted(
            (
                path for path in folder.rglob("*")
                if path.is_file()
                and path.suffix.lower() in {".xlsx", ".xlsm"}
                and not path.name.startswith("~$")
            ),
            key=lambda path: str(path).lower(),
        )
        if not paths:
            QMessageBox.information(self, "没有报名表", "所选文件夹及其子文件夹中没有找到 .xlsx 或 .xlsm 文件。")
            return
        self._import_registration_paths(paths, source_root=folder)

    def _import_registration_paths(self, paths, source_root=None) -> None:
        self.statusBar().showMessage("正在读取报名表……")
        try:
            imported = import_registration_workbooks(paths, self.database.load_config())
            self.registration = imported
            self.latest_report = imported.report
            blocking = len(imported.report.by_severity(Severity.BLOCKING))
            warnings = len(imported.report.by_severity(Severity.WARNING))
            fixed = len(imported.report.by_severity(Severity.AUTO_FIXED))
            if not blocking:
                self.database.backup()
                self.database.replace_registration(
                    imported.athletes,
                    imported.entries,
                    imported.source_snapshots,
                    source_root,
                )
                self._refresh_result_events()
                self._refresh_final_material_events()
            self._show_registration(
                imported.athletes, imported.entries, blocking, warnings, fixed, len(paths)
            )
            self.statusBar().showMessage("报名表导入完成" if not blocking else "发现阻断问题，数据尚未写入项目")
            self.logger.info(
                "registration imported",
                fields={"row_count": len(imported.entries), "validation_count": len(imported.report.issues)},
            )
            if blocking:
                QMessageBox.warning(self, "存在阻断问题", "报名数据未写入项目。请先导出核对报告并处理红色问题。")
        except Exception as exc:
            self.statusBar().showMessage("报名表导入失败")
            QMessageBox.critical(self, "导入失败", str(exc))

    def _show_registration(
        self, athletes, entries, blocking: int, warnings: int, fixed: int,
        source_count: int | None = None,
    ) -> None:
        entries_by_athlete = {}
        for entry in entries:
            entries_by_athlete.setdefault(entry.athlete_id, []).append(entry.event_name)
        units = {(item.grade, item.class_name) for item in athletes}
        source_text = f"文件 {source_count} 个｜" if source_count is not None else ""
        summary = (
            f"{source_text}班级 {len(units)} 个｜运动员 {len(athletes)} 人｜"
            f"报名明细 {len(entries)} 条｜阻断 {blocking}｜警告 {warnings}｜自动修正 {fixed}"
        )
        self.summary_label.setText(summary)
        self.registration_summary_label.setText(summary)
        rows = [
            [
                athlete.bib,
                athlete.name,
                athlete.sex,
                athlete.grade,
                athlete.class_name,
                athlete.unit,
                "、".join(entries_by_athlete.get(athlete.id, [])),
            ]
            for athlete in athletes
        ]
        headers = ["号码", "姓名", "性别", "年级", "班级", "单位", "报名项目"]
        self.athlete_table.setModel(SimpleTableModel(headers, rows, self))
        self.registration_table.setModel(SimpleTableModel(headers, rows, self))
        self.athlete_table.resizeColumnsToContents()
        self.registration_table.resizeColumnsToContents()

    def import_schedule(self) -> None:
        if not self.database:
            QMessageBox.information(self, "请先新建项目", "导入比赛日程前，请先新建或打开运动会项目。")
            return
        path, _ = QFileDialog.getOpenFileName(self, "选择比赛日程", "", "Excel 工作簿 (*.xlsx *.xlsm)")
        if not path:
            return
        self.statusBar().showMessage("正在读取比赛日程……")
        try:
            imported = import_schedule_workbook(path, self.database.load_config())
            self.schedule = imported
            self.latest_report = imported.report
            blocking = len(imported.report.by_severity(Severity.BLOCKING))
            warnings = len(imported.report.by_severity(Severity.WARNING))
            fixed = len(imported.report.by_severity(Severity.AUTO_FIXED))
            if not blocking:
                self.database.backup()
                self.database.replace_schedule(imported.rounds)
                self._refresh_result_events()
                self._refresh_final_material_events()
            self._show_schedule(imported.rounds, blocking, warnings, fixed)
            self.statusBar().showMessage("比赛日程导入完成" if not blocking else "日程存在阻断问题，尚未写入项目")
            self.logger.info(
                "schedule imported",
                fields={"row_count": len(imported.rounds), "validation_count": len(imported.report.issues)},
            )
            if blocking:
                QMessageBox.warning(self, "日程存在阻断问题", "日程尚未写入项目。请导出核对报告查看具体行。")
        except Exception as exc:
            self.statusBar().showMessage("比赛日程导入失败")
            QMessageBox.critical(self, "导入失败", str(exc))

    def _registration_sources_are_current(self) -> bool:
        if not self.database:
            return False
        changes = self.database.registration_source_changes()
        if not changes:
            return True
        details = "\n".join(f"- {item}" for item in changes[:8])
        if len(changes) > 8:
            details += f"\n- 另有 {len(changes) - 8} 项变化"
        QMessageBox.warning(
            self,
            "报名表已发生变化",
            "上次导入后，报名表文件发生了变化。为避免日程和秩序册使用旧数据，"
            "请重新批量导入报名表后再继续。\n\n" + details,
        )
        return False

    def export_schedule_template(self) -> None:
        if not self.database:
            QMessageBox.information(self, "请先新建项目", "请先新建或打开运动会项目。")
            return
        if not self._registration_sources_are_current():
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "生成标准日程模板", "比赛日程标准模板.xlsx", "Excel 工作簿 (*.xlsx)"
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            config = self.database.load_config()
            athletes, entries = self.database.load_registration()
            plan_rows = build_standard_schedule(config, athletes, entries)
            create_schedule_template(path, plan_rows)
            count_text = "已按报名表中的参赛年级展开，并按实际报名人数计算组数和时长"
            QMessageBox.information(
                self,
                "标准日程已生成",
                f"共生成 {len(plan_rows)} 条比赛安排，{count_text}。\n"
                "时间、人数、组数和可用道次均可在 Excel 中继续修改。",
            )
        except ScheduleCapacityError as exc:
            QMessageBox.warning(self, "运动会时间不足", str(exc))
        except Exception as exc:
            QMessageBox.critical(self, "模板生成失败", str(exc))

    def _load_meeting_dates(self, config: ProjectConfig) -> None:
        if config.meeting_start_date and config.meeting_end_date:
            start = QDate.fromString(config.meeting_start_date, "yyyy-MM-dd")
            end = QDate.fromString(config.meeting_end_date, "yyyy-MM-dd")
            if start.isValid() and end.isValid():
                self.meeting_start_date_edit.setDate(start)
                self.meeting_end_date_edit.setDate(end)
                self.meeting_dates_label.setText(
                    f"已保存：{start.toString('yyyy年M月d日')} 至 {end.toString('yyyy年M月d日')}"
                )
                return
        self.meeting_dates_label.setText("当前项目尚未设置运动会日期，请选择后保存。")

    def save_meeting_dates(self) -> None:
        if not self.database:
            QMessageBox.information(self, "请先新建项目", "请先新建或打开运动会项目，再保存日期。")
            return
        start = self.meeting_start_date_edit.date()
        end = self.meeting_end_date_edit.date()
        if end < start:
            QMessageBox.warning(self, "日期范围错误", "运动会结束日期不能早于开始日期。")
            return
        try:
            config = self.database.load_config()
            config.meeting_start_date = start.toString("yyyy-MM-dd")
            config.meeting_end_date = end.toString("yyyy-MM-dd")
            self.database.save_config(config)
            self._load_meeting_dates(config)
            self.statusBar().showMessage("运动会日期已保存")
        except Exception as exc:
            QMessageBox.critical(self, "日期保存失败", str(exc))

    def _load_registration_template_settings(self, config: ProjectConfig) -> None:
        for field_name, table in self.registration_event_editors.items():
            formats = getattr(config, field_name.replace("_events", "_event_formats"))
            table.blockSignals(True)
            table.setRowCount(0)
            for event_name in getattr(config, field_name):
                self._add_registration_event_row(table, event_name, formats.get(event_name))
            table.blockSignals(False)
        self.team_event_table.blockSignals(True)
        self.team_event_table.setRowCount(0)
        for team_event in config.team_events:
            self._add_team_event_row(self.team_event_table, team_event)
        self.team_event_table.blockSignals(False)

    @staticmethod
    def _set_combo_data(combo: QComboBox, value: str) -> None:
        index = combo.findData(value)
        if index >= 0:
            combo.setCurrentIndex(index)

    @staticmethod
    def _guess_team_event_values(event_name: str) -> tuple[str, str, str]:
        compact = event_name.replace(" ", "")
        if "教工" in compact:
            return TEAM_GROUP_STAFF, TEAM_COUNT_MANUAL, ROUND_MODE_TIMED_FINAL
        if "拔河" in compact:
            return TEAM_GROUP_MIXED, TEAM_COUNT_MANUAL, TEAM_ROUND_MODE_KNOCKOUT
        if "男子" in compact:
            group_mode = TEAM_GROUP_MEN
        elif "女子" in compact:
            group_mode = TEAM_GROUP_WOMEN
        else:
            group_mode = TEAM_GROUP_MIXED
        return group_mode, TEAM_COUNT_ONE_PER_CLASS, ROUND_MODE_TIMED_FINAL

    def _add_team_event_row(
        self,
        table: QTableWidget,
        team_event: TeamEventConfig | None = None,
    ) -> None:
        row = table.rowCount()
        table.insertRow(row)
        event_name = team_event.event_name if team_event else ""
        guessed_group, guessed_count, guessed_round = self._guess_team_event_values(event_name)
        table.setItem(row, 0, QTableWidgetItem(event_name))

        group_combo = QComboBox()
        for label, value in (
            ("男子", TEAM_GROUP_MEN),
            ("女子", TEAM_GROUP_WOMEN),
            ("混合", TEAM_GROUP_MIXED),
            ("教工组", TEAM_GROUP_STAFF),
        ):
            group_combo.addItem(label, value)
        self._set_combo_data(group_combo, team_event.group_mode if team_event else guessed_group)
        group_combo.setProperty("manual_choice", team_event is not None)
        table.setCellWidget(row, 1, group_combo)

        scope_item = QTableWidgetItem()
        scope_item.setFlags(scope_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        table.setItem(row, 2, scope_item)

        count_combo = QComboBox()
        count_combo.addItem("每班1队", TEAM_COUNT_ONE_PER_CLASS)
        count_combo.addItem("手动填写", TEAM_COUNT_MANUAL)
        self._set_combo_data(count_combo, team_event.count_mode if team_event else guessed_count)
        count_combo.setProperty("manual_choice", team_event is not None)
        table.setCellWidget(row, 3, count_combo)

        manual_count = team_event.manual_team_count if team_event else None
        table.setItem(row, 4, QTableWidgetItem(str(manual_count) if manual_count else ""))

        round_combo = QComboBox()
        round_combo.addItem("预决赛（一次比完）", ROUND_MODE_TIMED_FINAL)
        round_combo.addItem("预赛＋决赛（分两次）", ROUND_MODE_PRELIMINARY_FINAL)
        round_combo.addItem("淘汰赛", TEAM_ROUND_MODE_KNOCKOUT)
        self._set_combo_data(round_combo, team_event.round_mode if team_event else guessed_round)
        round_combo.setProperty("manual_choice", team_event is not None)
        table.setCellWidget(row, 5, round_combo)

        for combo in (group_combo, count_combo, round_combo):
            combo.currentIndexChanged.connect(
                lambda _index, target=combo, target_table=table: (
                    target.setProperty("manual_choice", True),
                    self._refresh_team_event_rows(target_table),
                )
            )
        self._refresh_team_event_rows(table)
        if team_event is None:
            table.setCurrentCell(row, 0)
            table.editItem(table.item(row, 0))

    def _refresh_team_event_rows(self, table: QTableWidget) -> None:
        for row in range(table.rowCount()):
            group_combo = table.cellWidget(row, 1)
            count_combo = table.cellWidget(row, 3)
            if not isinstance(group_combo, QComboBox) or not isinstance(count_combo, QComboBox):
                continue
            is_staff = group_combo.currentData() == TEAM_GROUP_STAFF
            if is_staff and count_combo.currentData() != TEAM_COUNT_MANUAL:
                count_combo.blockSignals(True)
                self._set_combo_data(count_combo, TEAM_COUNT_MANUAL)
                count_combo.blockSignals(False)
            scope_item = table.item(row, 2)
            if scope_item:
                scope_item.setText("教工" if is_staff else "实际报名年级")
            manual_item = table.item(row, 4)
            if manual_item:
                editable = count_combo.currentData() == TEAM_COUNT_MANUAL
                flags = manual_item.flags()
                manual_item.setFlags(
                    flags | Qt.ItemFlag.ItemIsEditable
                    if editable
                    else flags & ~Qt.ItemFlag.ItemIsEditable
                )
                if not editable:
                    manual_item.setText("")

    def _on_team_event_name_changed(self, item: QTableWidgetItem) -> None:
        if item.column() != 0:
            return
        table = self.team_event_table
        group_mode, count_mode, round_mode = self._guess_team_event_values(item.text())
        for column, value in ((1, group_mode), (3, count_mode), (5, round_mode)):
            combo = table.cellWidget(item.row(), column)
            if not isinstance(combo, QComboBox) or combo.property("manual_choice"):
                continue
            combo.blockSignals(True)
            self._set_combo_data(combo, value)
            combo.blockSignals(False)
        self._refresh_team_event_rows(table)

    @staticmethod
    def _team_event_rows(table: QTableWidget) -> tuple[TeamEventConfig, ...]:
        result: list[TeamEventConfig] = []
        for row in range(table.rowCount()):
            name_item = table.item(row, 0)
            event_name = name_item.text().strip() if name_item else ""
            if not event_name:
                continue
            group_combo = table.cellWidget(row, 1)
            count_combo = table.cellWidget(row, 3)
            round_combo = table.cellWidget(row, 5)
            if not all(isinstance(combo, QComboBox) for combo in (group_combo, count_combo, round_combo)):
                raise ValueError(f"集体项目“{event_name}”的设置不完整")
            manual_text = table.item(row, 4).text().strip() if table.item(row, 4) else ""
            if manual_text and not manual_text.isdigit():
                raise ValueError(f"集体项目“{event_name}”的手动队数必须填写正整数或留空")
            result.append(
                TeamEventConfig(
                    event_name,
                    group_combo.currentData(),
                    count_combo.currentData(),
                    round_combo.currentData(),
                    int(manual_text) if manual_text else None,
                )
            )
        return tuple(result)

    @staticmethod
    def _round_mode_index(combo: QComboBox, round_mode: str) -> int:
        index = combo.findData(round_mode)
        return index if index >= 0 else 0

    def _add_registration_event_row(
        self,
        table: QTableWidget,
        event_name: str = "",
        round_mode: str | None = None,
    ) -> None:
        row = table.rowCount()
        table.insertRow(row)
        table.setItem(row, 0, QTableWidgetItem(event_name))
        combo = QComboBox()
        combo.addItem("预决赛（一次比完）", ROUND_MODE_TIMED_FINAL)
        combo.addItem("预赛＋决赛（分两次）", ROUND_MODE_PRELIMINARY_FINAL)
        selected_mode = round_mode or guess_event_round_mode(event_name)
        combo.setCurrentIndex(self._round_mode_index(combo, selected_mode))
        combo.setProperty("manual_round_mode", round_mode is not None)
        combo.currentIndexChanged.connect(
            lambda _index, target=combo: target.setProperty("manual_round_mode", True)
        )
        table.setCellWidget(row, 1, combo)
        if not event_name:
            table.setCurrentCell(row, 0)
            table.editItem(table.item(row, 0))

    def _on_registration_event_name_changed(
        self, table: QTableWidget, item: QTableWidgetItem
    ) -> None:
        if item.column() != 0:
            return
        combo = table.cellWidget(item.row(), 1)
        if not isinstance(combo, QComboBox) or combo.property("manual_round_mode"):
            return
        combo.blockSignals(True)
        combo.setCurrentIndex(
            self._round_mode_index(combo, guess_event_round_mode(item.text().strip()))
        )
        combo.blockSignals(False)

    def _batch_add_registration_events(self, table: QTableWidget) -> None:
        text, accepted = QInputDialog.getMultiLineText(
            self,
            "批量添加项目",
            "每行一个项目，也可用逗号或顿号分隔：",
        )
        if not accepted:
            return
        for event_name in self._parse_event_list(text):
            self._add_registration_event_row(table, event_name)

    @staticmethod
    def _delete_selected_registration_events(table: QTableWidget) -> None:
        rows = sorted({index.row() for index in table.selectedIndexes()}, reverse=True)
        for row in rows:
            table.removeRow(row)

    @staticmethod
    def _registration_event_rows(table: QTableWidget) -> tuple[tuple[str, str], ...]:
        rows: list[tuple[str, str]] = []
        for row in range(table.rowCount()):
            item = table.item(row, 0)
            event_name = item.text().strip() if item else ""
            if not event_name:
                continue
            combo = table.cellWidget(row, 1)
            round_mode = (
                combo.currentData()
                if isinstance(combo, QComboBox)
                else guess_event_round_mode(event_name)
            )
            rows.append((event_name, round_mode))
        return tuple(rows)

    @staticmethod
    def _parse_event_list(text: str) -> tuple[str, ...]:
        return tuple(
            item.strip()
            for item in re.split(r"[\n,，、;；]+", text)
            if item.strip()
        )

    def _save_registration_template_settings(self, show_confirmation: bool) -> ProjectConfig | None:
        if not self.database:
            QMessageBox.information(self, "请先新建项目", "请先新建或打开运动会项目，再保存报名表设置。")
            return None
        try:
            config = self.database.load_config()
            for field_name, table in self.registration_event_editors.items():
                rows = self._registration_event_rows(table)
                setattr(config, field_name, tuple(event_name for event_name, _mode in rows))
                setattr(
                    config,
                    field_name.replace("_events", "_event_formats"),
                    {event_name: mode for event_name, mode in rows},
                )
            config.team_events = self._team_event_rows(self.team_event_table)
            config.validate()
            self.database.save_config(config)
            self._load_registration_template_settings(config)
            self.statusBar().showMessage("报名项目设置已保存")
            if show_confirmation:
                QMessageBox.information(self, "保存成功", "个人项目和集体项目设置已保存。")
            return config
        except Exception as exc:
            QMessageBox.warning(self, "设置无法保存", str(exc))
            return None

    def save_registration_template_settings(self) -> None:
        self._save_registration_template_settings(show_confirmation=True)

    def export_level_registration_template(self, school_level: str) -> None:
        config = self._save_registration_template_settings(show_confirmation=False)
        if config is None:
            return
        default_name = f"{school_level}班级报名表模板.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self, f"导出{school_level}报名表模板", default_name, "Excel 工作簿 (*.xlsx)"
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        prefix = "junior" if school_level == "初中" else "senior"
        try:
            create_school_level_registration_template(
                path,
                school_level,
                getattr(config, f"{prefix}_boys_events"),
                getattr(config, f"{prefix}_girls_events"),
            )
            self.statusBar().showMessage(f"{school_level}报名表模板已生成")
            QMessageBox.information(
                self,
                "模板已生成",
                f"已生成{school_level}模板，包含男生、女生两个工作表及各自的项目下拉选项。",
            )
        except Exception as exc:
            QMessageBox.critical(self, "模板生成失败", str(exc))

    def _show_schedule(self, rounds, blocking: int, warnings: int, fixed: int) -> None:
        self.schedule_summary_label.setText(
            f"日程 {len(rounds)} 项｜阻断 {blocking}｜警告 {warnings}｜自动建议 {fixed}"
        )
        type_names = {"TRACK": "径赛", "FIELD": "田赛", "TEAM": "集体项目"}
        round_names = {"PRELIMINARY": "预赛", "FINAL": "决赛", "TIMED_FINAL": "预决赛"}
        rows = [
            [
                item.scheduled_time, type_names[item.event_type.value], item.group_name,
                item.event_name, round_names[item.round_type.value], item.declared_count,
                item.heat_count or "待确认",
            ]
            for item in rounds
        ]
        self.schedule_table.setModel(
            SimpleTableModel(["比赛时间", "类型", "组别", "项目", "赛次", "人数/队数", "组数"], rows, self)
        )
        self.schedule_table.resizeColumnsToContents()

    def run_project_validation(self) -> None:
        if not self.database:
            QMessageBox.information(self, "请先新建项目", "校验前，请先新建或打开运动会项目。")
            return
        try:
            athletes, entries = self.database.load_registration()
            rounds = self.database.load_schedule()
            report = validate_project(self.database.load_config(), athletes, entries, rounds)
            self.latest_report = report
            blocking = len(report.by_severity(Severity.BLOCKING))
            warnings = len(report.by_severity(Severity.WARNING))
            fixed = len(report.by_severity(Severity.AUTO_FIXED))
            self.validation_summary_label.setText(
                f"阻断 {blocking}｜警告 {warnings}｜自动修正 {fixed}｜共 {len(report.issues)} 条"
            )
            severity_names = {
                Severity.BLOCKING: "阻断", Severity.WARNING: "警告", Severity.AUTO_FIXED: "自动修正"
            }
            rows = [
                [issue.code, severity_names[issue.severity], issue.message, issue.entity_type, issue.entity_id]
                for issue in report.issues
            ]
            self.validation_table.setModel(
                SimpleTableModel(["规则代码", "级别", "说明", "对象类型", "对象ID"], rows, self)
            )
            self.validation_table.resizeColumnsToContents()
            self.statusBar().showMessage("全项目校验完成")
        except Exception as exc:
            QMessageBox.critical(self, "校验失败", str(exc))

    def generate_heats(self) -> None:
        if not self.database:
            QMessageBox.information(self, "请先新建项目", "编排前，请先新建或打开运动会项目。")
            return
        if not self._registration_sources_are_current():
            return
        try:
            athletes, entries = self.database.load_registration()
            rounds = self.database.load_schedule()
            validation = validate_project(self.database.load_config(), athletes, entries, rounds)
            if validation.has_blocking:
                self.latest_report = validation
                QMessageBox.warning(
                    self, "暂时不能编排",
                    f"全项目校验仍有 {len(validation.by_severity(Severity.BLOCKING))} 条阻断问题，请先到“核对修正”处理。",
                )
                return
            existing = self.database.load_heat_assignments()
            if existing:
                answer = QMessageBox.question(
                    self, "确认重新抽签", "重新抽签会替换当前全部分组，但会先自动备份项目。是否继续？"
                )
                if answer != QMessageBox.StandardButton.Yes:
                    return
            generated = generate_heat_assignments(
                self.database.load_config(), athletes, entries, rounds
            )
            if generated.report.has_blocking:
                self.latest_report = generated.report
                QMessageBox.warning(self, "编排未完成", "部分项目无法编排，请导出当前核对报告查看详情。")
                return
            self.database.backup()
            self.database.replace_heat_assignments(generated.assignments, generated.seed)
            self.latest_report = generated.report
            self._show_assignments(generated.assignments)
            self._refresh_result_events()
            self._refresh_final_material_events()
            warning_count = len(generated.report.by_severity(Severity.WARNING))
            self.statusBar().showMessage(f"分组已生成；集体项目提示 {warning_count} 条")
        except Exception as exc:
            QMessageBox.critical(self, "编排失败", str(exc))

    def _show_assignments(self, assignments) -> None:
        if not self.database:
            return
        athletes, _ = self.database.load_registration()
        rounds = self.database.load_schedule()
        athlete_map = {item.id: item for item in athletes}
        round_map = {item.id: item for item in rounds}
        ordered = sorted(
            assignments,
            key=lambda item: (
                round_map[item.event_round_id].scheduled_time,
                round_map[item.event_round_id].group_name,
                round_map[item.event_round_id].event_name,
                item.heat_no,
                item.lane if item.lane is not None else item.order or 0,
            ),
        )
        rows = []
        for item in ordered:
            round_ = round_map[item.event_round_id]
            athlete = athlete_map[item.participant_id]
            rows.append([
                round_.scheduled_time, round_.group_name, round_.event_name, item.heat_no,
                item.lane if item.lane is not None else item.order,
                "道次" if item.lane is not None else "出场序", athlete.bib, athlete.name, athlete.unit,
            ])
        seed = next(
            (item.random_seed for item in assignments if item.random_seed is not None), ""
        )
        self.grouping_summary_label.setText(
            f"编排 {len(assignments)} 人次｜随机种子 {seed}｜结果已保存到项目文件"
        )
        self.grouping_table.setModel(
            SimpleTableModel(
                ["比赛时间", "组别", "项目", "第几组", "序号", "含义", "号码", "姓名", "单位"], rows, self
            )
        )
        self.grouping_table.resizeColumnsToContents()

    def _refresh_result_events(self) -> None:
        self.result_event_combo.blockSignals(True)
        self.result_event_combo.clear()
        if self.database:
            assigned_round_ids = {
                item.event_round_id for item in self.database.load_heat_assignments()
            }
            for round_ in self.database.load_schedule():
                if (
                    round_.round_type is RoundType.PRELIMINARY
                    and round_.id in assigned_round_ids
                ):
                    self.result_event_combo.addItem(
                        f"{round_.group_name} {round_.event_name} 预赛（{round_.scheduled_time}）",
                        round_.id,
                    )
        self.result_event_combo.blockSignals(False)
        self.load_result_event()

    def load_result_event(self, _index: int = 0) -> None:
        self.result_rows = []
        self.manual_rank_by_participant = {}
        self.manual_rank_button.setChecked(False)
        self.result_table.setRowCount(0)
        self.final_preview_table.setModel(SimpleTableModel([], [], self))
        if not self.database:
            return
        preliminary_id = self.result_event_combo.currentData()
        if not preliminary_id:
            return
        athletes, _ = self.database.load_registration()
        preliminary = next(
            item for item in self.database.load_schedule() if item.id == preliminary_id
        )
        saved_unit = self.database.load_result_input_unit(
            preliminary_id,
            recommended_input_unit(preliminary.event_name, preliminary.performance_kind),
        )
        unit_index = self.result_unit_combo.findData(saved_unit)
        self.result_unit_combo.setCurrentIndex(unit_index if unit_index >= 0 else 1)
        athlete_map = {item.id: item for item in athletes}
        assignments = sorted(
            (
                item for item in self.database.load_heat_assignments()
                if item.event_round_id == preliminary_id
            ),
            key=lambda item: (item.heat_no, item.lane or item.order or 0),
        )
        existing_results = {
            item.participant_id: item for item in self.database.load_results(preliminary_id)
        }
        status_names = [
            ("有效", ResultStatus.VALID),
            ("未起跑", ResultStatus.DNS),
            ("未完成", ResultStatus.DNF),
            ("犯规", ResultStatus.DQ),
            ("无成绩", ResultStatus.NM),
        ]
        self.result_table.setRowCount(len(assignments))
        self.result_rows = assignments
        for row_index, assignment in enumerate(assignments):
            athlete = athlete_map[assignment.participant_id]
            values = [
                assignment.heat_no,
                assignment.lane if assignment.lane is not None else assignment.order,
                athlete.bib,
                athlete.name,
                athlete.unit,
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.result_table.setItem(row_index, column, item)
            existing = existing_results.get(assignment.participant_id)
            self.result_table.setItem(
                row_index, ScoreTableWidget.SCORE_COLUMN,
                QTableWidgetItem(existing.raw_value if existing else ""),
            )
            status_combo = QComboBox()
            for name, status in status_names:
                status_combo.addItem(name, status.value)
            if existing:
                status_combo.setCurrentIndex(status_combo.findData(existing.status.value))
            self.result_table.setCellWidget(row_index, 6, status_combo)
            rank_item = QTableWidgetItem()
            rank_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            rank_item.setFlags(rank_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            if existing and existing.manual_rank is not None:
                self.manual_rank_by_participant[assignment.participant_id] = existing.manual_rank
                rank_item.setText(str(existing.manual_rank))
            self.result_table.setItem(row_index, 7, rank_item)
        self.result_table.resizeColumnsToContents()
        self._show_final_preview(preliminary_id)

    def toggle_manual_ranking(self, enabled: bool) -> None:
        self.manual_rank_button.setText("结束手动排名" if enabled else "手动决定名次")
        self.result_unit_combo.setEnabled(not enabled)
        for row in range(self.result_table.rowCount()):
            item = self.result_table.item(row, ScoreTableWidget.SCORE_COLUMN)
            if item is None:
                continue
            if enabled:
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            else:
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
        if enabled:
            self.statusBar().showMessage("手动排名已开启：点击运动员所在行，按各组分别赋予第1、2、3名……")

    def assign_manual_rank(self, row: int, _column: int) -> None:
        if not self.manual_rank_button.isChecked() or not 0 <= row < len(self.result_rows):
            return
        assignment = self.result_rows[row]
        if assignment.participant_id in self.manual_rank_by_participant:
            self.statusBar().showMessage("该运动员已有组内名次；如需重排，请点击“清除手动名次”")
            return
        used = [
            rank
            for item, rank in (
                (candidate, self.manual_rank_by_participant.get(candidate.participant_id))
                for candidate in self.result_rows
            )
            if item.heat_no == assignment.heat_no and rank is not None
        ]
        rank = max(used, default=0) + 1
        self.manual_rank_by_participant[assignment.participant_id] = rank
        self.result_table.item(row, 7).setText(str(rank))
        self.result_table.item(row, ScoreTableWidget.SCORE_COLUMN).setText("")
        status_widget = self.result_table.cellWidget(row, 6)
        status_widget.setCurrentIndex(status_widget.findData(ResultStatus.VALID.value))
        self.statusBar().showMessage(f"已指定第 {assignment.heat_no} 组第 {rank} 名")

    def clear_manual_ranks(self) -> None:
        self.manual_rank_button.setChecked(False)
        self.manual_rank_by_participant.clear()
        for row in range(self.result_table.rowCount()):
            rank_item = self.result_table.item(row, 7)
            if rank_item:
                rank_item.setText("")
            status_widget = self.result_table.cellWidget(row, 6)
            if status_widget:
                status_widget.setCurrentIndex(status_widget.findData(ResultStatus.VALID.value))
        self.statusBar().showMessage("手动名次已清除，可以重新点击排名或录入成绩")

    def save_results_and_generate_final(self) -> None:
        if not self.database:
            QMessageBox.information(self, "请先新建项目", "录入成绩前，请先新建或打开运动会项目。")
            return
        preliminary_id = self.result_event_combo.currentData()
        rounds = self.database.load_schedule()
        preliminary = next((item for item in rounds if item.id == preliminary_id), None)
        if preliminary is None:
            QMessageBox.information(self, "没有可录入项目", "请先完成径赛预赛编排。")
            return
        final_round = next(
            (
                item for item in rounds
                if item.group_name == preliminary.group_name
                and item.event_name == preliminary.event_name
                and item.round_type is RoundType.FINAL
            ),
            None,
        )
        if final_round is None:
            QMessageBox.warning(
                self, "缺少决赛日程",
                f"日程中没有“{preliminary.group_name} {preliminary.event_name}”的决赛场次，请先补充日程。",
            )
            return

        results = []
        errors = []
        status_text = {
            ResultStatus.DNS: "DNS",
            ResultStatus.DNF: "DNF",
            ResultStatus.DQ: "DQ",
            ResultStatus.NM: "NM",
        }
        manual_mode = bool(self.manual_rank_by_participant)
        selected_unit = self.result_unit_combo.currentData()
        unit_kind = {
            "MINUTE": PerformanceKind.TIME,
            "SECOND": PerformanceKind.TIME,
            "METER": PerformanceKind.DISTANCE,
            "COUNT": PerformanceKind.COUNT,
        }
        performance_kind = unit_kind[selected_unit]
        default_unit = "m" if selected_unit == "METER" else "cm"
        for row_index, assignment in enumerate(self.result_rows):
            score_item = self.result_table.item(row_index, ScoreTableWidget.SCORE_COLUMN)
            raw_value = score_item.text().strip() if score_item else ""
            status_widget = self.result_table.cellWidget(row_index, 6)
            status = ResultStatus(status_widget.currentData())
            bib = self.result_table.item(row_index, 2).text()
            manual_rank = self.manual_rank_by_participant.get(assignment.participant_id)
            if not manual_mode and status is ResultStatus.VALID and not raw_value:
                errors.append(f"第 {row_index + 1} 行（号码 {bib}）：有效成绩不能为空")
                continue
            try:
                if manual_mode:
                    if status is ResultStatus.VALID and manual_rank is not None:
                        results.append(
                            Result(
                                assignment.participant_id, assignment.heat_no, "",
                                PerformanceKind.MANUAL_RANK, None, None,
                                f"第{manual_rank}名", ResultStatus.VALID, manual_rank,
                            )
                        )
                    else:
                        value = status_text.get(status, "NM")
                        results.append(
                            parse_result(
                                assignment.participant_id, assignment.heat_no, value,
                                PerformanceKind.MANUAL_RANK,
                            )
                        )
                else:
                    value = raw_value if status is ResultStatus.VALID else status_text[status]
                    results.append(
                        parse_result(
                            assignment.participant_id,
                            assignment.heat_no,
                            value,
                            performance_kind,
                            default_unit=default_unit,
                            display_unit=selected_unit,
                        )
                    )
            except PerformanceParseError as exc:
                errors.append(f"第 {row_index + 1} 行（号码 {bib}）：{exc}")
        if errors:
            details = "\n".join(errors[:12])
            if len(errors) > 12:
                details += f"\n……另有 {len(errors) - 12} 条"
            QMessageBox.warning(self, "成绩未保存", details)
            return

        try:
            computed = compute_track_final(
                preliminary, final_round, results, self.database.load_config()
            )
            self.database.backup()
            self.database.save_track_final(
                preliminary.id,
                final_round.id,
                results,
                computed.qualifications,
                computed.assignments,
            )
            self.database.save_result_input_unit(preliminary.id, selected_unit)
            self._show_assignments(self.database.load_heat_assignments())
            self._show_final_preview(preliminary.id)
            self._refresh_final_material_events()
            self.statusBar().showMessage("成绩已保存，决赛名单和道次已生成")
            QMessageBox.information(
                self,
                "生成完成",
                f"已保存 {len(results)} 条成绩，生成 {len(computed.qualifications)} 名晋级运动员、"
                f"{len(computed.assignments)} 条决赛道次。",
            )
        except Exception as exc:
            QMessageBox.critical(self, "决赛编排失败", str(exc))

    def _show_final_preview(self, preliminary_id: str) -> None:
        if not self.database:
            return
        rounds = self.database.load_schedule()
        preliminary = next((item for item in rounds if item.id == preliminary_id), None)
        if preliminary is None:
            return
        final_round = next(
            (
                item for item in rounds
                if item.group_name == preliminary.group_name
                and item.event_name == preliminary.event_name
                and item.round_type is RoundType.FINAL
            ),
            None,
        )
        if final_round is None:
            self.final_preview_table.setModel(SimpleTableModel([], [], self))
            return
        athletes, _ = self.database.load_registration()
        athlete_map = {item.id: item for item in athletes}
        qualification_map = {
            item.participant_id: item
            for item in self.database.load_qualifications(preliminary_id)
        }
        result_map = {
            item.participant_id: item for item in self.database.load_results(preliminary_id)
        }
        assignments = sorted(
            (
                item for item in self.database.load_heat_assignments()
                if item.event_round_id == final_round.id
            ),
            key=lambda item: (item.heat_no, item.lane or 0),
        )
        rows = []
        for assignment in assignments:
            athlete = athlete_map[assignment.participant_id]
            qualification = qualification_map[assignment.participant_id]
            result = result_map.get(assignment.participant_id)
            rows.append([
                assignment.heat_no,
                assignment.lane if assignment.lane is not None else assignment.order,
                athlete.bib,
                athlete.name,
                athlete.unit,
                result.standard_display if result else format_canonical(
                    preliminary.performance_kind, qualification.canonical_value
                ),
                qualification.overall_rank,
                qualification.reason,
            ])
        self.final_preview_table.setModel(
            SimpleTableModel(
                ["决赛组", "道次/出场序", "号码", "姓名", "单位", "预赛成绩", "总名次", "晋级原因"],
                rows,
                self,
            )
        )
        self.final_preview_table.resizeColumnsToContents()

    def _final_rounds_with_assignments(self):
        if not self.database:
            return []
        assigned_round_ids = {
            item.event_round_id for item in self.database.load_heat_assignments()
        }
        return [
            round_
            for round_ in self.database.load_schedule()
            if round_.round_type is RoundType.FINAL and round_.id in assigned_round_ids
        ]

    def _refresh_final_material_events(self) -> None:
        self.final_material_event_combo.blockSignals(True)
        self.final_material_event_combo.clear()
        for round_ in self._final_rounds_with_assignments():
            self.final_material_event_combo.addItem(
                f"{round_.group_name} {round_.event_name} 决赛（{round_.scheduled_time}）",
                round_.id,
            )
        self.final_material_event_combo.blockSignals(False)
        self.load_final_material_event()

    def load_final_material_event(self, _index: int = 0) -> None:
        self.final_result_rows = []
        self.final_material_table.setRowCount(0)
        final_round_id = self.final_material_event_combo.currentData()
        if not self.database or not final_round_id:
            self.final_material_summary_label.setText("尚未生成决赛名单")
            return
        final_round = next(
            (item for item in self.database.load_schedule() if item.id == final_round_id),
            None,
        )
        if final_round is None:
            self.final_material_summary_label.setText("所选决赛项目已不存在")
            return
        athletes, _ = self.database.load_registration()
        athlete_map = {item.id: item for item in athletes}
        saved_unit = self.database.load_result_input_unit(
            final_round_id,
            recommended_input_unit(final_round.event_name, final_round.performance_kind),
        )
        unit_index = self.final_result_unit_combo.findData(saved_unit)
        self.final_result_unit_combo.setCurrentIndex(unit_index if unit_index >= 0 else 1)
        assignments = sorted(
            (
                item for item in self.database.load_heat_assignments()
                if item.event_round_id == final_round_id
            ),
            key=lambda item: (
                item.heat_no,
                item.lane if item.lane is not None else item.order or 0,
            ),
        )
        position_name = "出场序" if final_round.event_type is EventType.FIELD else "道次"
        existing_results = {
            item.participant_id: item for item in self.database.load_results(final_round_id)
        }
        ranks = {
            result.participant_id: rank
            for rank, result in rank_final_results(list(existing_results.values()))
            if rank is not None
        }
        status_names = [
            ("有效", ResultStatus.VALID),
            ("未起跑", ResultStatus.DNS),
            ("未完成", ResultStatus.DNF),
            ("犯规", ResultStatus.DQ),
            ("无成绩", ResultStatus.NM),
        ]
        self.final_result_rows = assignments
        self.final_material_table.setRowCount(len(assignments))
        for row_index, assignment in enumerate(assignments):
            athlete = athlete_map[assignment.participant_id]
            values = [
                assignment.heat_no,
                assignment.lane if assignment.lane is not None else assignment.order,
                athlete.bib,
                athlete.name,
                athlete.unit,
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.final_material_table.setItem(row_index, column, item)
            existing = existing_results.get(assignment.participant_id)
            self.final_material_table.setItem(
                row_index,
                ScoreTableWidget.SCORE_COLUMN,
                QTableWidgetItem(existing.raw_value if existing else ""),
            )
            status_combo = QComboBox()
            for name, status in status_names:
                status_combo.addItem(name, status.value)
            if existing:
                status_combo.setCurrentIndex(status_combo.findData(existing.status.value))
            self.final_material_table.setCellWidget(row_index, 6, status_combo)
            rank_item = QTableWidgetItem(
                str(ranks.get(assignment.participant_id, ""))
            )
            rank_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            rank_item.setFlags(rank_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.final_material_table.setItem(row_index, 7, rank_item)
        self.final_material_table.resizeColumnsToContents()
        score_count = len(existing_results)
        self.final_material_summary_label.setText(
            f"{final_round.group_name} {final_round.event_name}｜名单 {len(assignments)} 人｜"
            f"已录成绩 {score_count} 人｜{position_name}已确定"
        )

    def _selected_final_round(self):
        final_round_id = self.final_material_event_combo.currentData()
        return next(
            (item for item in self._final_rounds_with_assignments() if item.id == final_round_id),
            None,
        )

    def save_final_results(self) -> None:
        if not self.database:
            QMessageBox.information(self, "请先打开项目", "请先新建或打开运动会项目。")
            return
        final_round = self._selected_final_round()
        if final_round is None:
            QMessageBox.information(self, "没有决赛名单", "请先在成绩录入页生成决赛名单。")
            return
        selected_unit = self.final_result_unit_combo.currentData()
        performance_kind = {
            "MINUTE": PerformanceKind.TIME,
            "SECOND": PerformanceKind.TIME,
            "METER": PerformanceKind.DISTANCE,
            "COUNT": PerformanceKind.COUNT,
        }[selected_unit]
        default_unit = "m" if selected_unit == "METER" else "cm"
        status_text = {
            ResultStatus.DNS: "DNS",
            ResultStatus.DNF: "DNF",
            ResultStatus.DQ: "DQ",
            ResultStatus.NM: "NM",
        }
        results = []
        errors = []
        for row_index, assignment in enumerate(self.final_result_rows):
            score_item = self.final_material_table.item(
                row_index, ScoreTableWidget.SCORE_COLUMN
            )
            raw_value = score_item.text().strip() if score_item else ""
            status_widget = self.final_material_table.cellWidget(row_index, 6)
            status = ResultStatus(status_widget.currentData())
            bib = self.final_material_table.item(row_index, 2).text()
            if status is ResultStatus.VALID and not raw_value:
                errors.append(f"第 {row_index + 1} 行（号码 {bib}）：有效成绩不能为空")
                continue
            try:
                value = raw_value if status is ResultStatus.VALID else status_text[status]
                results.append(parse_result(
                    assignment.participant_id,
                    assignment.heat_no,
                    value,
                    performance_kind,
                    default_unit=default_unit,
                    display_unit=selected_unit,
                ))
            except PerformanceParseError as exc:
                errors.append(f"第 {row_index + 1} 行（号码 {bib}）：{exc}")
        if errors:
            details = "\n".join(errors[:12])
            if len(errors) > 12:
                details += f"\n……另有 {len(errors) - 12} 条"
            QMessageBox.warning(self, "决赛成绩未保存", details)
            return
        try:
            rank_final_results(results)
            self.database.backup()
            self.database.save_round_results(final_round.id, results)
            self.database.save_result_input_unit(final_round.id, selected_unit)
            self.load_final_material_event()
            self.statusBar().showMessage("决赛成绩已保存，合并名次已计算")
            QMessageBox.information(
                self, "保存完成", f"已保存 {len(results)} 条决赛成绩并计算名次。"
            )
        except Exception as exc:
            QMessageBox.critical(self, "决赛成绩保存失败", str(exc))

    def _participants_for_documents(self) -> dict[str, Participant]:
        if not self.database:
            return {}
        athletes, _ = self.database.load_registration()
        return {
            item.id: Participant(item.id, item.bib, item.name, item.unit)
            for item in athletes
        }

    def _material_for_round(self, final_round) -> FinalMaterial:
        assignments = [
            item for item in self.database.load_heat_assignments()
            if item.event_round_id == final_round.id
        ]
        heat_numbers = sorted({item.heat_no for item in assignments})
        heats = []
        for heat_no in heat_numbers:
            heat = [
                (
                    item.participant_id,
                    item.lane if item.lane is not None else item.order or 0,
                )
                for item in assignments
                if item.heat_no == heat_no
            ]
            heats.append(sorted(heat, key=lambda item: item[1]))
        return FinalMaterial(
            event_title=f"{final_round.group_name} {final_round.event_name}",
            scheduled_time=final_round.scheduled_time,
            heats=heats,
            position_header=(
                "出场序" if final_round.event_type is EventType.FIELD else "道次"
            ),
        )

    def export_current_final_material(self) -> None:
        final_round_id = self.final_material_event_combo.currentData()
        final_round = next(
            (item for item in self._final_rounds_with_assignments() if item.id == final_round_id),
            None,
        )
        if final_round is None:
            QMessageBox.information(self, "没有决赛名单", "请先在成绩录入页生成决赛名单。")
            return
        self._export_final_materials(
            [self._material_for_round(final_round)],
            f"{final_round.group_name}{final_round.event_name}决赛名单.docx",
        )

    def export_all_final_materials(self) -> None:
        rounds = self._final_rounds_with_assignments()
        if not rounds:
            QMessageBox.information(self, "没有决赛名单", "请先在成绩录入页生成决赛名单。")
            return
        self._export_final_materials(
            [self._material_for_round(item) for item in rounds],
            "全部决赛名单.docx",
        )

    def export_current_on_site_package(self) -> None:
        final_round = self._selected_final_round()
        if final_round is None:
            QMessageBox.information(self, "没有决赛名单", "请先在成绩录入页生成决赛名单。")
            return
        self._export_on_site_packages(
            [self._material_for_round(final_round)],
            f"{final_round.group_name}{final_round.event_name}决赛现场材料",
        )

    def export_all_on_site_packages(self) -> None:
        rounds = self._final_rounds_with_assignments()
        if not rounds:
            QMessageBox.information(self, "没有决赛名单", "请先在成绩录入页生成决赛名单。")
            return
        self._export_on_site_packages(
            [self._material_for_round(item) for item in rounds],
            "全部决赛现场材料",
        )

    def _export_on_site_packages(
        self, materials: list[FinalMaterial], folder_name: str
    ) -> None:
        parent = QFileDialog.getExistingDirectory(self, "选择现场材料包保存位置")
        if not parent:
            return
        base = Path(parent) / folder_name
        target = base
        copy_number = 2
        while target.exists():
            target = base.with_name(f"{base.name}-{copy_number}")
            copy_number += 1
        try:
            self.statusBar().showMessage("正在生成决赛现场材料包……")
            QApplication.processEvents()
            files = render_on_site_package(
                materials, self._participants_for_documents(), target
            )
            self.statusBar().showMessage("决赛现场材料包生成完成")
            file_names = "\n".join(
                f"- {path.name}"
                for path in (
                    files.assignment_sheet,
                    files.check_in_sheet,
                    files.official_sheet,
                    files.broadcast_script,
                )
            )
            QMessageBox.information(
                self,
                "生成完成",
                f"现场材料包已生成：\n{target}\n\n包含：\n{file_names}",
            )
        except Exception as exc:
            self.statusBar().showMessage("决赛现场材料包生成失败")
            QMessageBox.critical(self, "生成失败", str(exc))

    def _ranking_material_for_round(self, final_round) -> FinalRankingMaterial | None:
        results = self.database.load_results(final_round.id)
        if not results:
            return None
        participants = self._participants_for_documents()
        ranked = [
            (rank, participants[result.participant_id], result)
            for rank, result in rank_final_results(results, maximum_rank=8)
        ]
        return FinalRankingMaterial(
            f"{final_round.group_name} {final_round.event_name}", ranked
        )

    def export_current_final_ranking(self) -> None:
        final_round = self._selected_final_round()
        material = (
            self._ranking_material_for_round(final_round) if final_round is not None else None
        )
        if material is None:
            QMessageBox.information(self, "没有决赛成绩", "请先保存当前项目的决赛成绩。")
            return
        self._export_final_rankings(
            [material], f"{final_round.group_name}{final_round.event_name}成绩排名表.docx"
        )

    def export_all_final_rankings(self) -> None:
        materials = [
            material
            for round_ in self._final_rounds_with_assignments()
            if (material := self._ranking_material_for_round(round_)) is not None
        ]
        if not materials:
            QMessageBox.information(self, "没有决赛成绩", "请先保存至少一个项目的决赛成绩。")
            return
        self._export_final_rankings(materials, "全部决赛成绩排名表.docx")

    def _export_final_rankings(
        self, materials: list[FinalRankingMaterial], default_name: str
    ) -> None:
        path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "生成决赛成绩排名表",
            default_name,
            "Word 文档 (*.docx);;PDF 文件 (*.pdf)",
        )
        if not path:
            return
        suffix = ".pdf" if "PDF" in selected_filter else ".docx"
        if Path(path).suffix.lower() != suffix:
            path = str(Path(path).with_suffix(suffix))
        try:
            self.statusBar().showMessage("正在生成决赛成绩排名表……")
            QApplication.processEvents()
            if suffix == ".pdf":
                with tempfile.TemporaryDirectory(prefix="sport-ranking-") as temp_directory:
                    temporary_docx = Path(temp_directory) / "final-rankings.docx"
                    render_result_rankings_batch(materials, temporary_docx)
                    QApplication.processEvents()
                    self.export_manager.export_pdf(temporary_docx, path)
            else:
                render_result_rankings_batch(materials, path)
            self.statusBar().showMessage("决赛成绩排名表生成完成")
            QMessageBox.information(self, "生成完成", f"成绩排名表已生成：\n{path}")
        except Exception as exc:
            self.statusBar().showMessage("决赛成绩排名表生成失败")
            QMessageBox.critical(self, "生成失败", str(exc))

    def _export_final_materials(
        self, materials: list[FinalMaterial], default_name: str
    ) -> None:
        path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "生成决赛材料",
            default_name,
            "Word 文档 (*.docx);;PDF 文件 (*.pdf)",
        )
        if not path:
            return
        suffix = ".pdf" if "PDF" in selected_filter else ".docx"
        if Path(path).suffix.lower() != suffix:
            path = str(Path(path).with_suffix(suffix))
        try:
            participants = self._participants_for_documents()
            self.statusBar().showMessage("正在生成决赛材料……")
            QApplication.processEvents()
            if suffix == ".pdf":
                with tempfile.TemporaryDirectory(prefix="sport-final-") as temp_directory:
                    temporary_docx = Path(temp_directory) / "final-materials.docx"
                    render_final_assignments_batch(materials, participants, temporary_docx)
                    QApplication.processEvents()
                    self.export_manager.export_pdf(temporary_docx, path)
            else:
                render_final_assignments_batch(materials, participants, path)
            self.statusBar().showMessage("决赛材料生成完成")
            QMessageBox.information(self, "生成完成", f"决赛材料已生成：\n{path}")
        except Exception as exc:
            self.statusBar().showMessage("决赛材料生成失败")
            QMessageBox.critical(self, "生成失败", str(exc))

    def generate_competition_document(self) -> None:
        if not self.database:
            QMessageBox.information(self, "请先新建项目", "生成文档前，请先新建或打开运动会项目。")
            return
        if not self._registration_sources_are_current():
            return
        assignments = self.database.load_heat_assignments()
        if not assignments:
            QMessageBox.information(self, "尚未编排", "请先到“抽签编排”生成并确认分组。")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "生成竞赛分组表", "竞赛分组表.docx", "Word 文档 (*.docx)"
        )
        if not path:
            return
        if not path.lower().endswith(".docx"):
            path += ".docx"
        try:
            athletes, _ = self.database.load_registration()
            stats = render_competition_groups(
                path, self.database.load_schedule(), assignments, athletes
            )
            self.booklet_summary_label.setText(
                f"已生成 {stats.event_count} 个项目、{stats.table_count} 张分组表、{stats.assignment_count} 人次。"
            )
            self._show_booklet_render_result("竞赛分组表", path, stats)
        except Exception as exc:
            QMessageBox.critical(self, "生成失败", str(exc))

    def generate_full_booklet_document(self) -> None:
        if not self.database:
            QMessageBox.information(self, "请先新建项目", "生成文档前，请先新建或打开运动会项目。")
            return
        if not self._registration_sources_are_current():
            return
        assignments = self.database.load_heat_assignments()
        if not assignments:
            QMessageBox.information(self, "尚未编排", "请先到“抽签编排”生成并确认分组。")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "生成完整秩序册", "示例学校田径运动会秩序册.docx", "Word 文档 (*.docx)"
        )
        if not path:
            return
        if not path.lower().endswith(".docx"):
            path += ".docx"
        try:
            athletes, _ = self.database.load_registration()
            stats = render_full_booklet(
                path, self.database.load_schedule(), assignments, athletes,
                config=self.database.load_config(),
            )
            self.booklet_summary_label.setText(
                f"完整秩序册已生成：{len(athletes)} 名运动员、{stats.event_count} 个项目、"
                f"{stats.table_count} 张分组表、{stats.assignment_count} 人次。"
            )
            self._show_booklet_render_result("完整秩序册", path, stats)
        except Exception as exc:
            QMessageBox.critical(self, "生成失败", str(exc))

    def _show_booklet_render_result(self, document_name: str, path: str, stats) -> None:
        correction = f"\n自动缩小字号：{stats.font_adjustment_count} 个单元格。"
        if not stats.layout_warnings:
            QMessageBox.information(
                self, "生成完成", f"{document_name}已生成：\n{path}{correction}\n文本长度检查：通过。请在 Word/WPS 中预览最终分页。"
            )
            return
        details = "\n".join(f"- {warning}" for warning in stats.layout_warnings[:8])
        omitted = len(stats.layout_warnings) - 8
        if omitted > 0:
            details += f"\n- 另有 {omitted} 处"
        QMessageBox.warning(
            self,
            "生成完成（需检查）",
            f"{document_name}已生成：\n{path}{correction}\n"
            f"版式检查发现 {len(stats.layout_warnings)} 处可能换行：\n{details}",
        )

    def export_validation_report(self) -> None:
        if not self.latest_report:
            QMessageBox.information(self, "没有核对结果", "请先导入数据或执行全项目校验。")
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出核对报告", "导入核对报告.xlsx", "Excel 工作簿 (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            render_validation_report(self.latest_report, path)
            QMessageBox.information(self, "导出完成", "核对报告已生成。")
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))

    def closeEvent(self, event) -> None:
        self._close_database()
        super().closeEvent(event)
