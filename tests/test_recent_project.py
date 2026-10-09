import logging
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QDate, QSettings
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QDialog, QMessageBox, QPushButton, QTableWidget, QTableWidgetItem

from sport_meeting.core.database import ProjectDatabase
from sport_meeting.core.heat_generation import generate_heat_assignments
from sport_meeting.core.importers import import_registration_workbooks, import_schedule_workbook
from sport_meeting.core.legacy_booklet_content import LAST_YEAR_SCHOOL_RECORDS, LATEST_SCHOOL_RECORDS
from sport_meeting.core.models import ROUND_MODE_PRELIMINARY_FINAL, ProjectConfig
from sport_meeting.core.workbook_templates import (
    create_grouping_registration_sample,
    create_grouping_schedule_sample,
)
from sport_meeting.gui.main_window import MainWindow, ScoreInputDelegate


class RecentProjectTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.application = QApplication.instance() or QApplication([])

    def test_last_project_is_remembered_and_reopened(self):
        with tempfile.TemporaryDirectory() as directory:
            QSettings.setDefaultFormat(QSettings.Format.IniFormat)
            QSettings.setPath(QSettings.Format.IniFormat, QSettings.Scope.UserScope, directory)
            settings = QSettings("LeishiSchool", "SportMeetingBooklet")
            settings.clear()

            project_path = Path(directory) / "最近项目.sportsmeet"
            first = MainWindow(logging.getLogger("recent-project-test"))
            first.database = ProjectDatabase.create(project_path)
            first._set_project(str(project_path))
            first.close()

            second = MainWindow(logging.getLogger("recent-project-test"))
            second.open_last_project()
            self.assertIsNotNone(second.database)
            self.assertIn(project_path.name, second.project_label.text())
            second.close()
            settings.clear()

    def test_existing_project_can_load_latest_records_without_changing_other_chapters(self):
        with tempfile.TemporaryDirectory() as directory:
            project_path = Path(directory) / "旧项目.sportsmeet"
            window = MainWindow(logging.getLogger("latest-school-records-test"))
            window.database = ProjectDatabase.create(
                project_path,
                ProjectConfig(
                    opening_ceremony_text="本届自定义开幕式",
                    school_records_text=LAST_YEAR_SCHOOL_RECORDS,
                ),
            )

            def choose_latest(dialog):
                button = next(
                    button for button in dialog.findChildren(QPushButton)
                    if button.text() == "载入最新校纪录"
                )
                button.click()
                return QDialog.DialogCode.Accepted

            with (
                patch.object(QDialog, "exec", choose_latest),
                patch.object(QMessageBox, "question", return_value=QMessageBox.StandardButton.Yes),
                patch.object(QMessageBox, "information"),
            ):
                window.edit_booklet_content()
            config = window.database.load_config()
            self.assertEqual(config.school_records_text, LATEST_SCHOOL_RECORDS)
            self.assertEqual(config.opening_ceremony_text, "本届自定义开幕式")
            window.close()

    def test_time_score_editor_inserts_separators_from_unit(self):
        seconds_editor = ScoreInputDelegate(lambda: "SECOND").createEditor(None, None, None)
        QTest.keyClicks(seconds_editor, "5932")
        self.assertEqual(seconds_editor.text(), "59:32")

        minutes_editor = ScoreInputDelegate(lambda: "MINUTE").createEditor(None, None, None)
        QTest.keyClicks(minutes_editor, "15329")
        self.assertEqual(minutes_editor.text(), "1:53:29")

        table = QTableWidget(1, 1)
        table.setItem(0, 0, QTableWidgetItem("7:50"))
        index = table.model().index(0, 0)
        delegate = ScoreInputDelegate(lambda: "SECOND")
        existing_editor = delegate.createEditor(None, None, index)
        delegate.setEditorData(existing_editor, index)
        self.assertEqual(existing_editor.text(), "07:50")
        delegate.setModelData(existing_editor, table.model(), index)
        self.assertEqual(table.item(0, 0).text(), "7:50")

    def test_registration_folder_page_imports_workbooks_and_updates_preview(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            create_grouping_registration_sample(root / "班级报名.xlsx")
            window = MainWindow(logging.getLogger("registration-folder-test"))
            project_path = root / "报名项目.sportsmeet"
            window.database = ProjectDatabase.create(project_path)
            window._set_project(str(project_path))
            with (
                patch(
                    "sport_meeting.gui.main_window.QFileDialog.getExistingDirectory",
                    return_value=str(root),
                ),
                patch("sport_meeting.gui.main_window.QMessageBox.information"),
                patch("sport_meeting.gui.main_window.QMessageBox.warning"),
                patch("sport_meeting.gui.main_window.QMessageBox.critical") as critical,
            ):
                window.import_registration_folder()
            self.assertFalse(critical.called)
            self.assertEqual(window.registration_table.model().rowCount(), 25)
            self.assertIn("文件 1 个", window.registration_summary_label.text())
            self.assertIn("运动员 25 人", window.registration_summary_label.text())
            window.close()

    def test_registration_template_settings_are_saved_with_project(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            project_path = root / "模板设置项目.sportsmeet"
            window = MainWindow(logging.getLogger("template-settings-test"))
            window.database = ProjectDatabase.create(project_path)
            window._set_project(str(project_path))
            table = window.registration_event_editors["junior_boys_events"]
            table.setRowCount(0)
            window._add_registration_event_row(table, "50米")
            window._add_registration_event_row(
                table, "测试项目", ROUND_MODE_PRELIMINARY_FINAL
            )
            tug_row = next(
                row for row in range(window.team_event_table.rowCount())
                if window.team_event_table.item(row, 0).text() == "拔河比赛"
            )
            window.team_event_table.item(tug_row, 4).setText("2")
            saved = window._save_registration_template_settings(show_confirmation=False)
            self.assertIsNotNone(saved)
            self.assertEqual(
                window.database.load_config().junior_boys_events,
                ("50米", "测试项目"),
            )
            self.assertEqual(
                window.database.load_config().junior_boys_event_formats,
                {"50米": ROUND_MODE_PRELIMINARY_FINAL, "测试项目": ROUND_MODE_PRELIMINARY_FINAL},
            )
            self.assertEqual(
                next(
                    event.manual_team_count
                    for event in window.database.load_config().team_events
                    if event.event_name == "拔河比赛"
                ),
                2,
            )
            window.close()

            reopened = ProjectDatabase.open(project_path)
            self.assertEqual(reopened.load_config().junior_boys_events, ("50米", "测试项目"))
            self.assertEqual(
                reopened.load_config().junior_boys_event_formats["测试项目"],
                ROUND_MODE_PRELIMINARY_FINAL,
            )
            self.assertEqual(
                next(
                    event.manual_team_count
                    for event in reopened.load_config().team_events
                    if event.event_name == "拔河比赛"
                ),
                2,
            )
            reopened.close()

    def test_meeting_dates_are_saved_and_reloaded_on_home_page(self):
        with tempfile.TemporaryDirectory() as directory:
            project_path = Path(directory) / "日期项目.sportsmeet"
            window = MainWindow(logging.getLogger("meeting-date-test"))
            window.database = ProjectDatabase.create(project_path)
            window._set_project(str(project_path))
            window.meeting_start_date_edit.setDate(QDate(2026, 10, 3))
            window.meeting_end_date_edit.setDate(QDate(2026, 10, 5))
            self.assertEqual(window.meeting_start_date_edit.text(), "2026年10月3日")
            self.assertEqual(window.meeting_start_date_edit.locale().name(), "zh_CN")
            self.assertEqual(window.meeting_start_date_edit.font().family(), "Microsoft YaHei UI")
            self.assertEqual(
                window.meeting_start_date_edit.calendarWidget().font().family(),
                "Microsoft YaHei UI",
            )
            window.save_meeting_dates()
            config = window.database.load_config()
            self.assertEqual(config.meeting_start_date, "2026-10-03")
            self.assertEqual(config.meeting_end_date, "2026-10-05")
            self.assertIn("2026年10月3日", window.meeting_dates_label.text())
            window.close()

    def test_manual_click_ranks_each_heat_independently_and_persists(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            registration = import_registration_workbooks([
                create_grouping_registration_sample(root / "报名.xlsx")
            ])
            schedule = import_schedule_workbook(
                create_grouping_schedule_sample(root / "日程.xlsx")
            )
            generated = generate_heat_assignments(
                ProjectConfig(), registration.athletes, registration.entries,
                schedule.rounds, seed=2026,
            )
            project_path = root / "手动排名.sportsmeet"
            window = MainWindow(logging.getLogger("manual-ranking-test"))
            window.database = ProjectDatabase.create(project_path)
            window.database.replace_registration(registration.athletes, registration.entries)
            window.database.replace_schedule(schedule.rounds)
            window.database.replace_heat_assignments(generated.assignments, generated.seed)
            window._set_project(str(project_path))

            self.assertEqual(window.result_unit_combo.currentData(), "SECOND")
            window.manual_rank_button.setChecked(True)
            first_row_by_heat = {}
            for row, assignment in enumerate(window.result_rows):
                first_row_by_heat.setdefault(assignment.heat_no, row)
            for row in first_row_by_heat.values():
                window.assign_manual_rank(row, 0)
            self.assertEqual(
                [window.result_table.item(row, 7).text() for row in first_row_by_heat.values()],
                ["1", "1", "1"],
            )
            with (
                patch("sport_meeting.gui.main_window.QMessageBox.information"),
                patch("sport_meeting.gui.main_window.QMessageBox.critical") as critical,
            ):
                window.save_results_and_generate_final()
            self.assertFalse(critical.called)
            preliminary_id = window.result_event_combo.currentData()
            saved = window.database.load_results(preliminary_id)
            self.assertEqual(sorted(item.manual_rank for item in saved if item.manual_rank), [1, 1, 1])
            self.assertEqual(len(window.database.load_qualifications(preliminary_id)), 3)
            self.assertEqual(window.final_material_event_combo.count(), 1)
            self.assertEqual(window.final_material_table.model().rowCount(), 3)
            self.assertIn("道次已确定", window.final_material_summary_label.text())
            for row, score in enumerate(["10:10", "10:20", "10:20"]):
                window.final_material_table.setItem(row, 5, QTableWidgetItem(score))
            with (
                patch("sport_meeting.gui.main_window.QMessageBox.information"),
                patch("sport_meeting.gui.main_window.QMessageBox.critical") as critical,
            ):
                window.save_final_results()
            self.assertFalse(critical.called)
            final_round_id = window.final_material_event_combo.currentData()
            ranked = [window.final_material_table.item(row, 7).text() for row in range(3)]
            self.assertEqual(ranked, ["1", "2", "2"])
            self.assertEqual(len(window.database.load_results(final_round_id)), 3)
            window.close()


if __name__ == "__main__":
    unittest.main()
