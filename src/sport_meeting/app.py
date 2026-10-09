from __future__ import annotations

import os
import multiprocessing as mp
import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMessageBox

from .gui.main_window import MainWindow
from .logging_config import configure_logging
from .single_instance import SingleInstance


def _app_data_directory() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    return base / "LeishiSportMeeting"


def main() -> int:
    mp.freeze_support()
    smoke_test = "--smoke-test" in sys.argv
    if smoke_test:
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        sys.argv.remove("--smoke-test")
    application = QApplication(sys.argv)
    application.setApplicationName("学校运动会秩序册与决赛编排系统")
    application.setStyle("Fusion")
    if smoke_test:
        logger = configure_logging(_app_data_directory() / "logs")
        window = MainWindow(logger)
        window.close()
        return 0
    with SingleInstance() as instance:
        if instance.already_running:
            QMessageBox.information(None, "软件已在运行", "请切换到已经打开的运动会系统窗口。")
            return 0
        logger = configure_logging(_app_data_directory() / "logs")
        window = MainWindow(logger)
        window.show()
        return application.exec()


if __name__ == "__main__":
    raise SystemExit(main())
