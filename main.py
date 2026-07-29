"""
Doc Companion application entry point.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import nltk
from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QApplication

from ui.mainwindow import MainWindow


def get_resource_path(relative_path: str) -> str:
    """Resolve source or PyInstaller-packaged resource paths."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = Path(__file__).resolve().parent

    return str(base_path / relative_path)


def configure_nltk_data_path() -> None:
    """Add bundled NLTK data when the application is packaged."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        nltk_data_dir = os.path.join(sys._MEIPASS, "nltk_data")

        if nltk_data_dir not in nltk.data.path:
            nltk.data.path.append(nltk_data_dir)


def load_stylesheet(application: QApplication) -> None:
    """Load the optional Qt stylesheet without blocking application startup."""
    style_path = get_resource_path(os.path.join("ui", "style.qss"))

    try:
        with open(style_path, "r", encoding="utf-8") as style_file:
            application.setStyleSheet(style_file.read())
    except FileNotFoundError:
        print(
            f"Warning: style.qss not found at {style_path}. "
            "Using default application styling."
        )
    except Exception as exc:
        print(f"Warning: could not load stylesheet: {exc}")


def main() -> None:
    """Create the Qt application and show the main Doc Companion window."""
    configure_nltk_data_path()

    application = QApplication(sys.argv)
    load_stylesheet(application)

    main_window = MainWindow()
    main_window.show()

    QTimer.singleShot(0, main_window.prefetch_acronyms)

    sys.exit(application.exec_())


if __name__ == "__main__":
    main()
