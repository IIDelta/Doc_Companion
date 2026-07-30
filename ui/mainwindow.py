"""
mainwindow.py

Primary Doc Companion launcher window.
"""

from __future__ import annotations

import os
import sys
import threading
from pathlib import Path
from typing import Optional

import win32com.client
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from .acronymswindow import fetch_acronym_list_online

try:
    from macros.CleanDocument import process_word_document
except ImportError:
    process_word_document = None

from .windowchrome import applydarktitlebar



class MainWindow(QMainWindow):
    """Primary Doc Companion tool-selection window."""

    def __init__(self) -> None:
        super().__init__()

        self.active_doc_path: Optional[str] = None

        self.acronyms_window = None
        self.replace_values_window = None
        self.ref_finder_window = None
        self.please_review_window = None
        self.referencecheckwindow = None
        self.citation_migration_window = None


        self.setWindowTitle("Doc Companion")
        self.setFixedSize(460, 330)
        self.setWindowIcon(QIcon(self.get_resource_path("leaf.ico")))

        self._build_ui()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_active_document)
        self.timer.start(1000)

        self.update_active_document()

    @staticmethod
    def get_resource_path(relative_path: str) -> str:
        """
        Resolve a UI resource location in source and PyInstaller environments.
        """
        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            base_path = Path(sys._MEIPASS) / "ui"
        else:
            base_path = Path(__file__).resolve().parent

        return str(base_path / relative_path)

    def _build_ui(self) -> None:
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        top_bar_layout = QHBoxLayout()

        self.active_doc_label = QLabel("Active Document: None", self)
        self.active_doc_label.setWordWrap(True)

        self.stay_on_top_checkbox = QCheckBox("Stay on Top", self)
        self.stay_on_top_checkbox.stateChanged.connect(
            self.toggle_stay_on_top
        )

        top_bar_layout.addWidget(self.active_doc_label)
        top_bar_layout.addStretch()
        top_bar_layout.addWidget(self.stay_on_top_checkbox)

        layout.addLayout(top_bar_layout)

        replace_button = QPushButton(
            QIcon(self.get_resource_path("replace.ico")),
            " Replace Values",
            self,
        )
        replace_button.clicked.connect(
            self.open_replace_values_selection_window
        )
        layout.addWidget(replace_button)

        acronyms_button = QPushButton(
            QIcon(self.get_resource_path("acronym.ico")),
            " Acronyms Table",
            self,
        )
        acronyms_button.clicked.connect(self.open_acronyms_window)
        layout.addWidget(acronyms_button)

        clean_button = QPushButton(
            QIcon(self.get_resource_path("clean.ico")),
            " Clean && Protect Document",
            self,
        )
        clean_button.clicked.connect(self.run_clean_document)
        layout.addWidget(clean_button)

        reference_button = QPushButton(
            QIcon(self.get_resource_path("audit.ico")),
            " Reference/Open-Access Finder",
            self,
        )
        reference_button.clicked.connect(self.open_ref_finder_window)
        layout.addWidget(reference_button)

        please_review_button = QPushButton(
            QIcon(self.get_resource_path("audit.ico")),
            " Prepare PleaseReview Copy",
            self,
        )
        please_review_button.clicked.connect(
            self.open_please_review_window
        )
        layout.addWidget(please_review_button)

        self.referencecheckbutton = QPushButton(
            " Reference Library Check", self
        )
        self.referencecheckbutton.clicked.connect(
            self.openreferencecheckwindow
        )
        layout.addWidget(self.referencecheckbutton)
        citation_migration_button = QPushButton(
            " Citation Migration",
            self,
        )

        citation_migration_button.setObjectName(
            "UtilityButton"
        )

        citation_migration_button.clicked.connect(
            self.open_citation_migration_window
        )

        layout.addWidget(citation_migration_button)



        layout.addStretch()

        self.label = QLabel("", self)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setWordWrap(True)
        layout.addWidget(self.label)

    def update_active_document(self) -> None:
        """Refresh the visible name of the active Word document."""
        try:
            word_app = win32com.client.GetActiveObject("Word.Application")
            active_document = word_app.ActiveDocument

            self.active_doc_label.setText(
                f"Active Document: {active_document.Name}"
            )
            self.active_doc_path = os.path.join(
                active_document.Path,
                active_document.Name,
            )
        except Exception:
            self.active_doc_label.setText("Active Document: None")
            self.active_doc_path = None

    def open_replace_values_selection_window(self) -> None:
        """Open the Replace Values tool window."""
        from .replacevalues_selectionwindow import (
            ReplaceValuesSelectionWindow,
        )

        if (
            self.replace_values_window is not None
            and self.replace_values_window.isVisible()
        ):
            self.replace_values_window.activateWindow()
            self.replace_values_window.raise_()
            return

        self.replace_values_window = ReplaceValuesSelectionWindow(self)
        self.replace_values_window.show()
        applydarktitlebar(self.referencecheckwindow)


    def open_acronyms_window(self) -> None:
        """Open the Acronyms Table tool window."""
        from .acronymswindow import AcronymsWindow

        if (
            self.acronyms_window is not None
            and self.acronyms_window.isVisible()
        ):
            self.acronyms_window.activateWindow()
            self.acronyms_window.raise_()
            return

        self.acronyms_window = AcronymsWindow(self)
        self.acronyms_window.show()
        applydarktitlebar(self.referencecheckwindow)


    def open_ref_finder_window(self) -> None:
        """Open the Reference/Open-Access Finder window."""
        from .ref_finder_window import RefFinderWindow

        if (
            self.ref_finder_window is not None
            and self.ref_finder_window.isVisible()
        ):
            self.ref_finder_window.activateWindow()
            self.ref_finder_window.raise_()
            return

        self.ref_finder_window = RefFinderWindow(self)
        self.ref_finder_window.show()
        applydarktitlebar(self.referencecheckwindow)

    def open_please_review_window(self) -> None:
        """Open the tracked-changes PleaseReview preparation window."""
        from .pleasereviewwindow import PleaseReviewWindow

        if (
            self.please_review_window is not None
            and self.please_review_window.isVisible()
        ):
            self.please_review_window.activateWindow()
            self.please_review_window.raise_()
            return

        self.please_review_window = PleaseReviewWindow(self)
        self.please_review_window.show()
        applydarktitlebar(self.referencecheckwindow)

    def run_clean_document(self) -> None:
        """Run the existing Clean & Protect workflow."""
        if process_word_document is None:
            QMessageBox.critical(
                self,
                "Error",
                "Clean Document module failed to load.",
            )
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose a Word Document to Clean",
            "",
            "Word Documents (*.docx *.doc)",
        )

        if not file_path:
            self.label.setText("File selection cancelled.")
            QTimer.singleShot(3000, self.label.clear)
            return

        try:
            self.label.setText(
                f"Processing {os.path.basename(file_path)}. Please wait."
            )

            success, messages = process_word_document(file_path)
            log_message = "\n".join(messages)

            if success:
                QMessageBox.information(
                    self,
                    "Success",
                    f"Processing finished.\n\nLog:\n{log_message}",
                )
            else:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Processing failed.\n\nLog:\n{log_message}",
                )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Error",
                f"An unexpected error occurred: {exc}",
            )
        finally:
            self.label.clear()

    def toggle_stay_on_top(self, checked: int) -> None:
        """Apply stay-on-top state to the main window and open child windows."""
        flags = self.windowFlags()

        if checked:
            self.setWindowFlags(flags | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(flags & ~Qt.WindowStaysOnTopHint)

        self.show()

        for child_window in (
            self.replace_values_window,
            self.acronyms_window,
            self.ref_finder_window,
            self.please_review_window,
        ):
            if (
                child_window is not None
                and child_window.isVisible()
                and hasattr(child_window, "update_stay_on_top")
            ):
                child_window.update_stay_on_top()

        if (
            self.citation_migration_window is not None
            and self.citation_migration_window.isVisible()
        ):
            self.citation_migration_window.setWindowFlags(
                self.citation_migration_window.windowFlags()
                | Qt.WindowStaysOnTopHint
                if checked
                else self.citation_migration_window.windowFlags()
                & ~Qt.WindowStaysOnTopHint
            )

            self.citation_migration_window.show()


    def prefetch_acronyms(self) -> None:
        """
        Download/cache the base acronym list without blocking initial UI display.
        """
        thread = threading.Thread(
            target=self._prefetch_acronyms_worker,
            daemon=True,
        )
        thread.start()

    def openreferencecheckwindow(self):
        from .referencecheckwindow import ReferenceCheckWindow
        if (self.referencecheckwindow is not None and
                self.referencecheckwindow.isVisible()):
            self.referencecheckwindow.activateWindow()
            self.referencecheckwindow.raise_()
        else:
            self.referencecheckwindow = ReferenceCheckWindow(self)
            self.referencecheckwindow.show()
            applydarktitlebar(self.referencecheckwindow)

    def open_citation_migration_window(self):
        from .citationmigrationwindow import (
            CitationMigrationWindow,
        )

        if (
            self.citation_migration_window is not None
            and self.citation_migration_window.isVisible()
        ):
            self.citation_migration_window.activateWindow()
            self.citation_migration_window.raise_()

        else:
            self.citation_migration_window = (
                CitationMigrationWindow(self)
            )

            self.citation_migration_window.show()


    @staticmethod
    def _prefetch_acronyms_worker() -> None:
        try:
            url = (
                "https://raw.githubusercontent.com/IIDelta/Doc_Companion/"
                "main/acronyms/acronym%20list.txt"
            )
            cache_dir = Path.home() / ".doc_companion"
            cache_dir.mkdir(parents=True, exist_ok=True)

            cache_path = cache_dir / "base_acronym_list.txt"
            fetch_acronym_list_online(
                url,
                str(cache_path),
            )
        except Exception as exc:
            print(f"Acronym prefetch failed: {exc}")
