# ui/mainwindow.py
from PyQt5.QtWidgets import (
    QMainWindow, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QCheckBox,
    QMessageBox, QFileDialog
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QTimer, QCoreApplication
import os
import sys
import win32com.client
from .acronymswindow import fetch_acronym_list_online # Use relative import

try:
    from macros.CleanDocument import process_word_document
except ImportError:
    print("Warning: Could not import CleanDocument macro.")
    process_word_document = None


class MainWindow(QMainWindow):
    @staticmethod
    def get_resource_path(relative_path):
        """ Get the absolute path for a resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
            # For PyInstaller, resources are usually in the root or a subdir
            # We assume icons are in a 'ui' subdir *within* the bundle
            return os.path.join(base_path, 'ui', relative_path)
        except AttributeError:
            # If not running as a PyInstaller bundle, use the script's directory
            # Go up one level from 'ui' to the project root, then down to 'ui'
            base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            return os.path.join(base_path, 'ui', relative_path)

    def __init__(self):
        super().__init__()
        self.active_doc_path = None
        self.setWindowTitle("Nutrasource Copilot")
        self.setFixedSize(400, 250)

        # --- Icon Paths ---
        icon_path = self.get_resource_path("/icons/leaf.ico")
        replace_icon_path = self.get_resource_path("/icons/replace.ico")
        acronym_icon_path = self.get_resource_path("/icons/acronym.ico")
        clean_icon_path = self.get_resource_path("/icons/clean.ico")
        # --- End Icon Paths ---

        self.setWindowIcon(QIcon(icon_path))

        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(15, 15, 15, 15)
        self.layout.setSpacing(10)

        # --- Top Bar (Label & Checkbox) ---
        top_bar_layout = QHBoxLayout()
        self.active_doc_label = QLabel("Active Word Document: None", self)
        top_bar_layout.addWidget(self.active_doc_label)
        top_bar_layout.addStretch()
        self.stay_on_top_checkbox = QCheckBox("Stay on Top", self)
        self.stay_on_top_checkbox.stateChanged.connect(self.toggle_stay_on_top)
        top_bar_layout.addWidget(self.stay_on_top_checkbox)
        self.layout.addLayout(top_bar_layout)
        # --- End Top Bar ---

        self.replace_values_selection_button = QPushButton(
            QIcon(replace_icon_path), " Replace Values", self)
        self.replace_values_selection_button.clicked.connect(
            self.open_replace_values_selection_window)
        self.layout.addWidget(self.replace_values_selection_button)

        self.acronyms_button = QPushButton(
            QIcon(acronym_icon_path), " Acronyms Table", self)
        self.acronyms_button.clicked.connect(self.open_acronyms_window)
        self.layout.addWidget(self.acronyms_button)

        self.clean_doc_button = QPushButton(
            QIcon(clean_icon_path), " Clean && Protect Document", self)
        self.clean_doc_button.clicked.connect(self.run_clean_document)
        self.layout.addWidget(self.clean_doc_button)

        self.layout.addStretch()

        self.label = QLabel("", self)
        self.label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.label)

        self.widget = QWidget()
        self.widget.setLayout(self.layout)
        self.setCentralWidget(self.widget)

        self.timer = QTimer()
        self.timer.timeout.connect(self.update_active_document)
        self.timer.start(1000)
        self.acronyms_window = None
        self.replace_values_window = None
        self.update_active_document()

    def update_active_document(self):
        try:
            word_app = win32com.client.GetActiveObject('Word.Application')
            active_doc = word_app.ActiveDocument
            self.active_doc_label.setText(
                "Active Document: " + active_doc.Name)
            self.active_doc_path = os.path.join(
                active_doc.Path, active_doc.Name)
        except Exception:
            self.active_doc_label.setText("Active Document: None")
            self.active_doc_path = None

    def open_replace_values_selection_window(self):
        from .replacevalues_selectionwindow import ReplaceValuesSelectionWindow
        if self.replace_values_window is not None and self.replace_values_window.isVisible():
            self.replace_values_window.activateWindow()
        else:
            self.replace_values_window = ReplaceValuesSelectionWindow(self)
            self.replace_values_window.show()

    def open_acronyms_window(self):
        from .acronymswindow import AcronymsWindow
        if self.acronyms_window is not None and self.acronyms_window.isVisible():
            self.acronyms_window.activateWindow()
        else:
            self.acronyms_window = AcronymsWindow(self)
            self.acronyms_window.show()

    def run_clean_document(self):
        if process_word_document is None:
            QMessageBox.critical(self, "Error", "Clean Document module failed to load.")
            return

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Choose a Word Document to Clean", "",
            "Word Documents (*.docx *.doc)", options=options)

        if not file_path:
            self.label.setText("File selection cancelled.")
            QTimer.singleShot(3000, self.label.clear)
            return

        try:
            self.label.setText(f"Processing {os.path.basename(file_path)}... Please wait.")
            QCoreApplication.processEvents()

            success, messages = process_word_document(file_path)

            log_message = "\n".join(messages)
            if success:
                QMessageBox.information(self, "Success", f"Processing finished.\n\nLog:\n{log_message}")
            else:
                QMessageBox.critical(self, "Error", f"Processing failed.\n\nLog:\n{log_message}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred: {str(e)}")
        finally:
            self.label.clear()

    def toggle_stay_on_top(self, checked):
        flags = self.windowFlags()
        if checked:
            self.setWindowFlags(flags | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(flags & ~Qt.WindowStaysOnTopHint)
        self.show()

        if self.replace_values_window is not None and self.replace_values_window.isVisible():
             self.replace_values_window.update_stay_on_top()

        if self.acronyms_window is not None and self.acronyms_window.isVisible():
             self.acronyms_window.update_stay_on_top()

    def prefetch_acronyms(self):
        try:
            url = (
                "https://raw.githubusercontent.com/IIDelta/Doc_Companion/"
                "main/acronyms/acronym%20list.txt"
            )
            cache_dir = os.path.join(os.path.expanduser("~"), ".doc_companion")
            cache = os.path.join(cache_dir, "base_acronym_list.txt")
            os.makedirs(cache_dir, exist_ok=True)
            fetch_acronym_list_online(url, cache)
            print("Acronym list prefetched.")
        except Exception as e:
            print(f"Prefetch failed: {e}")