# ui/mainwindow.py
from PyQt5.QtWidgets import (
    QMainWindow, QLabel, QPushButton, QVBoxLayout, QWidget, QCheckBox,
    QInputDialog, QMessageBox, QLineEdit, QApplication, QFileDialog # Added QFileDialog
)
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
import os
import sys
import win32com.client
from ui.acronymswindow import fetch_acronym_list_online
# Import the new function
try:
    from macros.CleanDocument import process_word_document
except ImportError:
    print("Warning: Could not import CleanDocument macro.")
    process_word_document = None


class MainWindow(QMainWindow):
    @staticmethod
    def get_resource_path(relative_path):
        """ Get the absolute path for a resource, """ \
            """works for dev and for PyInstaller """
        base_path = getattr(
            sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)

    def __init__(self):
        super().__init__()
        self.active_doc_path = None
        self.setWindowTitle("Nutrasource Copilot")
        self.setFixedSize(400, 200) # Increased size for the new button
        self.layout = QVBoxLayout()
        dir_path = self.get_resource_path(".")
        icon_path = os.path.join(dir_path, "leaf.ico")
        self.setWindowIcon(QIcon(icon_path))
        self.active_doc_label = QLabel("Active Word Document: None", self)
        self.layout.addWidget(self.active_doc_label)
        self.stay_on_top_checkbox = QCheckBox("Stay on Top", self)
        self.stay_on_top_checkbox.stateChanged.connect(self.toggle_stay_on_top)
        self.layout.addWidget(self.stay_on_top_checkbox)
        self.replace_values_window = None

        self.replace_values_selection_button = QPushButton(
            "Replace Values", self)
        self.replace_values_selection_button.setStyleSheet("""
            QPushButton {
                background-color: #5E2D91;
                color: white;
            }
            QPushButton:hover {
                background-color: #3C1A56;
            }
        """)
        self.replace_values_selection_button.clicked.connect(
            self.open_replace_values_selection_window)
        self.layout.addWidget(self.replace_values_selection_button)

        self.acronyms_button = QPushButton("Acronyms Table", self)
        self.acronyms_button.setStyleSheet("""
            QPushButton {
                background-color: #5E2D91;
                color: white;
            }
            QPushButton:hover {
                background-color: #3C1A56;
            }
        """)
        self.acronyms_button.clicked.connect(self.open_acronyms_window)
        self.layout.addWidget(self.acronyms_button)

        # --- New Button Added ---
        self.clean_doc_button = QPushButton("Clean & Protect Document", self)
        self.clean_doc_button.setStyleSheet("""
            QPushButton {
                background-color: #5E2D91;
                color: white;
            }
            QPushButton:hover {
                background-color: #3C1A56;
            }
        """)
        self.clean_doc_button.clicked.connect(self.run_clean_document)
        self.layout.addWidget(self.clean_doc_button)
        # --- End New Button ---

        self.label = QLabel("", self)
        self.layout.addWidget(self.label)

        self.widget = QWidget()
        self.widget.setLayout(self.layout)
        self.setCentralWidget(self.widget)
        self.layout.addStretch()

        self.timer = QTimer()
        self.timer.timeout.connect(self.update_active_document)
        self.timer.start(1000)
        self.acronyms_window = None
        self.replace_values_window = None
        self.acronyms_data = {}

    def update_active_document(self):
        # This function still runs to update the label, but isn't
        # strictly needed for the new 'Clean & Protect' workflow.
        try:
            word_app = win32com.client.GetActiveObject('Word.Application')
            active_doc = word_app.ActiveDocument
            self.active_doc_label.setText(
                "Active Document: " + active_doc.Name)
            # We keep tracking this, but don't use it for cleaning now
            self.active_doc_path = os.path.join(
                active_doc.Path, active_doc.Name)
        except Exception:
            self.active_doc_label.setText("Active Document: None")
            self.active_doc_path = None

    def open_replace_values_selection_window(self):
        from .replacevalues_selectionwindow import ReplaceValuesSelectionWindow
        self.replace_values_window = ReplaceValuesSelectionWindow(self)
        self.replace_values_window.show()

    def open_acronyms_window(self):
        from .acronymswindow import AcronymsWindow
        self.acronyms_window = AcronymsWindow(self)
        self.acronyms_window.show()

    # --- Method Modified ---
    def run_clean_document(self):
        if process_word_document is None:
            QMessageBox.critical(self, "Error", "Clean Document module failed to load.")
            return

        # --- Open File Dialog to Select Document ---
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Choose a Word Document to Clean", "",
            "Word Documents (*.docx *.doc)", options=options)

        # If the user cancelled the dialog, file_path will be empty
        if not file_path:
            self.label.setText("File selection cancelled.")
            QTimer.singleShot(3000, self.label.clear)
            return
        # --- End File Dialog ---

        # Proceed to get the password
        password, ok = QInputDialog.getText(
            self, "Password Required", "Enter the document password:", QLineEdit.Password
        )

        if ok and password:
            try:
                self.label.setText(f"Processing {os.path.basename(file_path)}... Please wait.")
                QApplication.processEvents() # Update UI to show the message

                # Call the processing function with the selected file path
                success, messages = process_word_document(file_path, password)

                log_message = "\n".join(messages)
                if success:
                    QMessageBox.information(self, "Success", f"Processing finished for {os.path.basename(file_path)}.\n\nLog:\n{log_message}")
                else:
                    QMessageBox.critical(self, "Error", f"Processing failed for {os.path.basename(file_path)}.\n\nLog:\n{log_message}")

            except Exception as e:
                QMessageBox.critical(self, "Error", f"An unexpected error occurred: {str(e)}")
            finally:
                self.label.clear() # Clear the status label
        elif ok:
             QMessageBox.warning(self, "Input Error", "Password cannot be empty for processing.")
        else:
            self.label.setText("Cleaning cancelled (no password entered).")
            QTimer.singleShot(3000, self.label.clear)
    # --- End Modified Method ---

    def toggle_stay_on_top(self, checked):
        flags = self.windowFlags()
        if checked:
            self.setWindowFlags(flags | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(flags & ~Qt.WindowStaysOnTopHint)
        self.show()

        if self.replace_values_window is not None:
             if checked: self.replace_values_window.setWindowFlags(self.replace_values_window.windowFlags() | Qt.WindowStaysOnTopHint)
             else: self.replace_values_window.setWindowFlags(self.replace_values_window.windowFlags() & ~Qt.WindowStaysOnTopHint)
             self.replace_values_window.show()

        if self.acronyms_window is not None:
             if checked: self.acronyms_window.setWindowFlags(self.acronyms_window.windowFlags() | Qt.WindowStaysOnTopHint)
             else: self.acronyms_window.setWindowFlags(self.acronyms_window.windowFlags() & ~Qt.WindowStaysOnTopHint)
             self.acronyms_window.show()


    def prefetch_acronyms(self):
        try:
            url = (
                "https://raw.githubusercontent.com/IIDelta/Doc_Companion/"
                "main/acronyms/acronym%20list.txt"
            )
            cache = os.path.join(
                os.path.expanduser("~"), ".doc_companion", "acronym_list.txt"
            )
            fetch_acronym_list_online(url, cache)
        except Exception as e:
            print("Prefetch failed:", e)