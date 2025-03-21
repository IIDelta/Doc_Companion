# Import necessary modules from PyQt5.QtWidgets
from PyQt5.QtWidgets import (
    QMainWindow, QLabel, QPushButton, QVBoxLayout, QWidget, QCheckBox)
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
import os
import sys
import win32com.client
from ui.acronymswindow import fetch_acronym_list_online


# Define MainWindow class, which is a QMainWindow subclass
class MainWindow(QMainWindow):
    @staticmethod
    def get_resource_path(relative_path):
        """ Get the absolute path for a resource, """ \
            """works for dev and for PyInstaller """
        base_path = getattr(
            sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)

    # Initialize the class
    def __init__(self):
        super().__init__()  # Call the superclass's __init__ function
        self.active_doc_path = None
        self.setWindowTitle("Nutrasource Copilot")  # Set the window title
        self.setFixedSize(400, 150)  # Set the size of the window
        self.layout = QVBoxLayout()
        # Get the directory of this script
        dir_path = self.get_resource_path(".")
        # Build the full icon path
        icon_path = os.path.join(dir_path, "leaf.ico")
        self.setWindowIcon(QIcon(icon_path))
        self.active_doc_label = QLabel("Active Word Document: None", self)
        self.layout.addWidget(self.active_doc_label)
        self.stay_on_top_checkbox = QCheckBox("Stay on Top", self)
        self.stay_on_top_checkbox.stateChanged.connect(self.toggle_stay_on_top)
        self.layout.addWidget(self.stay_on_top_checkbox)
        self.replace_values_window = None  # Add this line

        # Repeat for "Replace Values From Selection" button
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

        self.label = QLabel("", self)  # QLabel to show messages
        self.layout.addWidget(self.label)

        # Create a QWidget, set its layout to the QVBoxLayout we created,
        # and set it as the central widget of the QMainWindow
        self.widget = QWidget()
        self.widget.setLayout(self.layout)
        self.setCentralWidget(self.widget)
        self.layout.addStretch()

        # Create a QTimer
        self.timer = QTimer()
        # Connect the timer's timeout
        # signal to the update_active_document function
        self.timer.timeout.connect(self.update_active_document)
        # Start the timer to trigger every 1000 ms (1 second)
        self.timer.start(1000)
        self.acronyms_window = None
        self.replace_values_window = None
        self.acronyms_data = {}

    def update_active_document(self):
        try:
            word_app = win32com.client.GetActiveObject('Word.Application')
            active_doc_name = word_app.ActiveDocument.Name
            self.active_doc_label.setText(
                "Active Document: " + active_doc_name)

            # Update active_doc_path
            active_doc_path = word_app.ActiveDocument.Path
            active_doc_filename = word_app.ActiveDocument.Name
            self.active_doc_path = os.path.join(
                active_doc_path, active_doc_filename)

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

    def toggle_stay_on_top(self, checked):
        flags = self.windowFlags()
        if checked:
            self.setWindowFlags(flags | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(flags & ~Qt.WindowStaysOnTopHint)
        self.show()
        # You need to call show() again for the change to take effect

        # Update replace_values_window's flags
        if self.replace_values_window is not None:
            if checked:
                self.replace_values_window.setWindowFlags(
                    flags | Qt.WindowStaysOnTopHint)
            else:
                self.replace_values_window.setWindowFlags(
                    flags & ~Qt.WindowStaysOnTopHint)
            self.replace_values_window.show()
        # Update acronyms_window's flags
        if self.acronyms_window is not None:
            if checked:
                self.acronyms_window.setWindowFlags(
                    flags | Qt.WindowStaysOnTopHint)
            else:
                self.acronyms_window.setWindowFlags(
                    flags & ~Qt.WindowStaysOnTopHint)
            self.acronyms_window.show()

    def prefetch_acronyms(self):
        # e.g. download & cache the GitHub list (no UI blocking)
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
