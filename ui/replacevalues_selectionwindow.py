# ui/replacevalues_selectionwindow.py

from PyQt5.QtWidgets import (QMainWindow, QVBoxLayout, QCheckBox,
                             QWidget, QPushButton, QLabel, QFileDialog,
                             QTextEdit)
import os
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QTimer


from macros.ReplaceValues_Selection import Macro_ReplaceValues_Selection


class WildcardsInfoWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Wildcards Information")

        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setHtml(
            """
            <table style="width:100%">
                <tr>
                    <th style="text-align:left">Wildcard</th>
                    <th style="text-align:left">Matches</th>
                </tr>
                <tr>
                    <td>*</td>
                    <td>Any sequence of characters</td>
                </tr>
                <tr>
                    <td>?</td>
                    <td>Any single character</td>
                </tr>
                <tr>
                    <td>#</td>
                    <td>Any single digit</td>
                </tr>
                <tr>
                    <td>[abc]</td>
                    <td>a, b, or c</td>
                </tr>
                <tr>
                    <td>[!abc]</td>
                    <td>Any character except a, b, or c</td>
                </tr>
            </table>
            """
        )
        self.setCentralWidget(self.text_edit)


class ReplaceValuesSelectionWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Replace Values in Selection")
        self.setFixedSize(400, 100)
        dir_path = os.path.dirname(os.path.realpath(__file__))
        # Build the full icon path
        icon_path = os.path.join(dir_path, "leaf.png")
        # Set the window icon
        self.setWindowIcon(QIcon(icon_path))
        self.layout = QVBoxLayout()

        # Update the window flags based on the parent's checkbox state
        if parent is not None and parent.stay_on_top_checkbox.isChecked():
            self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)

        self.excel_file_button = QPushButton("Choose Excel File", self)
        self.excel_file_button.setStyleSheet("""
            QPushButton {
                background-color: #5E2D91;
                color: white;
            }
            QPushButton:hover {
                background-color: #3C1A56;
            }
        """)
        self.excel_file_button.clicked.connect(self.choose_excel_file)
        self.layout.addWidget(self.excel_file_button)
        self.excel_file_label = QLabel("", self)  # Create the label
        self.layout.addWidget(self.excel_file_label)

        # Add the checkbox for enabling wildcards
        self.use_wildcards_checkbox = QCheckBox("Use Wildcards", self)
        self.layout.addWidget(self.use_wildcards_checkbox)
        self.use_wildcards_checkbox.stateChanged.connect(
            self.toggle_wildcards_info_window)
        self.run_macro_button = QPushButton("Run Macro", self)
        self.run_macro_button.setStyleSheet("""
            QPushButton {
                background-color: #5E2D91;
                color: white;
            }
            QPushButton:hover {
                background-color: #3C1A56;
            }
        """)
        self.run_macro_button.clicked.connect(self.run_macro)
        self.layout.addWidget(self.run_macro_button)

        self.widget = QWidget()
        self.widget.setLayout(self.layout)
        self.setCentralWidget(self.widget)
        self.layout.addStretch()

    def toggle_wildcards_info_window(self, checked):
        if checked:
            self.wildcards_info_window = WildcardsInfoWindow(self)
            self.wildcards_info_window.show()
        else:
            if self.wildcards_info_window is not None:
                self.wildcards_info_window.close()
                self.wildcards_info_window = None

    def choose_excel_file(self):
        options = QFileDialog.Options()
        self.excel_file, _ = QFileDialog.getOpenFileName(
            self, "Choose an Excel File", "",
            "Excel Files (*.xlsx)", options=options)
        if self.excel_file:
            self.excel_file_label.setText(
                "Excel File: " + os.path.basename(self.excel_file))

    def run_macro(self):
        try:
            if not self.excel_file:
                print("No Excel file chosen.")
                return
            macro = Macro_ReplaceValues_Selection()
            macro.load_excel_file(self.excel_file)
            use_wildcards = self.use_wildcards_checkbox.isChecked()
            macro.replace_values(use_wildcards)
            macro.save_document()
            print("Replacements made in selected text.")

        except Exception as e:
            print(str(e))
            self.parent().label.setText(f"Error: {str(e)}")
            QTimer.singleShot(5000, self.parent().label.clear)
