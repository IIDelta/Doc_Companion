# ui/replacevalueswindow.py

from PyQt5.QtWidgets import (QMainWindow, QFileDialog, QVBoxLayout, 
                            QWidget, QPushButton, QLabel, QMessageBox)
from macros.ReplaceValues import Macro_ReplaceValues
import os
from PyQt5.QtGui import QIcon

class ReplaceValuesWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Replace Values")
        self.setFixedSize(400, 100)
        # Get the directory of this script
        dir_path = os.path.dirname(os.path.realpath(__file__))
        # Build the full icon path
        icon_path = os.path.join(dir_path, "leaf.png")
        # Set the window icon
        self.setWindowIcon(QIcon(icon_path))  

        self.layout = QVBoxLayout()

        self.word_file_button = QPushButton("Choose Word Document", self)
        self.word_file_button.setStyleSheet("""
            QPushButton {
                background-color: #5E2D91;
                color: white;
            }
            QPushButton:hover {
                background-color: #3C1A56;
            }
        """)
        self.word_file_button.clicked.connect(self.choose_word_file)
        self.layout.addWidget(self.word_file_button)

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
        self.run_macro_button.setEnabled(False)  # Disable the button by default
        self.layout.addWidget(self.run_macro_button)

        self.widget = QWidget()
        self.widget.setLayout(self.layout)
        self.setCentralWidget(self.widget)
        self.layout.addStretch()

    def choose_word_file(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Choose a Word Document", 
                                        "", "Word Documents (*.docx)", options=options)
        if fileName:
            self.word_file = fileName
            self.layout.removeWidget(self.word_file_button)
            self.word_file_button.deleteLater()
            self.word_file_button = None
            self.word_file_label = QLabel("<b><font color=#5E2D91>\
                                        Word File:</font></b> "
                                        + os.path.basename(self.word_file), self)
            self.layout.addWidget(self.word_file_label)
            self.check_files_selected()

    def choose_excel_file(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Choose an Excel File", 
                                            "", "Excel Files (*.xlsx)", options=options)
        if fileName:
            self.excel_file = fileName
            self.layout.removeWidget(self.excel_file_button)
            self.excel_file_button.deleteLater()
            self.excel_file_button = None
            self.excel_file_label = QLabel("<b><font color=#5E2D91>\
                                        Excel File:</font></b> " 
                                        + os.path.basename(self.excel_file), self)
            self.layout.addWidget(self.excel_file_label)
            self.check_files_selected()

    def check_files_selected(self):
        # Enable the "Run Macro" button if both files have been selected
        if hasattr(self, 'word_file') and hasattr(self, 'excel_file'):
            self.run_macro_button.setEnabled(True)

    def run_macro(self):
        macro = Macro_ReplaceValues()
        macro.load_document(self.word_file)
        macro.load_excel_file(self.excel_file)
        macro.replace_values()
        macro.save_document(self.word_file)

        msg = QMessageBox()
        msg.setWindowTitle("Macro")
        msg.setText(f"Replace Values ran successfully on {os.path.basename(self.word_file)}")
        msg.exec_()
