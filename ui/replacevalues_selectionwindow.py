# ui/replacevalues_selectionwindow.py
from PyQt5.QtWidgets import (QMainWindow, QVBoxLayout, QHBoxLayout, QCheckBox,
                             QWidget, QPushButton, QLabel, QFileDialog,
                             QTextEdit, QMessageBox)
import os
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QTimer

try:
    from macros.ReplaceValues_Selection import Macro_ReplaceValues_Selection
except ImportError:
    print("Error: Could not import ReplaceValues_Selection macro.")
    Macro_ReplaceValues_Selection = None


class WildcardsInfoWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Wildcards Information")
        self.setMinimumSize(350, 250) # Increased height

        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setHtml(
            """
            <h3>Word Wildcard Characters</h3>
            <table style="width:100%; border-collapse: collapse;">
                <tr style="background-color:#E8E0F1; font-weight:bold;">
                    <th style="text-align:left; padding: 5px; border: 1px solid #C0C0C0;">Wildcard</th>
                    <th style="text-align:left; padding: 5px; border: 1px solid #C0C0C0;">Matches</th>
                </tr>
                <tr><td style="padding: 5px; border: 1px solid #C0C0C0;">*</td><td style="padding: 5px; border: 1px solid #C0C0C0;">Any sequence of characters</td></tr>
                <tr><td style="padding: 5px; border: 1px solid #C0C0C0;">?</td><td style="padding: 5px; border: 1px solid #C0C0C0;">Any single character</td></tr>
                <tr><td style="padding: 5px; border: 1px solid #C0C0C0;">#</td><td style="padding: 5px; border: 1px solid #C0C0C0;">Any single digit</td></tr>
                <tr><td style="padding: 5px; border: 1px solid #C0C0C0;">[abc]</td><td style="padding: 5px; border: 1px solid #C0C0C0;">a, b, or c</td></tr>
                <tr><td style="padding: 5px; border: 1px solid #C0C0C0;">[a-z]</td><td style="padding: 5px; border: 1px solid #C0C0C0;">Any character from a to z</td></tr>
                <tr><td style="padding: 5px; border: 1px solid #C0C0C0;">[!abc]</td><td style="padding: 5px; border: 1px solid #C0C0C0;">Any character except a, b, or c</td></tr>
                <tr><td style="padding: 5px; border: 1px solid #C0C0C0;">~*</td><td style="padding: 5px; border: 1px solid #C0C0C0;">Finds a literal *</td></tr>
                <tr><td style="padding: 5px; border: 1px solid #C0C0C0;">~?</td><td style="padding: 5px; border: 1px solid #C0C0C0;">Finds a literal ?</td></tr>
                <tr><td style="padding: 5px; border: 1px solid #C0C0C0;">~~</td><td style="padding: 5px; border: 1px solid #C0C0C0;">Finds a literal ~</td></tr>
            </table>
            <p><i>Note: Wildcards are only used if the 'Use Wildcards' box is checked.</i></p>
            """
        )
        self.setCentralWidget(self.text_edit)
        self.update_stay_on_top()

    def update_stay_on_top(self):
        parent = self.parent()
        if parent and hasattr(parent, 'parent') and hasattr(parent.parent(), 'stay_on_top_checkbox'):
             main_win = parent.parent()
             flags = self.windowFlags()
             if main_win.stay_on_top_checkbox.isChecked():
                self.setWindowFlags(flags | Qt.WindowStaysOnTopHint)
             else:
                self.setWindowFlags(flags & ~Qt.WindowStaysOnTopHint)
             self.show()

class ReplaceValuesSelectionWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.excel_file = None
        self.wildcards_info_window = None
        self.setWindowTitle("Replace Values in Selection")
        self.setFixedSize(450, 180)

        dir_path = os.path.dirname(os.path.realpath(__file__))
        leaf_icon_path = os.path.join(dir_path, "leaf.png")
        excel_icon_path = os.path.join(dir_path, "excel.png")
        self.setWindowIcon(QIcon(leaf_icon_path))

        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(10, 10, 10, 10)
        self.layout.setSpacing(10)

        excel_layout = QHBoxLayout()
        self.excel_file_button = QPushButton(QIcon(excel_icon_path), " Choose Excel File", self)
        self.excel_file_button.clicked.connect(self.choose_excel_file)
        excel_layout.addWidget(self.excel_file_button)
        self.excel_file_label = QLabel("No file chosen.", self)
        excel_layout.addWidget(self.excel_file_label)
        excel_layout.addStretch()
        self.layout.addLayout(excel_layout)

        self.use_wildcards_checkbox = QCheckBox("Use Wildcards (See info)", self)
        self.use_wildcards_checkbox.stateChanged.connect(self.toggle_wildcards_info_window)
        self.layout.addWidget(self.use_wildcards_checkbox)

        self.run_macro_button = QPushButton("Run Macro", self)
        self.run_macro_button.clicked.connect(self.run_macro)
        self.layout.addWidget(self.run_macro_button)

        self.layout.addStretch()

        self.widget = QWidget()
        self.widget.setLayout(self.layout)
        self.setCentralWidget(self.widget)
        self.update_stay_on_top()

    def toggle_wildcards_info_window(self, checked):
        if checked:
            if self.wildcards_info_window is None or not self.wildcards_info_window.isVisible():
                self.wildcards_info_window = WildcardsInfoWindow(self)
                # Position it relative to the main window or checkbox
                # self.wildcards_info_window.move(self.pos().x() + self.width(), self.pos().y())
                self.wildcards_info_window.show()
            else:
                 self.wildcards_info_window.activateWindow()
        else:
            if self.wildcards_info_window is not None:
                self.wildcards_info_window.close()
                self.wildcards_info_window = None

    def choose_excel_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Choose an Excel File", "",
            "Excel Files (*.xlsx *.xls)", options=options)
        if file_path:
            self.excel_file = file_path
            self.excel_file_label.setText("File: " + os.path.basename(self.excel_file))
        else:
            self.excel_file = None
            self.excel_file_label.setText("No file chosen.")

    def run_macro(self):
        if Macro_ReplaceValues_Selection is None:
            QMessageBox.critical(self, "Error", "Replace Values module failed to load.")
            return

        if not self.excel_file:
            QMessageBox.warning(self, "Warning", "Please choose an Excel file first.")
            return

        try:
            macro = Macro_ReplaceValues_Selection()
            macro.load_excel_file(self.excel_file)
            use_wildcards = self.use_wildcards_checkbox.isChecked()
            result = macro.replace_values(use_wildcards)

            if result is not None:
                 QMessageBox.critical(self, "Error", f"Macro execution failed:\n{result}")
            else:
                 macro.save_document()
                 QMessageBox.information(self, "Success", "Replacements made in selected text.")
                 if self.parent() and hasattr(self.parent(), 'label'):
                     self.parent().label.setText("Replacements completed.")
                     QTimer.singleShot(3000, self.parent().label.clear)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An unexpected error occurred: {str(e)}")
            print(f"Error running replace macro: {e}")
            if self.parent() and hasattr(self.parent(), 'label'):
                self.parent().label.setText(f"Error: {str(e)}")
                QTimer.singleShot(5000, self.parent().label.clear)

    def update_stay_on_top(self):
        parent = self.parent()
        if parent and hasattr(parent, 'stay_on_top_checkbox'):
            flags = self.windowFlags()
            if parent.stay_on_top_checkbox.isChecked():
                self.setWindowFlags(flags | Qt.WindowStaysOnTopHint)
            else:
                self.setWindowFlags(flags & ~Qt.WindowStaysOnTopHint)
            self.show()

    def closeEvent(self, event):
        if self.wildcards_info_window is not None:
            self.wildcards_info_window.close()
        super().closeEvent(event)