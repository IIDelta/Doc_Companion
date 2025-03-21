# ui/acronymswindow.py

from PyQt5.QtWidgets import (QMainWindow, QVBoxLayout, QWidget,
                             QPushButton, QFileDialog, QSizePolicy,
                             QTableWidget, QTableWidgetItem, QCheckBox,
                             QTabWidget)
import os
import sys
import requests
import win32com.client
from PyQt5.QtGui import QIcon
from docx import Document
from PyQt5.QtCore import Qt, QTimer
from macros.Acronyms import find_acronyms, get_definition


def fetch_acronym_list_online(url, cache_path):
    """
    Fetch the acronym list from the online URL.
    On success, write the file to cache_path and return the path.
    On failure, if a cached file exists, return that;
    otherwise, raise an exception.
    """
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        os.makedirs(os.path.dirname(cache_path), exist_ok=True)
        with open(cache_path, "w", encoding="utf-8") as f:
            f.write(response.text)
        return cache_path
    except Exception as e:
        # If fetching failed but a cached file exists, use it.
        if os.path.exists(cache_path):
            return cache_path
        else:
            raise Exception(f"Failed to fetch acronym list online: {e}")


class AcronymsWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Acronym Finder")
        self.setMinimumSize(800, 600)
        # Get the directory of this script
        dir_path = os.path.dirname(os.path.realpath(__file__))
        # Build the full icon path
        icon_path = os.path.join(dir_path, "leaf.png")
        # Set the window icon
        self.setWindowIcon(QIcon(icon_path))
        self.layout = QVBoxLayout()
        self.run_macro_button = QPushButton("Find Acronyms", self)
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
        self.tab_widget = QTabWidget(self)
        self.tab_widget.setMinimumHeight(400)
        self.tab_widget.setSizePolicy(
            QSizePolicy.Expanding, QSizePolicy.Expanding)  # Add this line
        self.layout.addWidget(self.tab_widget)
        self.likely_table = self.create_table()
        self.possible_table = self.create_table()
        self.unlikely_table = self.create_table()
        # Create a QVBoxLayout for each table and add the table and "+" button
        likely_layout = QVBoxLayout()
        likely_layout.addWidget(self.likely_table)
        self.add_row_button_likely = QPushButton("+", self)
        self.add_row_button_likely.clicked.connect(
            lambda: self.add_new_row(self.likely_table))
        likely_layout.addWidget(self.add_row_button_likely)
        likely_widget = QWidget()
        likely_widget.setLayout(likely_layout)

        possible_layout = QVBoxLayout()
        possible_layout.addWidget(self.possible_table)
        self.add_row_button_possible = QPushButton("+", self)
        self.add_row_button_possible.clicked.connect(
            lambda: self.add_new_row(self.possible_table))
        possible_layout.addWidget(self.add_row_button_possible)
        possible_widget = QWidget()
        possible_widget.setLayout(possible_layout)

        unlikely_layout = QVBoxLayout()
        unlikely_layout.addWidget(self.unlikely_table)
        self.add_row_button_unlikely = QPushButton("+", self)
        self.add_row_button_unlikely.clicked.connect(
            lambda: self.add_new_row(self.unlikely_table))
        unlikely_layout.addWidget(self.add_row_button_unlikely)
        unlikely_widget = QWidget()
        unlikely_widget.setLayout(unlikely_layout)
        self.tab_widget.addTab(likely_widget, "Likely")
        self.tab_widget.addTab(possible_widget, "Possible")
        self.tab_widget.addTab(unlikely_widget, "Unlikely")
        # Set the default column widths for each table
        self.likely_table.setColumnWidth(0, 100)
        self.likely_table.setColumnWidth(1, 150)
        self.likely_table.setColumnWidth(2, 60)
        self.likely_table.setColumnWidth(3, 400)

        self.possible_table.setColumnWidth(0, 100)
        self.possible_table.setColumnWidth(1, 150)
        self.possible_table.setColumnWidth(2, 60)
        self.possible_table.setColumnWidth(3, 400)

        self.unlikely_table.setColumnWidth(0, 100)
        self.unlikely_table.setColumnWidth(1, 150)
        self.unlikely_table.setColumnWidth(2, 60)
        self.unlikely_table.setColumnWidth(3, 400)

        # Inside the __init__ method after creating the tables
        self.likely_check_all_button = QPushButton(
            'Likely: Check/Uncheck All', self)
        self.likely_check_all_button.setStyleSheet(
            'QPushButton {color: #5E2D91; font-weight: bold;}')
        self.likely_check_all_button.clicked.connect(
            lambda: self.check_uncheck_all(self.likely_table))
        self.layout.addWidget(self.likely_check_all_button)

        self.possible_check_all_button = QPushButton(
            'Possible: Check/Uncheck All', self)
        self.possible_check_all_button.setStyleSheet(
            'QPushButton {color: #5E2D91; font-weight: bold;}')
        self.possible_check_all_button.clicked.connect(
            lambda: self.check_uncheck_all(self.possible_table))
        self.layout.addWidget(self.possible_check_all_button)

        self.unlikely_check_all_button = QPushButton(
            'Unlikely: Check/Uncheck All', self)
        self.unlikely_check_all_button.setStyleSheet(
            'QPushButton {color: #5E2D91; font-weight: bold;}')
        self.unlikely_check_all_button.clicked.connect(
            lambda: self.check_uncheck_all(self.unlikely_table))
        self.layout.addWidget(self.unlikely_check_all_button)

        self.generate_table_button = QPushButton(
            "Generate Table", self)
        self.generate_table_button.setStyleSheet("""
            QPushButton {
                background-color: #5E2D91;
                color: white;
            }
            QPushButton:hover {
                background-color: #3C1A56;
            }
        """)
        self.generate_table_button.clicked.connect(self.generate_table)
        self.layout.addWidget(self.generate_table_button)

        self.widget = QWidget()
        self.widget.setLayout(self.layout)
        self.setCentralWidget(self.widget)

    def create_table(self):
        table = QTableWidget(0, 4)  # 0 rows, 4 columns
        table.setHorizontalHeaderLabels(
            ["Acronym", "Likely Definition", "Include", "Context"])
        table.setMinimumHeight(300)
        table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        return table

    @staticmethod
    def get_resource_path(relative_path):
        """ Get the absolute path for a
          resource, works for dev and for PyInstaller """
        base_path = getattr(
            sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)

    def run_macro(self):
        try:
            word_app = win32com.client.Dispatch('Word.Application')

            if not word_app.Documents.Count:
                raise Exception("No active word document.")

            # URL for the acronym list hosted on GitHub.
            # (Replace 'yourusername' and 'yourrepo' with 
            # your actual GitHub username and repository.)
            url = (
                "https://raw.githubusercontent.com/IIDelta/Doc_Companion/"
                "main/acronyms/acronym%20list.txt"
            )
            # Define a cache path in the user's home directory (hidden folder).
            cache_path = os.path.join(
                os.path.expanduser("~"), ".doc_companion", "acronym_list.txt")

            # Try fetching the online acronym
            #  list (with caching and error handling)
            definition_path = fetch_acronym_list_online(url, cache_path)
            print(f"Using acronym list from: {definition_path}")

            acronyms = find_acronyms(word_app, definition_path)

            for category, table in [("likely", self.likely_table),
                                    ("possible", self.possible_table),
                                    ("unlikely", self.unlikely_table)]:
                table.setRowCount(0)  # Clear the table
                for acronym, context in acronyms[category].items():
                    table.insertRow(table.rowCount())
                    table.setItem(
                        table.rowCount()-1, 0, QTableWidgetItem(acronym))
                    table.setItem(table.rowCount()-1, 1, QTableWidgetItem(get_definition(acronym, definition_path)))
                    checkbox = QCheckBox()
                    checkbox.setChecked(category != "unlikely")
                    table.setCellWidget(table.rowCount()-1, 2, checkbox)
                    table.setItem(table.rowCount()-1, 3, QTableWidgetItem(context))
        except Exception as e:
            print(str(e))
            self.parent().label.setText(f"Error: {str(e)}")
            QTimer.singleShot(5000, self.parent().label.clear)

    def generate_table(self):
        try:
            doc = Document()

            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save Word Document", "", "Word Documents (*.docx)")
            if file_path:
                table = doc.add_table(rows=0, cols=2)
                table.style = 'Table Grid'
                rows = []
                for acronym_table in [
                        self.likely_table,
                        self.possible_table,
                        self.unlikely_table]:
                    for i in range(acronym_table.rowCount()):
                        acronym_item = acronym_table.item(i, 0)
                        definition_item = acronym_table.item(i, 1)
                        checkbox = acronym_table.cellWidget(i, 2)
                        if checkbox.isChecked():
                            rows.append(
                                (acronym_item.text(), definition_item.text()))

                # Sort the rows
                rows.sort()

                # Add the sorted rows to the table
                for acronym, definition in rows:
                    cells = table.add_row().cells
                    cells[0].text = acronym
                    cells[1].text = definition

                doc.save(file_path)
        except Exception as e:
            print(str(e))
            self.parent().label.setText(f"Error: {str(e)}")
            QTimer.singleShot(5000, self.parent().label.clear)

    # In AcronymsWindow:
    def update_stay_on_top(self):
        if self.parent().stay_on_top_checkbox.isChecked():
            self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)
        self.show()
        # You need to call show() again for the change to take effect

    def check_uncheck_all(self, table):
        # This will keep track of the new state for checkboxes
        new_state = None

        for i in range(table.rowCount()):
            checkbox = table.cellWidget(i, 2)
            # On the first iteration,
            # decide whether to check or uncheck all boxes
            if new_state is None:
                new_state = not checkbox.isChecked()
            checkbox.setChecked(new_state)

    def add_new_row(self, table):
        # Insert a new row at the end of the table
        table.insertRow(table.rowCount())

        # Add empty cells for acronym and definition
        table.setItem(table.rowCount()-1, 0, QTableWidgetItem(""))
        table.setItem(table.rowCount()-1, 1, QTableWidgetItem(""))

        # Add a checkbox in the "Include" column
        checkbox = QCheckBox()
        checkbox.setChecked(True)
        table.setCellWidget(table.rowCount()-1, 2, checkbox)

        # Add an empty cell for context
        table.setItem(table.rowCount()-1, 3, QTableWidgetItem(""))