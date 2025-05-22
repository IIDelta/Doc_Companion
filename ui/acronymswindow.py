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


# Modify the function signature and internal usage
def fetch_acronym_list_online(url, base_cache_path): # Renamed parameter
    """
    Fetch the acronym list from the online URL.
    On success, write the file to base_cache_path and return the path.
    On failure, if a cached base file exists, return that;
    otherwise, raise an exception.
    """
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        # Ensure parent directory for the base cache path exists
        os.makedirs(os.path.dirname(base_cache_path), exist_ok=True)
        with open(base_cache_path, "w", encoding="utf-8") as f: # Writes to base_cache_path
            f.write(response.text)
        return base_cache_path
    except Exception as e:
        if os.path.exists(base_cache_path): # Checks existing base_cache_path
            return base_cache_path
        else:
            raise Exception(f"Failed to fetch base acronym list online: {e}")


class AcronymsWindow(QMainWindow):
    def __init__(self, parent=None):
        self.base_acronym_file_path = None
        self.user_acronym_file_path = os.path.join(os.path.expanduser("~"), ".doc_companion", "user_acronyms.txt")
        # Ensure the .doc_companion directory exists
        os.makedirs(os.path.dirname(self.user_acronym_file_path), exist_ok=True)
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
            # Define path for the base list fetched from online
            doc_companion_dir = os.path.join(os.path.expanduser("~"), ".doc_companion")
            base_list_filename = "base_acronym_list.txt" # New filename for base list
            self.base_acronym_file_path = os.path.join(doc_companion_dir, base_list_filename)

            # self.user_acronym_file_path is already set in __init__

            # Fetch the online list into the base file path
            # fetch_acronym_list_online now ensures its parent directory exists
            effective_base_path = fetch_acronym_list_online(url, self.base_acronym_file_path)
            print(f"Using base acronym list from: {effective_base_path}")
            # Ensure self.base_acronym_file_path reflects the actual path used (could be cache on error)
            self.base_acronym_file_path = effective_base_path


            # find_acronyms and get_definition will need to accept two paths
            acronyms = find_acronyms(word_app, self.base_acronym_file_path, self.user_acronym_file_path)

            for category, table in [("likely", self.likely_table),
                                    ("possible", self.possible_table),
                                    ("unlikely", self.unlikely_table)]:
                table.setRowCount(0)
                for acronym, context in acronyms[category].items():
                    table.insertRow(table.rowCount())
                    table.setItem(table.rowCount()-1, 0, QTableWidgetItem(acronym))
                    # get_definition will also need to check both files
                    definition_text = get_definition(acronym, self.base_acronym_file_path, self.user_acronym_file_path)
                    table.setItem(table.rowCount()-1, 1, QTableWidgetItem(definition_text))
                    checkbox = QCheckBox()
                    checkbox.setChecked(category != "unlikely")
                    table.setCellWidget(table.rowCount()-1, 2, checkbox)
                    table.setItem(table.rowCount()-1, 3, QTableWidgetItem(context))
        except Exception as e:
            print(str(e))
            self.parent().label.setText(f"Error: {str(e)}")
            QTimer.singleShot(5000, self.parent().label.clear)

    def generate_table(self):
        # --- Part 1: Save new/updated definitions from UI to user_acronyms.txt ---
        if self.user_acronym_file_path: # Check if user_acronym_file_path is set
            user_definitions = {} # Store definitions from user_acronyms.txt
            try:
                # Ensure the directory for user_acronym_file_path exists
                os.makedirs(os.path.dirname(self.user_acronym_file_path), exist_ok=True)
                with open(self.user_acronym_file_path, 'r', encoding='utf-8') as f_read:
                    for line in f_read:
                        line = line.strip()
                        if '\t' in line:
                            acr, deph = line.split('\t', 1)
                            user_definitions[acr] = deph
            except FileNotFoundError:
                print(f"Info: User acronym file '{self.user_acronym_file_path}' not found. Will be created if new definitions are added.")
            except Exception as e:
                print(f"Error reading user acronym list {self.user_acronym_file_path}: {e}")

            made_changes_to_user_list = False
            for table_widget in [self.possible_table, self.unlikely_table]: # Or all tables if you want to save from "Likely" too
                for r in range(table_widget.rowCount()):
                    acronym_item = table_widget.item(r, 0)
                    definition_item = table_widget.item(r, 1)

                    if acronym_item and definition_item:
                        ui_acronym = acronym_item.text().strip()
                        ui_definition = definition_item.text().strip()

                        if ui_acronym and ui_definition: # Only save if both are non-empty
                            if ui_acronym not in user_definitions or \
                            user_definitions[ui_acronym] != ui_definition:
                                user_definitions[ui_acronym] = ui_definition
                                made_changes_to_user_list = True

            if made_changes_to_user_list:
                try:
                    with open(self.user_acronym_file_path, 'w', encoding='utf-8') as f_write:
                        for acr, deph in sorted(user_definitions.items()): # Sort for consistency
                            f_write.write(f"{acr}\t{deph}\n")
                    print(f"Info: User acronym list '{self.user_acronym_file_path}' updated.")
                    if self.parent() and hasattr(self.parent(), 'label'):
                        self.parent().label.setText("User acronym list updated.")
                        QTimer.singleShot(3000, self.parent().label.clear)
                except Exception as e:
                    print(f"Error writing user acronym list to {self.user_acronym_file_path}: {e}")
                    if self.parent() and hasattr(self.parent(), 'label'):
                        self.parent().label.setText(f"Error updating user list: {e}")
                        QTimer.singleShot(5000, self.parent().label.clear)
        else:
            print("Warning: User acronym file path not set. Cannot save new definitions.")
            # ... (optional status message) ...

        # --- Part 2: logic to generate the Word document table ---
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
