# ui/ref_finder_window.py
from PyQt5.QtWidgets import (
    QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QPushButton, QLabel,
    QFileDialog, QTextEdit, QMessageBox, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QSplitter, QApplication
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, QCoreApplication

import os
import sys
import queue
import threading
import traceback

# Import the backend processing functions and constants
from macros.RefFinder import UNPAYWALL_EMAIL, process_references_backend

# This function should ideally be imported or provided by the main application
# For now, replicate it here for standalone testing/clarity.
# In your main app, you should just get it from MainWindow.get_resource_path
def get_resource_path(relative_path):
    """
    Get the absolute path for resources (like icons)
    This static method helps find files whether running from source or
    as a bundle.
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
        # For PyInstaller, resources are usually in the root or a subdir
        # We assume icons are in a 'ui' subdir *within* the bundle
        return os.path.join(base_path, 'ui', relative_path)
    except AttributeError:
        # If not running as a PyInstaller bundle,
        # use the script's directory
        # Go up one level from 'ui' to the project root,
        # then down to 'ui'
        base_path = os.path.dirname(
            os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, 'ui', relative_path)


class RefFinderWindow(QMainWindow):
    # Define a custom signal for thread-safe GUI updates
    # This signal will carry a dictionary with 'type' (log/summary_add) and data
    gui_update_signal = pyqtSignal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Open Access Reference Checker")
        self.setMinimumSize(900, 700) # Increased size for better layout

        self.input_file_path = ""
        self.output_folder_path = ""
        self.processing_thread = None
        self.gui_update_queue = queue.Queue()

        # Connect the custom signal to the queue processing slot
        self.gui_update_signal.connect(self.process_gui_queue_item)

        # --- Icon Paths ---
        # Using the helper function to get resource paths
        icon_path = get_resource_path("leaf.ico")
        folder_icon_path = get_resource_path("folder.ico")
        file_icon_path = get_resource_path("file.ico")
        clipboard_icon_path = get_resource_path("clipboard.ico") # Assuming you have this icon
        self.setWindowIcon(QIcon(icon_path))
        # --- End Icon Paths ---

        # Central Widget and Layout
        self.central_widget = QWidget()
        self.main_layout = QVBoxLayout(self.central_widget)
        self.setCentralWidget(self.central_widget)

        # --- Top frame for controls ---
        controls_frame_layout = QHBoxLayout()

        # Input via File
        self.select_input_file_button = QPushButton(QIcon(file_icon_path), "Select Input File (.txt)", self)
        self.select_input_file_button.clicked.connect(self.select_input_file)
        controls_frame_layout.addWidget(self.select_input_file_button)

        self.input_file_label = QLabel("No file selected.", self)
        self.input_file_label.setObjectName("FilePathLabel") # For QSS styling
        self.input_file_label.setMinimumWidth(150) # Give it some width
        self.input_file_label.setWordWrap(True)
        controls_frame_layout.addWidget(self.input_file_label)
        
        # Spacer
        controls_frame_layout.addStretch(1)

        # Output Folder
        self.select_output_folder_button = QPushButton(QIcon(folder_icon_path), "Select Output Folder", self)
        self.select_output_folder_button.clicked.connect(self.select_output_folder)
        controls_frame_layout.addWidget(self.select_output_folder_button)

        self.output_folder_label = QLabel("No folder selected.", self)
        self.output_folder_label.setObjectName("FilePathLabel") # For QSS styling
        self.output_folder_label.setMinimumWidth(150) # Give it some width
        self.output_folder_label.setWordWrap(True)
        controls_frame_layout.addWidget(self.output_folder_label)
        
        # Spacer
        controls_frame_layout.addStretch(1)

        self.main_layout.addLayout(controls_frame_layout)

        # --- Processing Buttons Layout ---
        # This will hold the file/clipboard buttons and the copy summary button
        processing_actions_layout = QHBoxLayout() 

        # Left group: File and Clipboard processing buttons (stacked vertically)
        file_clipboard_buttons_layout = QVBoxLayout()
        self.start_file_processing_button = QPushButton("Start Processing (File)", self)
        self.start_file_processing_button.clicked.connect(self.start_file_processing)
        self.start_file_processing_button.setMinimumHeight(30)
        self.start_file_processing_button.setObjectName("PrimaryButton")
        file_clipboard_buttons_layout.addWidget(self.start_file_processing_button)

        self.process_from_clipboard_button = QPushButton(QIcon(clipboard_icon_path), "Process from Clipboard", self)
        self.process_from_clipboard_button.clicked.connect(self.start_clipboard_processing)
        self.process_from_clipboard_button.setMinimumHeight(30)
        file_clipboard_buttons_layout.addWidget(self.process_from_clipboard_button)
        
        processing_actions_layout.addLayout(file_clipboard_buttons_layout) # Add the vertical group to the horizontal layout

        # Spacer between processing buttons and copy summary
        processing_actions_layout.addStretch(1)
        
        # Right group: Copy Summary button
        self.copy_summary_button = QPushButton("Copy Summary", self)
        self.copy_summary_button.clicked.connect(self.copy_summary_to_clipboard)
        self.copy_summary_button.setEnabled(False) # Disabled until processing finishes
        processing_actions_layout.addWidget(self.copy_summary_button)

        self.main_layout.addLayout(processing_actions_layout) # Add this new overall horizontal layout to main


        # --- Main Splitter to divide summary and log ---
        self.splitter = QSplitter(Qt.Vertical)
        self.main_layout.addWidget(self.splitter)

        # --- Summary Treeview Frame (Top pane of Splitter) ---
        summary_widget = QWidget()
        summary_layout = QVBoxLayout(summary_widget)
        summary_layout.setContentsMargins(0, 0, 0, 0)
        
        summary_label = QLabel("Live Summary:", self)
        summary_label.setObjectName("SectionHeader")
        summary_layout.addWidget(summary_label)
        
        self.summary_tree = QTreeWidget(self)
        self.summary_tree.setHeaderLabels(["DOI", "Title / Status / File", "URL"])
        self.summary_tree.header().setSectionResizeMode(0, QHeaderView.Interactive)
        self.summary_tree.header().setSectionResizeMode(1, QHeaderView.Stretch) # Title/Status column stretches
        self.summary_tree.header().setSectionResizeMode(2, QHeaderView.Interactive)
        self.summary_tree.setColumnWidth(0, 150)
        self.summary_tree.setColumnWidth(2, 250)
        self.summary_tree.setSelectionBehavior(QTreeWidget.SelectRows) # Select full row
        summary_layout.addWidget(self.summary_tree)
        self.splitter.addWidget(summary_widget)

        self.tree_categories = {
            'downloaded_pdf': self.create_category_item("✅ Downloaded PDFs"),
            'open_access_no_pdf_link': self.create_category_item("🔑 OA (No PDF/Download Failed)"),
            'paywalled': self.create_category_item("💰 Paywalled"),
            'unpaywall_data_unavailable': self.create_category_item("⚠️ Unpaywall/SS Data Failed"),
            'doi_not_found': self.create_category_item("❌ DOI Not Found")
        }
        # Add categories to the tree
        for item in self.tree_categories.values():
            self.summary_tree.addTopLevelItem(item)
            item.setExpanded(True) # Start expanded

        # --- Log Area Frame (Bottom pane of Splitter) ---
        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(0, 0, 0, 0)

        log_label = QLabel("Detailed Log:", self)
        log_label.setObjectName("SectionHeader")
        log_layout.addWidget(log_label)
        self.log_text_area = QTextEdit(self)
        self.log_text_area.setReadOnly(True)
        log_layout.addWidget(self.log_text_area)
        self.splitter.addWidget(log_widget)

        # Set initial splitter sizes (ratio)
        self.splitter.setSizes([3, 2]) # Sets initial ratio to 60/40

        # Initial GUI state update
        self.update_stay_on_top()
        self.update_paths_labels()

    def create_category_item(self, text):
        """Helper to create QTreeWidget top-level items for categories."""
        item = QTreeWidgetItem([text, "", ""]) # Category items have no DOI/URL, only text
        item.setFlags(item.flags() & ~Qt.ItemIsSelectable) # Make categories non-selectable
        return item

    def update_paths_labels(self):
        self.input_file_label.setText(os.path.basename(self.input_file_path) if self.input_file_path else "No file selected.")
        self.output_folder_label.setText(os.path.basename(self.output_folder_path) if self.output_folder_path else "No folder selected.")


    def select_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select the .txt file with references", "",
            "Text files (*.txt);;All files (*.*)")
        if file_path:
            self.input_file_path = os.path.normpath(file_path)
            self.update_paths_labels()

    def select_output_folder(self):
        folder_path = QFileDialog.getExistingDirectory(
            self, "Select Folder to Save Downloaded PDFs", "")
        if folder_path:
            self.output_folder_path = os.path.normpath(folder_path)
            self.update_paths_labels()

    def process_gui_queue_item(self, queued_item):
        """Processes a single item from the GUI update queue."""
        if isinstance(queued_item, dict) and 'type' in queued_item:
            if queued_item['type'] == 'log':
                self.update_log_area(queued_item['message'])
            elif queued_item['type'] == 'summary_add':
                self.update_summary_treeview(queued_item['category_key'], queued_item['item_data'])
        else:
            self.update_log_area(str(queued_item))

    def update_log_area(self, message):
        self.log_text_area.insertPlainText(message + "\n")
        self.log_text_area.verticalScrollBar().setValue(self.log_text_area.verticalScrollBar().maximum())

    def update_summary_treeview(self, category_key, item_data):
        parent_item = self.tree_categories.get(category_key)
        if parent_item:
            doi_val = item_data.get('doi', 'N/A')
            title_val = item_data.get('title', 'No Title')
            status_info_val = item_data.get('status_info', 'N/A')
            url_val = item_data.get('url_info', 'N/A')
            
            # Combine title and status for the second column
            title_status_combined = f"{title_val} -- {status_info_val}"

            item = QTreeWidgetItem([doi_val, title_status_combined, url_val])
            parent_item.addChild(item)
            self.summary_tree.scrollToItem(item)


    def gui_update_callback(self, item_to_log):
        # Emit the signal to process the item on the main GUI thread
        self.gui_update_signal.emit(item_to_log)

    def start_file_processing(self):
        """Starts processing using the selected input file."""
        if not self.input_file_path:
            QMessageBox.warning(self, "Error", "Please select an input file.")
            return
        
        try:
            with open(self.input_file_path, 'r', encoding='utf-8') as f:
                references_from_file = [line.strip() for line in f if line.strip()]
            
            if not references_from_file:
                QMessageBox.warning(self, "Warning", f"No references found in '{os.path.basename(self.input_file_path)}' or the file is empty.")
                return

            self._initiate_processing(references_from_file)

        except FileNotFoundError:
            QMessageBox.critical(self, "Error", f"Input file '{os.path.basename(self.input_file_path)}' not found.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error reading input file: {str(e)}")


    def start_clipboard_processing(self):
        """Starts processing using content from the clipboard."""
        clipboard_text = QApplication.clipboard().text()
        if not clipboard_text:
            QMessageBox.warning(self, "Warning", "Clipboard is empty. Please copy references to the clipboard first.")
            return

        references_from_clipboard = [line.strip() for line in clipboard_text.splitlines() if line.strip()]

        if not references_from_clipboard:
            QMessageBox.warning(self, "Warning", "No valid references found in clipboard content.")
            return

        self._initiate_processing(references_from_clipboard)


    def _initiate_processing(self, references_list):
        """Common logic to start the background processing thread."""
        if not self.output_folder_path:
            QMessageBox.warning(self, "Error", "Please select an output folder.")
            return
        
        if UNPAYWALL_EMAIL == 'YOUR_EMAIL@example.com' or not UNPAYWALL_EMAIL:
             QMessageBox.critical(self, "Error", "Please set your Unpaywall email in 'macros/RefFinder.py' (UNPAYWALL_EMAIL variable).")
             return

        # Clear previous summary
        for category_item in self.tree_categories.values():
            category_item.takeChildren()
        
        self.log_text_area.clear()
        
        self.gui_update_callback({'type': 'log', 'message': "Starting processing..."})
        
        # Disable both processing buttons
        self.start_file_processing_button.setText("Processing...")
        self.start_file_processing_button.setEnabled(False)
        self.process_from_clipboard_button.setEnabled(False)
        self.copy_summary_button.setEnabled(False)

        self.processing_thread = threading.Thread(
            target=self.run_backend_task,
            args=(references_list, self.output_folder_path),
            daemon=True
        )
        self.processing_thread.start()

    def run_backend_task(self, references_list, output_folder):
        """The actual background task that calls the macro."""
        try:
            process_references_backend(references_list, output_folder, self.gui_update_callback)

        except Exception as e:
            self.gui_update_callback({'type': 'log', 'message': f"🚨 An unexpected error occurred in backend processing: {e}"})
            self.gui_update_callback({'type': 'log', 'message': traceback.format_exc()}) # Log full traceback
        finally:
            self.processing_finished()
            
    def processing_finished(self):
        # Use QTimer.singleShot(0, ...) to ensure these GUI updates happen on the main thread
        QTimer.singleShot(0, lambda: self.start_file_processing_button.setText("Start Processing (File)"))
        QTimer.singleShot(0, lambda: self.start_file_processing_button.setEnabled(True))
        QTimer.singleShot(0, lambda: self.process_from_clipboard_button.setEnabled(True))
        QTimer.singleShot(0, lambda: self.copy_summary_button.setEnabled(True))
        self.gui_update_callback({'type': 'log', 'message': "\n--- All Processing Finished. Check summary above. ---"})

    def copy_summary_to_clipboard(self):
        summary_lines = []
        try:
            for category_key, parent_item in self.tree_categories.items():
                category_text = parent_item.text(0) # Get text from the first column of the parent item
                summary_lines.append(f"\n{category_text}")
                
                children_count = parent_item.childCount()
                if children_count == 0:
                    summary_lines.append("\t(No items in this category)")
                    continue

                for i in range(children_count):
                    child_item = parent_item.child(i)
                    item_text_display = child_item.text(1) # This is the "Title / Status / File" combined column
                    doi_val = child_item.text(0) # DOI is in the first column
                    url_val = child_item.text(2) # URL is in the third column

                    summary_lines.append(f"\t- {item_text_display}")
                    if doi_val != 'N/A':
                        summary_lines.append(f"\t  DOI: {doi_val}")
                    summary_lines.append(f"\t  URL: {url_val}")
            
            clipboard_text = "\n".join(summary_lines)
            
            if not clipboard_text.strip():
                self.gui_update_callback({'type': 'log', 'message': "Summary is empty. Nothing to copy."})
                return

            QApplication.clipboard().setText(clipboard_text)
            self.gui_update_callback({'type': 'log', 'message': "Summary copied to clipboard!"})
            QMessageBox.information(self, "Copied", "Summary copied to clipboard!")

        except Exception as e:
            self.gui_update_callback({'type': 'log', 'message': f"Error copying summary: {e}"})
            QMessageBox.critical(self, "Error", f"Could not copy summary to clipboard: {e}")

    def update_stay_on_top(self):
        # This method assumes the parent of this window is the MainWindow
        # which has the stay_on_top_checkbox.
        parent = self.parent()
        if parent and hasattr(parent, 'stay_on_top_checkbox'):
             main_win = parent # In this case, parent is MainWindow
             flags = self.windowFlags()
             if main_win.stay_on_top_checkbox.isChecked():
                self.setWindowFlags(flags | Qt.WindowStaysOnTopHint)
             else:
                self.setWindowFlags(flags & ~Qt.WindowStaysOnTopHint)
             self.show()

# You can uncomment and run this block for standalone testing of the RefFinderWindow
# if __name__ == "__main__":
#     app = QApplication(sys.argv)
#
#     # Mock parent for stay_on_top checkbox (replace with your actual MainWindow setup if testing)
#     main_window_mock = QMainWindow()
#     main_window_mock.stay_on_top_checkbox = QCheckBox("Stay on Top", main_window_mock)
#     main_window_mock.stay_on_top_checkbox.setChecked(False)
#     main_window_mock.show()
#
#     # Ensure 'clipboard.ico' exists in 'ui' directory for local testing
#     # Or comment out the icon part if you don't have it
#     # Example if you don't have clipboard.ico:
#     # self.process_from_clipboard_button = QPushButton("Process from Clipboard", self)
#
#     ref_finder_win = RefFinderWindow(main_window_mock)
#     ref_finder_win.show()
#     sys.exit(app.exec_())