# ui/mainwindow.py
# import the necessary classes from PyQt5 for building the GUI
from PyQt5.QtWidgets import (
    QMainWindow,    # the main application window
    QLabel,         # for displaying text or images
    QPushButton,    # for clickable buttons
    QVBoxLayout,    # arranges widgets vertically
    QHBoxLayout,    # arranges widgets horizontally
    QWidget,        # a base class for UI elements; used here at a container
    QCheckBox,      # a checkbox for on/off options
    QMessageBox,    # for showing standard dialog boxes (info, warning, error)
    QFileDialog     # for opening file selection dialogs
)
# for setting icons on windows and buttons
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import (
    Qt,                 # for core functionalities like alignmen
                        # and window flags
    QTimer,             # for scheduled tasks
    QCoreApplication    # for processing events during long tasks
)
# for system interactions, like file paths
import os
# for system specific parameters and functions, like path finding
import sys
# to interact with COM objects, specifically Microsoft Word.
# Allows the app to see the active document
import win32com.client
from .acronymswindow import fetch_acronym_list_online  # Use relative import

# attempt to import the 'process word document' function from the
# CleanDocument macro.
try:
    from macros.CleanDocument import process_word_document
# if this import fails then print a warning and set the function to none
# so we can check if it loaded later.
except ImportError:
    print("Warning: Could not import CleanDocument macro.")
    process_word_document = None


# set the main window class, inheriting from QMainWindow
class MainWindow(QMainWindow):
    # a special decorator that refers to the class in general.
    # and it doesn't need a specific window instance (self)]
    # to work. It's utility is related to MainWindow but
    # but it can be called without creating an actual window.
    @staticmethod
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

    # the constructor method, called when the MainWindow object is created.
    def __init__(self):
        # call the constructor of the parent class (QMainWindow). This allows
        # for the main setup process for QMainWindow module to run first,
        # before some personal touches
        super().__init__()
        # initialize instance variables to None. These will hold references.
        self.active_doc_path = None
        # set the text that appears in the window's title bar.
        self.setWindowTitle("Doc Companion")
        # set a fixed size for the window (width, height)
        self.setFixedSize(400, 250)

        # --- Icon Paths ---
        # gets the paths of the various icons using the helper function
        icon_path = self.get_resource_path("leaf.ico")
        replace_icon_path = self.get_resource_path("replace.ico")
        acronym_icon_path = self.get_resource_path("acronym.ico")
        clean_icon_path = self.get_resource_path("clean.ico")
        # --- End Icon Paths ---

        # Set the main window icon (title bar and taskbar)
        self.setWindowIcon(QIcon(icon_path))
        # creates a vertical box layout. Helps to stack widgets on top
        # eachother
        self.layout = QVBoxLayout()
        # set margins (left, top, right, bottom) around the layout
        self.layout.setContentsMargins(15, 15, 15, 15)
        # sets the amount of space between the widgets
        self.layout.setSpacing(10)

        # --- Top Bar (Label & Checkbox) ---
        # create a horizontal  box layout. This arranges widgets side
        # by side
        top_bar_layout = QHBoxLayout()
        # create a label to show the active document name
        # use 'self' to show that this label belongs to the MainWindow
        self.active_doc_label = QLabel("Active Word Document: None", self)
        # add the label to the horizontal layout. It will be on the left.
        top_bar_layout.addWidget(self.active_doc_label)
        # add a stretchable space. Pushes everything
        # added after it to the right
        top_bar_layout.addStretch()
        # create a checkbox with the text 'stay on top'
        self.stay_on_top_checkbox = QCheckBox("Stay on Top", self)
        # connect the 'stateChanged' signal of the checkbox  to the
        # toggle stay on top method.
        self.stay_on_top_checkbox.stateChanged.connect(self.toggle_stay_on_top)
        # add the checkbox to the horizontal top bar layout
        top_bar_layout.addWidget(self.stay_on_top_checkbox)
        # add this entire horizontal top bar layout to our main vertical
        # layout.
        # This whole row now becomes the first item in the vertical stack.
        self.layout.addLayout(top_bar_layout)
        # --- End Top Bar ---

        # create the replace values button, pass it to the QIcon object
        # a space, and the button text.
        self.replace_values_selection_button = QPushButton(
            QIcon(replace_icon_path), " Replace Values", self)
        # connect its 'clicked' signal (emitted when the button is pressed)
        # to the method which opens that part of the program's ui
        self.replace_values_selection_button.clicked.connect(
            self.open_replace_values_selection_window)
        # add this button to our main vertical layout
        # to appear under the top bar.
        self.layout.addWidget(self.replace_values_selection_button)

        # create the acronyms button
        self.acronyms_button = QPushButton(
            QIcon(acronym_icon_path), " Acronyms Table", self)
        # opens the acronyms ui when the button is clicked
        self.acronyms_button.clicked.connect(self.open_acronyms_window)
        # add the button the vertical layout
        self.layout.addWidget(self.acronyms_button)

        # create the clean doc button
        self.clean_doc_button = QPushButton(
            QIcon(clean_icon_path), " Clean && Protect Document", self)
        # when this button is pressed launch the clean document
        # part of the program directly (currently no GUI)
        self.clean_doc_button.clicked.connect(self.run_clean_document)
        # add the button to the layout
        self.layout.addWidget(self.clean_doc_button)

        # Add a stretchable space. Pushes the label down towards the bottom
        # of the window, filling any vertical space above
        self.layout.addStretch()

        # create a status label. Currently empty.
        # used later to display status messages like 'Processing'
        self.label = QLabel("", self)
        # Center the text within the label.
        self.label.setAlignment(Qt.AlignCenter)
        # add the label to the layout (it will be at the bottom)
        self.layout.addWidget(self.label)

        # Create a basic QWidget.
        self.widget = QWidget()
        # set our main vertical layout onto this widget
        # This contains the top bar, buttons, stretch, and label
        self.widget.setLayout(self.layout)
        # sets this as the central widget.
        self.setCentralWidget(self.widget)

        # create our stopwatch
        self.timer = QTimer()
        # connect our stopwatch to our function for updating
        # the active document
        self.timer.timeout.connect(self.update_active_document)
        # start the timer, and set it's signal to emit every
        # 1000 milliseconds or 1 second
        self.timer.start(1000)

        # we initialize variables to hold references to our
        # other windows
        self.acronyms_window = None
        self.replace_values_window = None
        # call the function to update our active document once
        # immediately when the window starts
        self.update_active_document()

    def update_active_document(self):
        """
        For detecting the currently active word document.
        We use try...except because interacting with other programs
        can be error prone.
        """
        try:
            # win32com magic, .GetActiveObject tries to grab
            # the currently running word doc
            word_app = win32com.client.GetActiveObject('Word.Application')
            # if successful, it stores that active document
            active_doc = word_app.ActiveDocument
            # then we update the text in the main window to show the name
            # of that active document
            self.active_doc_label.setText(
                "Active Document: " + active_doc.Name)
            # we also store the full path of the active document
            # in case we need to use it later
            self.active_doc_path = os.path.join(
                active_doc.Path, active_doc.Name)
        # if there are any errors in the try block...
        except Exception:
            # we set the label as none and clears the path.
            # prevents the program from crashing if word is not ready
            self.active_doc_label.setText("Active Document: None")
            self.active_doc_path = None

    def open_replace_values_selection_window(self):
        # Open the replace values ui.
        from .replacevalues_selectionwindow import ReplaceValuesSelectionWindow
        # check if we already have the window stored and if it is
        # currently visible to teh user
        if (self.replace_values_window is not None and
                self.replace_values_window.isVisible()):
            # if the window exists but it's just not shown
            # then we show it to the user
            # stops us from opening multiple instances of the same window.
            self.replace_values_window.activateWindow()
        else:
            # if it doesn't exist then we create one. With the main window
            # as the parent
            self.replace_values_window = ReplaceValuesSelectionWindow(self)
            # show the newly created window
            self.replace_values_window.show()

    def open_acronyms_window(self):
        # opens the acronym window, same stuff
        from .acronymswindow import AcronymsWindow
        if (self.acronyms_window is not None and
                self.acronyms_window.isVisible()):
            self.acronyms_window.activateWindow()
        else:
            self.acronyms_window = AcronymsWindow(self)
            self.acronyms_window.show()

    def run_clean_document(self):
        """
        Runs the clean document macro.
        First we check if the process word document function loaded.
        If it did not, it means the import failed.
        """
        if process_word_document is None:
            # if it failed show the error message. Self makes it appear
            # centered over the main window
            QMessageBox.critical(self,
                                 "Error",
                                 "Clean Document module failed to load.")
            # this stops the function from going any further
            return

        # set some options for the file dialog
        options = QFileDialog.Options()

        # open a standard open file dialog
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Choose a Word Document to Clean", "",
            "Word Documents (*.docx *.doc)", options=options)

        # checks if the file path is empty
        if not file_path:
            # show a temporary message
            self.label.setText("File selection cancelled.")
            QTimer.singleShot(3000, self.label.clear)
            # stop the function
            return

        # for safety we wrap the main processing in try...except...finally
        try:
            # update the label to show the work is being done
            self.label.setText(
                f"Processing {os.path.basename(file_path)}... Please wait.")
            # This should show the events as they come in,
            # instead of at the end. Though it is not working
            QCoreApplication.processEvents()

            # call the actual 'macro' function
            # it returns success (true or false)
            # and a list of messages (the logs)
            success, messages = process_word_document(file_path)

            # join all the log messages with newline characters
            log_message = "\n".join(messages)
            # if the macro reported success
            if success:
                # show an information message box in the log
                QMessageBox.information(self,
                                        "Success",
                                        f"Processing finished.\n\nLog:\n{
                                            log_message}")
            # otherwise if it failed
            else:
                # we show the error message
                QMessageBox.critical(self,
                                     "Error",
                                     f"Processing failed.\n\nLog:\n{
                                        log_message}")

        # if any unexpected error happens during the try block
        except Exception as e:
            # show the critical error message with the details of the error
            QMessageBox.critical(self,
                                 "Error",
                                 f"An unexpected error occurred: {str(e)}")
        # the finally block always runs
        finally:
            # ensure oru status label is cleared after whatever happened
            self.label.clear()

    # this method handles the stay on top checkbox state
    def toggle_stay_on_top(self, checked):
        # get the current set of window flags
        # like if there is a title bar, minimize button, etc.
        flags = self.windowFlags()

        # if the checkbox is checked
        if checked:
            # we add a new flag which is the stay on top hint
            # while leaving all other flags alone
            self.setWindowFlags(flags | Qt.WindowStaysOnTopHint)
        else:
            # We want to *remove* the 'WindowStaysOnTopHint' flag.
            # We use bitwise AND ('&') and NOT ('~').
            # '~Qt.WindowStaysOnTopHint' creates a 'mask' where every
            # switch is ON *except* the StayOnTop one.
            # '&' means "keep only the switches that are ON in *both*
            # the current flags *and* the mask." This effectively turns
            # *off* only the StayOnTop switch.
            self.setWindowFlags(flags & ~Qt.WindowStaysOnTopHint)

        # we show the window again to make sure the flags actually update
        self.show()

        # if any of our other two windows are open and visible currently
        # to the user, also apply the stay on top to them
        if (self.replace_values_window is not None and
                self.replace_values_window.isVisible()):
            self.replace_values_window.update_stay_on_top()
        if (self.acronyms_window is not None and
                self.acronyms_window.isVisible()):
            self.acronyms_window.update_stay_on_top()

    def prefetch_acronyms(self):
        # to prefetch our acronym list, wrapped in try except in case it fails
        # like if the internet is down or anything
        try:
            # the URL where the acronym list is found
            url = (
                "https://raw.githubusercontent.com/IIDelta/Doc_Companion/"
                "main/acronyms/acronym%20list.txt"
            )
            # the path to the cached directory for the
            # user-defined acronym list
            cache_dir = os.path.join(os.path.expanduser("~"), ".doc_companion")
            cache = os.path.join(cache_dir, "base_acronym_list.txt")
            # create the cache directory if it doesn't already exist
            os.makedirs(cache_dir, exist_ok=True)
            # we pass the fetched acronym list url and local cache
            # to the function that needs that information
            fetch_acronym_list_online(url, cache)
            print("Acronym list prefetched.")
        # print the error if failed
        except Exception as e:
            print(f"Prefetch failed: {e}")
