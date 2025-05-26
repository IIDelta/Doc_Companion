"""
Initialization
    It sets up the core QApplication.
Resource Handling
    It includes a smart way (get_resource_path) to find files
    (like icons and stylesheet) whether you're running the code directly
    or as a packaged application.
Styling
    It tries to load a style.qss file to give the application a custom
    look.
Window Creation
    It creates and shows the MainWindow.
Deferred Loading
    It uses QTimer to load some data after the window appears,
    which improves startup time.
Event Loop
    app.exec_() is crucial. It keeps the window open and allow
    it to respond to user actions (clicks, key presses, etc.).
"""
# sys - for interacting with the python interpreter
import sys
# os - for interacting with the OS
import os
# nltk - natural language toolkit. For text processing
import nltk
# pyQT5 - a library for creating GUIs
# QApplication manages the GUI's app control flow and main settings
from PyQt5.QtWidgets import QApplication
# QTimer - used for scheduled events, tasks to run periodically/later
from PyQt5.QtCore import QTimer
# Imports my main window module
from ui.mainwindow import MainWindow

# --- NLTK data path for PyInstaller ---
# This block checks if the program is running bundled as a .exe
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    # if the program is bundled, it finds the path for the nltk data
    nltk_data_dir = os.path.join(sys._MEIPASS, 'nltk_data')
    # then we add this path to nltk's list of data paths
    # so nltk can find it's necessary files
    if nltk_data_dir not in nltk.data.path:
        nltk.data.path.append(nltk_data_dir)
# --- End NLTK data path for PyInstaller ---


def get_resource_path(relative_path):
    """
    Get the absolute path for a resource (like icons or stylesheets).
    This function handles both running the script directly (development)
    and running it as a bundled PyInstaller executable.
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        # this is the 'base_path' when running as a bundled app.
        base_path = sys._MEIPASS
    except AttributeError:
        # if '_MEIPASS' does not exist, we are likely running in a dev
        # environment. So we can use the current working directory or the
        # script's directory as 'base_path'
        base_path = os.path.abspath(".")
    # Joins the base path with the relatvie path provided
    # to get the full path to the resource
    return os.path.join(base_path, relative_path)


# This is the main function that starts the application
def main():
    # Create a QApplication, which is necessary for any PyQt application
    # sys.arg allows the passage of command-line arguments to the app
    app = QApplication(sys.argv)

    # --- Load and apply the stylesheet ---
    try:
        # Construct the full path to teh 'style.qss' file (qt style sheet)
        # using the get_resource_path function
        style_path = get_resource_path(os.path.join('ui', 'style.qss'))
        # open the stylesheet as a file
        with open(style_path, "r") as f:
            # read the content of the style sheet and apply it
            # to the entire application
            app.setStyleSheet(f.read())
            # print a message to confirm loading (for debugging)
            print(f"Stylesheet loaded from: {style_path}")

    except FileNotFoundError:
        # if the stylesheet isn't found, print a warning but continue running
        # with the default styles
        print(f"Warning: style.qss not found at {style_path}."
              "Using default styles.")
        # Log the path
    except Exception as e:
        print(f"Error loading stylesheet: {e}")
    # --- End stylesheet ---

    # Create an instance of my custom main window
    main_window = MainWindow()
    # display the main window
    main_window.show()

    """
    Defer loading the acronym list or anything else
    Qtimer is used to schedule the pre-fetch acronyms method
    to run as soon as teh app's event loop starts with no delay.
    This prevents the potentially slow task from blocking the initial display
    of the main window.
    """
    QTimer.singleShot(0, main_window.prefetch_acronyms)
    # this ensure a clean exit when the app finishes
    sys.exit(app.exec_())


"""
A standard python construct. CHecks if this script is being run directly by
the user. If it is, it calls the main function to start the app.
If this script were imported as a module into another script,
this would not run.
"""
if __name__ == "__main__":
    main()
