import sys
import os
import nltk
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer
from ui.mainwindow import MainWindow

# --- NLTK data path for PyInstaller ---
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    nltk_data_dir = os.path.join(sys._MEIPASS, 'nltk_data')
    if nltk_data_dir not in nltk.data.path:
        nltk.data.path.append(nltk_data_dir)
# --- End NLTK data path for PyInstaller ---

def get_resource_path(relative_path):
    """ Get the absolute path for a resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # If not running as a PyInstaller bundle, use the script's directory
        base_path = os.path.abspath(".") # Use current working dir or os.path.dirname(__file__)

    return os.path.join(base_path, relative_path)

# This is the main function that starts the application
def main():
    # Create a QApplication, which is necessary for any PyQt application
    app = QApplication(sys.argv)

    # --- Load and apply the stylesheet ---
    try:
        # Construct path using the new function
        style_path = get_resource_path(os.path.join('ui', 'style.qss'))
        with open(style_path, "r") as f:
            app.setStyleSheet(f.read())
            print(f"Stylesheet loaded from: {style_path}")
    except FileNotFoundError:
        print(f"Warning: style.qss not found at {style_path}. Using default styles.") # Log the path
    except Exception as e:
        print(f"Error loading stylesheet: {e}")
    # --- End stylesheet ---


    # Create an instance of MainWindow and show it
    main_window = MainWindow()
    main_window.show()

    # Defer loading the acronym list (or any other expensive startup work)
    QTimer.singleShot(0, main_window.prefetch_acronyms)
    # Start the application's event loop
    sys.exit(app.exec_())


# This block ensures that the main function
# is only called if this script is run directly,
# and not if it's imported as a module
if __name__ == "__main__":
    main()