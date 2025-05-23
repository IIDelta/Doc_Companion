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

# This is the main function that starts the application
def main():
    # Create a QApplication, which is necessary for any PyQt application
    app = QApplication(sys.argv)

    # --- Load and apply the stylesheet ---
    try:
        # Construct path relative to the script location
        dir_path = os.path.dirname(os.path.realpath(__file__))
        style_path = os.path.join(dir_path, 'ui', 'style.qss')
        with open(style_path, "r") as f:
            app.setStyleSheet(f.read())
            print(f"Stylesheet loaded from: {style_path}")
    except FileNotFoundError:
        print("Warning: style.qss not found. Using default styles.")
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