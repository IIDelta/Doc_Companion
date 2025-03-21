# Import necessary modules from PyQt5
from PyQt5.QtWidgets import (QApplication)
import sys
from ui.mainwindow import MainWindow


# This is the main function that starts the application
def main():
    # Create a QApplication, which is necessary for any PyQt application
    app = QApplication(sys.argv)

    # Create an instance of MainWindow and show it
    main_window = MainWindow()
    main_window.show()

    # Start the application's event loop
    sys.exit(app.exec_())


# This block ensures that the main function
# is only called if this script is run directly,
# and not if it's imported as a module
if __name__ == "__main__":
    main()
