# ui/removeblanklineswindow.py

# Import necessary components from PyQt5.QtWidgets
from PyQt5.QtWidgets import (QMainWindow, QVBoxLayout, QLabel, QWidget)
from PyQt5.QtCore import QTimer

# Define RemoveBlankLinesWindow class, which is a QMainWindow subclass
class RemoveBlankLinesWindow(QMainWindow):
    # Initialize the class
    def __init__(self, parent=None):
        # Call the superclass's __init__ function
        super().__init__(parent)

        # Set the window title
        self.setWindowTitle("Remove Blank Lines")
        
        # Create a QVBoxLayout to layout the widgets vertically
        self.layout = QVBoxLayout()
        
        # Create a QLabel with no initial text and add it to the layout
        self.label = QLabel("Removing multiple blank lines from selected text in open Word document...", self)
        self.layout.addWidget(self.label)

        # Create a QWidget, set its layout to the QVBoxLayout we created, 
        # and set it as the central widget of the QMainWindow
        self.widget = QWidget()
        self.widget.setLayout(self.layout)
        self.setCentralWidget(self.widget)
        
        # Create a QLabel with no initial text and add it to the layout
        self.label = QLabel("", self)
        self.layout.addWidget(self.label)
        self.remove_blank_lines()

    # Define method to remove blank lines from the selected text in the open Word document
    def remove_blank_lines(self):
        # Import Macro_RemoveBlankLines class from the macros package
        from macros.RemoveBlankLines import Macro_RemoveBlankLines
        macro = Macro_RemoveBlankLines()
        message = macro.remove_blank_lines()

        macro.remove_blank_lines()
        macro.save_document()
        # Set the QLabel text to inform the user that the blank lines were removed
        self.parent().label.setText(message)
