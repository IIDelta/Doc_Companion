# ui/removemultiplespaceswindow.py

from PyQt5.QtWidgets import (QMainWindow, QLabel, QVBoxLayout, 
                            QWidget)


class RemoveMultipleSpacesWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.setWindowTitle("Remove Multiple Spaces")
        
        self.layout = QVBoxLayout()
        
        self.label = QLabel("", self)
        self.layout.addWidget(self.label)
        
        self.widget = QWidget()
        self.widget.setLayout(self.layout)
        self.setCentralWidget(self.widget)

        self.label = QLabel("", self)
        self.layout.addWidget(self.label)
        self.remove_multiple_spaces()

    def remove_multiple_spaces(self):
        from macros.RemoveMultipleSpaces import Macro_RemoveMultipleSpaces
        macro = Macro_RemoveMultipleSpaces()
        message = macro.remove_multiple_spaces()
        
        macro.remove_multiple_spaces()
        macro.save_document()

        self.parent().label.setText(message)