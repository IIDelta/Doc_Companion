import win32com.client

class Macro_RemoveMultipleSpaces:
    def __init__(self):
        self.word_app = win32com.client.gencache.EnsureDispatch('Word.Application')

    def remove_multiple_spaces(self):
        if not self.word_app.Documents.Count:
            return "No document open in Word."

        selection = self.word_app.Selection
        if selection is None:
            return "No text is selected in Word."

        while '  ' in selection.Text:
            selection.Text = selection.Text.replace('  ', ' ')
        
        return "Blank spaces removed from the selected text in the open word document"

    def save_document(self):
        if not self.word_app.Documents.Count:
            return "No document open in Word."

        self.word_app.ActiveDocument.Save()
        return "Document Saved"
