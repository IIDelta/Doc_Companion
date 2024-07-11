import win32com

class Macro_RemoveBlankLines:
    def __init__(self):
        self.word_app = win32com.client.gencache.EnsureDispatch('Word.Application')

    def remove_blank_lines(self):
        if not self.word_app.Documents.Count:
            return "No document open in Word."

        selection = self.word_app.Selection
        if selection is None:
            return "No text is selected in Word."

        find = selection.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        find.Text = "^p^p"
        find.Replacement.Text = "^p"
        find.Execute(Replace=2)  # wdReplaceAll

        return "Blank lines removed from the selected text in the open Word document."

    def save_document(self):
        if not self.word_app.Documents.Count:
            return "No document open in Word."
        
        self.word_app.ActiveDocument.Save()
        return "Document saved."

