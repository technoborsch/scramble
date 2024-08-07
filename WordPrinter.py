from win32com import client


class WordPrinter:

    def __init__(self):
        self.app = None

    def convert(self, doc_path):
        if not self.app:
            self._setup()
        doc = self.app.Documents.Open(doc_path)
        output_path = doc_path.replace(".docx", "").replace(".doc", "") + ".pdf"
        doc.SaveAs(output_path, FileFormat=17)

    def close(self):
        self.app = None

    def _setup(self):
        self.app = client.Dispatch("Word.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = False
