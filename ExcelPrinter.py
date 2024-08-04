from win32com import client


class ExcelPrinter:

    def __init__(self):
        self.app = None

    def convert(self, xlsx_path):
        if not self.app:
            self._setup()
        wb = self.app.Workbooks.Open(xlsx_path)
        try:
            wb.WorkSheets("sheet1").Select()
        except Exception:
            wb.WorkSheets("CommonList").Select()
        output_path = xlsx_path.replace(".xlsx", "").replace(".xls", "").replace(".XLS", "") + ".pdf"
        wb.ActiveSheet.ExportAsFixedFormat(0, output_path)

    def close(self):
        self.app = None

    def _setup(self):
        self.app = client.Dispatch("Excel.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = False
