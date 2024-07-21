from docx2pdf import convert


class PdfPrinter:

    def __init__(self):
        pass

    @staticmethod
    def print_file(self, file_path, result_path):
        convert(file_path, result_path)
