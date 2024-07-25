import io
import sys
import os

from pypdf import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

if hasattr(sys, "_MEIPASS"):
    font_path = os.path.join(sys._MEIPASS, r"timesnrcyrmt.ttf")
else:
    font_path = r"materials\fonts\timesnrcyrmt.ttf"

pdfmetrics.registerFont(TTFont("Times", font_path))

styles = getSampleStyleSheet()
styles["Normal"].fontName = "Times"


class StampDrawer:

    def __init__(self):
        self._reset()
        self.table_style = TableStyle([('BACKGROUND', (0, 0), (-1, -1), colors.white),
                                       ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                                       ('ALIGN', (0, 0), (-1, -1), "CENTER"),
                                       ('VALIGN', (0, 0), (-1, -1), "MIDDLE"),
                                       ('FONT', (0, 0), (-1, -1), 'Times', self._to_su(2)),
                                       ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                                       ('GRID', (0, 0), (-1, -1), self._to_su(0.5), colors.black)])

    def draw(self, stamp_type, ii_number, ii_author, ii_date, destination_path, number_of_sections=None):
        stamp_types = {
            "replace": ("-", "Зам."),
            "new": ("-", "Нов."),
            "cancel": ("-", "Анн."),
            "patch": (number_of_sections, "-")
        }
        data = [
            ["Изм.01", *stamp_types[stamp_type], ii_number, ii_date],
            ["Изм", "Кол. уч", "Лист", "№ док.", "Дата"],
            ["Rev.", "Q-ty\nof prt.", "Sheet", "Doc. No", "Date"]
        ]
        table = Table(
            data,
            [self._to_su(7), self._to_su(7), self._to_su(10), self._to_su(16), self._to_su(10)],
            [self._to_su(10), self._to_su(6), self._to_su(6)]
        )
        table.setStyle(self.table_style)
        table.wrapOn(self.can, 0, 0)
        table.drawOn(self.can, self._to_su(0.25), self._to_su(0.22))
        data1 = [
            [" ", " ", " "],
            [ii_author, "Нестеров\nNesterov", "Полянская\nPolianskaia"],
            ["Изм. внес", "Пров.", "Н. контроль"],
            ["Changes made", "Checked by", "Examined by"]
        ]
        table1 = Table(
            data1,
            [self._to_su(17), self._to_su(17), self._to_su(17)],
            [self._to_su(5), self._to_su(5), self._to_su(6), self._to_su(6)]
        )
        table1.setStyle(self.table_style)
        table1.wrapOn(self.can, 0, 0)
        table1.drawOn(self.can, self._to_su(50.25), self._to_su(0.25))

        self.write_result(destination_path)

    def write_result(self, destination):
        self.can.save()
        self.packet.seek(0)
        new_pdf = PdfReader(self.packet)
        self.output.add_page(new_pdf.pages[0])
        output_stream = open(destination, "wb")
        self.output.write(output_stream)
        output_stream.close()
        self._reset()

    def _draw_text(self, text, size, x, y):
        self.can.setFont("Times", self._to_su(size))
        self.can.drawString(self._to_su(x), self._to_su(y), text)

    def _draw_line(self, x1, y1, x2, y2):
        self.can.line(self._to_su(x1), self._to_su(y1), self._to_su(x2), self._to_su(y2))

    def _sign(self, image):
        self.can.drawImage(image, self._to_su(56), self._to_su(17), self._to_su(7), self._to_su(7),
                           [0, 50, 0, 50, 0, 50])

    def _reset(self):
        self.packet = io.BytesIO()
        self.can = canvas.Canvas(self.packet, pagesize=(self._to_su(101.5), self._to_su(22.5)))
        self.output = PdfWriter()

    @staticmethod
    def _to_su(number):
        return number * 72 / 25.4


if __name__ == "__main__":
    signature = r"materials\signature1.png"
    drawer = StampDrawer()
    drawer.draw(
        "replace",
        "03885",
        "Петров\nPetrov",
        "11.11.2011",
        signature
    )
