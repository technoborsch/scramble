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

    def draw(self, changes_info_dict, destination_path):
        template = {
            1: {
                "stamp_type": "",
                "ii_number": "",
                "ii_author": "",
                "ii_date": "",
                "number_of_sections": ""
            }
        }
        data = []
        data1 = []
        table_heights = []
        table1_heights = []
        for change_number, change_info in dict(sorted(changes_info_dict.items(), reverse=True)).items():
            stamp_type = change_info["stamp_type"]
            ii_number = change_info["ii_number"]
            ii_author = change_info["ii_author"]
            ii_date = change_info["ii_date"]
            number_of_sections = change_info["number_of_sections"]
            stamp_types = {
                "replace": ("-", "Зам."),
                "new": ("-", "Нов."),
                "cancel": ("-", "Анн."),
                "patch": (number_of_sections, "-")
            }
            data.append([f"Изм.0{change_number}", *stamp_types[stamp_type], ii_number, ii_date])
            data1.append([" ", " ", " "])
            data1.append([ii_author, "Нестеров\nNesterov", "Полянская\nPolianskaia"])
            table_heights.append(self._to_su(10))
            table1_heights.append(self._to_su(5))
            table1_heights.append(self._to_su(5))

        data.append(["Изм", "Кол. уч", "Лист", "№ док.", "Дата"]),
        data.append(["Rev.", "Q-ty\nof prt.", "Sheet", "Doc. No", "Date"])
        data1.append(["Изм. внес", "Пров.", "Н. контроль"])
        data1.append(["Changes made", "Checked by", "Examined by"])
        table_heights.append(self._to_su(6))
        table_heights.append(self._to_su(6))
        table1_heights.append(self._to_su(6))
        table1_heights.append(self._to_su(6))

        table = Table(
            data,
            [self._to_su(7), self._to_su(7), self._to_su(10), self._to_su(16), self._to_su(10)],
            table_heights
        )

        table.setStyle(self.table_style)
        table.wrapOn(self.can, 0, 0)
        table.drawOn(self.can, self._to_su(0.25), self._to_su(0.22))
        table1 = Table(
            data1,
            [self._to_su(17), self._to_su(17), self._to_su(17)],
            table1_heights
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

    def _reset(self):
        self.packet = io.BytesIO()
        self.can = canvas.Canvas(self.packet)
        self.output = PdfWriter()

    @staticmethod
    def _to_su(number):
        return number * 72 / 25.4
