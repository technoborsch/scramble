import io
import os
import sys

from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from StampDrawer import StampDrawer
from config import INITIAL_SIZES, DOC_SIZES_MAP

if hasattr(sys, "_MEIPASS"):
    font_path = os.path.join(sys._MEIPASS, r"timesnrcyrmt.ttf")
    nesterov_sign = os.path.join(sys._MEIPASS, r"nesterov_sign.png")
else:
    font_path = r"materials\fonts\timesnrcyrmt.ttf"
    nesterov_sign = r"materials\nesterov_sign.png"
pdfmetrics.registerFont(TTFont("Times", font_path))
output = PdfWriter()  # TODO WTF


class Stamper:
    def __init__(self):
        self.stamp_drawer = StampDrawer()
        self.stamp_paths = []

    def create_stamps(self, changes, cn_date, cn_number, author, directory_path):
        new_stamp = False
        replace_stamp = False
        cancel_stamp = False
        patch_stamps = set()
        for set_info in changes.values():
            for doc_info in set_info.values():
                for change in doc_info["changes"]:
                    change_type = change["change_type"]
                    if not new_stamp and change_type == "new":
                        new_stamp = True
                    if not replace_stamp and change_type == "replace":
                        replace_stamp = True
                    if not cancel_stamp and change_type == "cancel":
                        cancel_stamp = True
                    if change_type == "patch":
                        change_numbers = [x[1] for x in change["sections_number"]]
                        for number in change_numbers:
                            patch_stamps.add(number)
        if new_stamp:
            new_stamp_path = os.path.join(directory_path, "new_stamp.pdf")
            self.stamp_drawer.draw("new", cn_number, author, cn_date, new_stamp_path)
            self.stamp_paths.append(new_stamp_path)
        if replace_stamp:
            replace_stamp_path = os.path.join(directory_path, "replace_stamp.pdf")
            self.stamp_drawer.draw("replace", cn_number, author, cn_date, replace_stamp_path)
            self.stamp_paths.append(replace_stamp_path)
        if cancel_stamp:
            cancel_stamp_path = os.path.join(directory_path, "cancel_stamp.pdf")
            self.stamp_drawer.draw("cancel", cn_number, author, cn_date, cancel_stamp_path)
            self.stamp_paths.append(cancel_stamp_path)
        if patch_stamps:
            for changes_number in patch_stamps:
                patch_stamp_path = os.path.join(directory_path, f"patch_stamp_{changes_number}.pdf")
                self.stamp_drawer.draw("patch", cn_number, author, cn_date,
                                       patch_stamp_path, number_of_sections=changes_number)
                self.stamp_paths.append(patch_stamp_path)

    def build_change_notice(self, directory_path, pdf_paths, changes, author_signature, output_path, do_stamp, do_note):
        for doc_code, pdf_path_list in pdf_paths.items():
            set_code = list(filter(lambda x: doc_code in x[1], changes.items()))[0][0]
            doc_changes = changes[set_code][doc_code]["changes"]
            doc_info = changes[set_code][doc_code]
            geometry = doc_info["geometry"]
            if len(pdf_path_list) == 1:
                doc = PdfReader(pdf_path_list[0][0])
                this_doc_page = 0
                for doc_change in doc_changes:
                    change_type = doc_change["change_type"]
                    pages = doc_change["pages"]
                    this_stamp_name = ""
                    this_section = 0
                    if change_type == "patch":
                        this_stamp_name = "patch_stamp_"
                    elif change_type == "new":
                        this_stamp_name = "new_stamp"
                    elif change_type == "replace":
                        this_stamp_name = "replace_stamp"
                    elif change_type == "cancel":
                        this_stamp_name = "cancel_stamp"
                    for page_number in pages:
                        if change_type == "patch":
                            sections_number = doc_change["sections_number"]
                            this_section = list(filter(lambda x: int(x[0]) == page_number, sections_number))[0][1]
                        this_geometry = list(filter(lambda x: x[0] == page_number, geometry))[0][1]
                        stamp_x = this_geometry[0]
                        stamp_y = this_geometry[1]
                        note_x = this_geometry[2]
                        note_y = this_geometry[3]
                        scale = this_geometry[4]
                        if change_type == "patch":
                            this_stamp_filename = this_stamp_name + str(this_section) + ".pdf"
                        else:
                            this_stamp_filename = this_stamp_name + ".pdf"
                        this_stamp_path = os.path.join(directory_path, this_stamp_filename)
                        if len(doc.pages) >= int(doc_info["number_of_sheets"]):
                            stamped_page = doc.pages[page_number - 1]
                        else:
                            stamped_page = doc.pages[this_doc_page]
                            this_doc_page += 1
                        calibration = self._get_calibration(stamped_page, doc_code)

                        if do_stamp:
                            s_x, s_y = self._stamp_page(this_stamp_path, stamped_page, calibration, stamp_x, stamp_y, scale)

                            self._sign(stamped_page, author_signature, s_x, s_y,
                                       self._to_su(50), self._to_su(16), self._to_su(17), self._to_su(8), scale)

                            self._sign(stamped_page, nesterov_sign, s_x, s_y,
                                       self._to_su(67), self._to_su(16), self._to_su(17), self._to_su(8), scale)

                        if do_note:
                            note_text = f"{set_code}/{doc_info['set_position']}.{page_number}"
                            self._add_note(stamped_page, note_text, 3.5, note_x, note_y)  # TODO different sizes for notes

                        output.add_page(doc.pages[page_number - 1])
            else:
                pass  # TODO make process multi-pdf docs

        output_stream = open(output_path, "wb")
        output.write(output_stream)
        output_stream.close()

    def _stamp_page(self, stamp_path, stamped_page, calibration, x, y, scale):
        this_stamp = PdfReader(stamp_path)
        stamped_page.scale_by(calibration)

        this_stamp.pages[0].scale(scale, scale)
        stamped_page.transfer_rotation_to_content()
        stamped_page.merge_translated_page(
            this_stamp.pages[0],
            stamped_page.cropbox.width - self._to_su(x),
            self._to_su(y),
        )
        return stamped_page.cropbox.width - self._to_su(x), self._to_su(y)

    def _add_note(self, stamped_page, note_text, font_size, x, y):
        stamped_page.merge_translated_page(
            self._make_note(note_text, font_size).pages[0],
            stamped_page.cropbox.width - self._to_su(x),
            self._to_su(y),
        )

    def _get_calibration(self, stamped_page, doc_code):
        this_doc_letters = DOC_SIZES_MAP[doc_code.split("-")[-1][:3].upper()]
        intended_height = self._to_su(INITIAL_SIZES[this_doc_letters]["height"])
        if this_doc_letters.endswith("V"):
            real_height = min(stamped_page.cropbox.width, stamped_page.cropbox.height)
        else:
            real_height = max(stamped_page.cropbox.width, stamped_page.cropbox.height)
        return intended_height / real_height

    def _make_note(self, text, size):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(self._to_su(70), self._to_su(10)))
        can.setFont("Times", self._to_su(size))
        can.drawString(0, 0, text)
        can.save()
        return PdfReader(packet)

    @staticmethod
    def _sign(page, signature, base_x, base_y, x, y, width, height, scale):
        img_temp = io.BytesIO()
        img_doc = canvas.Canvas(img_temp)
        img_doc.drawImage(signature,
                          0,
                          0,
                          width * scale,
                          height * scale,
                          [200, 255, 200, 255, 200, 255],
                          preserveAspectRatio=True
                          )
        img_doc.save()
        overlay = PdfReader(img_temp).get_page(0)
        page.transfer_rotation_to_content()
        page.merge_translated_page(
            overlay,
            base_x + x * scale,
            base_y + y * scale
        )

    def sign_title(self, title_path, author_signature):
        doc = PdfReader(title_path)
        page = doc.pages[0]
        self._sign(page, author_signature, 0, 0,
                   self._to_su(21.5), self._to_su(6), self._to_su(32), self._to_su(12), 1)
        self._sign(page, nesterov_sign, 0, 0,
                   self._to_su(53.5), self._to_su(6), self._to_su(32), self._to_su(12), 1)
        return doc

    @staticmethod
    def _to_su(number):
        return number * 72 / 25.4

    @staticmethod
    def _from_su(number):
        return number * 25.4 / 72

    def delete_stamps(self):
        for path in self.stamp_paths:
            os.remove(path)
        self.stamp_paths = []
