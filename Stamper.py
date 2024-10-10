import io
import os
import sys

from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from StampDrawer import StampDrawer
from config import FORMAT_INFO
from tools import chunk_string, pack_change_info

if hasattr(sys, "_MEIPASS"):
    font_path = os.path.join(sys._MEIPASS, r"timesnrcyrmt.ttf")
    nesterov_sign = os.path.join(sys._MEIPASS, r"nesterov_sign.png")
    goncharok_sign = os.path.join(sys._MEIPASS, r"goncharok_sign.png")
    gnelitskiy_sign = os.path.join(sys._MEIPASS, r"gnelitskiy_sign.png")  # TODO now not actual
else:
    font_path = r"materials\fonts\timesnrcyrmt.ttf"
    nesterov_sign = r"materials\nesterov_sign.png"
    goncharok_sign = r"materials\goncharok_sign.png"
pdfmetrics.registerFont(TTFont("Times", font_path))


class Stamper:
    def __init__(self):
        self._reset()
        self.stamp_drawer = StampDrawer()
        self.stamp_paths = []

    def create_stamps(self, changes, full_changes, change_info, directory_path):
        stamp_ids = set()
        for set_code, set_info in changes.items():
            for doc_code, doc_info in set_info.items():
                previous_changes = None
                if set_code in full_changes.keys() and doc_code in full_changes[set_code].keys():
                    previous_changes = full_changes[set_code][doc_code]["changes"]
                for change in doc_info["changes"]:
                    this_change_number = change["change_number"]
                    for page in change["pages"]:
                        this_stamp_info = [self._get_change_id(change, page)]
                        if previous_changes:
                            for previous_change in previous_changes:
                                if previous_change["change_number"] < this_change_number \
                                        and page in previous_change["pages"]:
                                    this_stamp_info.append(self._get_change_id(previous_change, page))
                        this_id = "".join(sorted(this_stamp_info, reverse=True))
                        stamp_ids.add(this_id)
        for stamp_id in stamp_ids:
            stamp_list = chunk_string(stamp_id, 3)
            this_change_info = {}
            for stamp_part in stamp_list:
                this_change_info[int(stamp_part[0])] = pack_change_info(stamp_part, change_info)
            stamp_path = os.path.join(directory_path, stamp_id + ".pdf")
            self.stamp_drawer.draw(this_change_info, stamp_path)
            self.stamp_paths.append(stamp_path)

    def build_change_notice(self, directory_path, pdf_paths, changes, full_changes, author_signature,
                            output_path, do_stamp, do_note, do_archive_note,
                            archive_number, archive_date, previous_set):
        for doc_code, pdf_path_list in pdf_paths.items():
            set_code = list(filter(lambda x: doc_code in x[1], changes.items()))[0][0]
            doc_changes = changes[set_code][doc_code]["changes"]
            previous_changes = None
            if set_code in full_changes.keys() and doc_code in full_changes[set_code].keys():
                previous_changes = full_changes[set_code][doc_code]["changes"]
            doc_info = changes[set_code][doc_code]
            geometry = doc_info["geometry"]
            if len(pdf_path_list) == 1:
                doc = PdfReader(pdf_path_list[0][0])
                this_doc_page = 0
                for doc_change in doc_changes:
                    pages = doc_change["pages"]
                    for page_number in pages:
                        # geometry
                        this_geometry = list(filter(lambda x: x[0] == page_number, geometry))[0][1]
                        stamp_x = this_geometry[0]
                        stamp_y = this_geometry[1]
                        note_x = this_geometry[2]
                        note_y = this_geometry[3]
                        scale = this_geometry[4]
                        # stamp
                        this_stamp_path, number_of_changes = self._resolve_stamp(
                            directory_path, doc_change, page_number, previous_changes
                        )

                        # stamped page
                        if len(doc.pages) >= int(doc_info["number_of_sheets"]):
                            stamped_page = doc.pages[page_number - 1]
                        else:
                            stamped_page = doc.pages[this_doc_page]
                        calibration = self._get_calibration(stamped_page, doc_info["page_size"])

                        if do_stamp:
                            s_x, s_y = self._stamp_page(this_stamp_path, stamped_page, calibration, stamp_x, stamp_y, scale)

                            self._sign(stamped_page, author_signature, s_x, s_y,
                                       self._to_su(50), self._to_su(16 + 10 * (number_of_changes -1)), self._to_su(17), self._to_su(8), scale)

                            self._sign(stamped_page, nesterov_sign, s_x, s_y,
                                       self._to_su(67), self._to_su(16 + 10 * (number_of_changes - 1)), self._to_su(17), self._to_su(8), scale)

                        if do_note and int(doc_info["set_position"]) != 1:
                            note_text = f"{set_code}/{doc_info['set_position']}.{page_number}"
                            self._add_note(stamped_page, note_text, 3.5, note_x, note_y)  # TODO different sizes for notes
                        if do_archive_note and bool(doc_info["has_archive_number"]):
                            self._add_archive_note(stamped_page, archive_number, archive_date, previous_set)  # TODO иногда плохо расставляет, нужно добавить логику с "Номером пакета" в файлах MDB
                        if len(doc.pages) >= int(doc_info["number_of_sheets"]):
                            self.output.add_page(doc.pages[page_number - 1])
                        else:
                            self.output.add_page(doc.pages[this_doc_page])
                            this_doc_page += 1
            else:
                pass  # TODO make process multi-pdf docs

        if os.path.exists(output_path):
            os.remove(output_path)
        with open(output_path, "wb") as f:
            self.output.write(f)
        self._reset()

    def _stamp_page(self, stamp_path, stamped_page, calibration, x, y, scale):
        this_stamp = PdfReader(stamp_path)
        stamped_page.scale_by(calibration)

        this_stamp.pages[0].scale(scale, scale)
        stamped_page.transfer_rotation_to_content()
        stamped_page.merge_translated_page(
            this_stamp.pages[0],
            stamped_page.cropbox.width - self._to_su(x),
            self._to_su(y),
            over=False
        )
        return stamped_page.cropbox.width - self._to_su(x), self._to_su(y)

    def _add_note(self, stamped_page, note_text, font_size, x, y):
        stamped_page.merge_translated_page(
            self._make_note(note_text, font_size).pages[0],
            stamped_page.cropbox.width - self._to_su(x),
            self._to_su(y),
        )

    def _add_archive_note(self, stamped_page, number, date, previous_number):
        number_page = self._make_note(number, self._to_su(2)).pages[0].rotate(270)
        number_page.transfer_rotation_to_content()
        date_page = self._make_note(date, self._to_su(2)).pages[0].rotate(270)
        date_page.transfer_rotation_to_content()
        previous_number_page = self._make_note(previous_number, self._to_su(2)).pages[0].rotate(270)
        previous_number_page.transfer_rotation_to_content()
        stamped_page.merge_translated_page(number_page, self._to_su(9), self._to_su(9))
        stamped_page.merge_translated_page(date_page, self._to_su(9), self._to_su(35))
        stamped_page.merge_translated_page(previous_number_page, self._to_su(9), self._to_su(66))

    def _get_calibration(self, stamped_page, page_size):
        intended_height = self._to_su(FORMAT_INFO[page_size]["height"])
        if page_size.endswith("V"):
            real_height = min(stamped_page.cropbox.width, stamped_page.cropbox.height)
        else:
            real_height = max(stamped_page.cropbox.width, stamped_page.cropbox.height)
        return intended_height / real_height

    def _make_note(self, text, size):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(self._to_su(78), self._to_su(10)))
        can.setFont("Times", self._to_su(size))
        can.drawString(0, 0, text)
        can.save()
        return PdfReader(packet)

    def _resolve_stamp(self, directory_path, doc_change, page_number, previous_changes):
        this_change_number = doc_change["change_number"]
        id_list = [self._get_change_id(doc_change, page_number)]
        if previous_changes:
            for change in previous_changes:
                if change["change_number"] < this_change_number and page_number in change["pages"]:
                    id_list.append(self._get_change_id(change, page_number))
        id_string = "".join(sorted(id_list, reverse=True))
        return os.path.join(directory_path, id_string + ".pdf"), len(id_list)

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

    def sign_title(self, title_path, author_signature, set_130=False):
        doc = PdfReader(title_path)
        page = doc.pages[0]
        # self._sign(page, author_signature, 0, 0,
        #            self._to_su(21.5), self._to_su(6), self._to_su(32), self._to_su(12), 1)
        # self._sign(page, nesterov_sign, 0, 0,
        #            self._to_su(53.5), self._to_su(6), self._to_su(32), self._to_su(12), 1)
        approve_sign = goncharok_sign
        if set_130:
            approve_sign = gnelitskiy_sign
        self._sign(page, approve_sign, 0, 0,
                   self._to_su(112.5), self._to_su(6), self._to_su(32), self._to_su(12), 1)
        return doc

    def _reset(self):
        self.output = PdfWriter()

    @staticmethod
    def _to_su(number):
        return number * 72 / 25.4

    @staticmethod
    def _from_su(number):
        return number * 25.4 / 72

    @staticmethod
    def _get_change_id(change, page):
        change_number = change["change_number"]
        change_type = change["change_type"]
        if change_type == "new":
            return f"{change_number}n0"
        elif change_type == "replace":
            return f"{change_number}r0"
        elif change_type == "cancel":
            return f"{change_number}c0"
        elif change_type == "patch":
            number_of_sections = list(filter(lambda x: x[0] == page, change["sections_number"]))[0][1]
            return f"{change_number}p{number_of_sections}"
        else:
            return

    def delete_stamps(self):
        for path in self.stamp_paths:
            os.remove(path)
        self.stamp_paths = []
