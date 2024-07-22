import io
import os
import re

from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from StampDrawer import StampDrawer


pdfmetrics.registerFont(TTFont("Times", r"materials\fonts\timesnrcyrmt.ttf"))
output = PdfWriter()  # TODO WTF


class Stamper:
    def __init__(self):
        self.stamp_drawer = StampDrawer()
        self.stamp_paths = []
        self.pdf_paths = {}

    def stamp_directory(self, directory_path, changes, cn_number, cn_date, author, author_signature, error_callback):
        all_is_ok, absent_files, more_sheets, less_sheets = self._check_consistency(directory_path, changes)
        if all_is_ok:
            self._create_stamps(changes, cn_date, cn_number, author, directory_path)
            self._do_stamping(directory_path, changes, author_signature)
            self._delete_stamps()
        else:
            text = ""
            if absent_files:
                absent_files_string = ", ".join(absent_files)
                text += f"Отсутствуют PDF для документов: {absent_files_string}\n"
            if more_sheets:
                more_sheets_string = ", ".join(more_sheets)
                text += f"Для следующих документов листов в PDF больше, чем нужно: {more_sheets_string}\n"
            if less_sheets:
                less_sheets_string = ", ".join(less_sheets)
                text += f"Для следующих документов листов в PDF меньше, чем нужно: {less_sheets_string}\n"

            error_callback("Проблемы с файлами PDF", f"{text}")

    def _check_consistency(self, directory_path, changes):  # TODO check pages
        not_in_directory = []
        more_sheets = []
        less_sheets = []
        ok = True
        for set_code, set_docs in changes.items():
            for doc_code in set_docs:
                contains_doc = False
                actual_filepath = ""
                for dir_filename in os.listdir(directory_path):
                    if dir_filename.endswith(".pdf") and doc_code in dir_filename:
                        contains_doc = True
                        actual_filepath = os.path.join(directory_path, dir_filename)
                        break
                if not contains_doc:
                    not_in_directory.append(doc_code)
                    ok = False
                else:
                    page_number_str = re.findall(r"\d+", actual_filepath.replace(doc_code, ""))[0]
                    page_number = None
                    if page_number_str:
                        page_number = int(page_number_str)
                    if doc_code not in self.pdf_paths.keys():
                        self.pdf_paths[doc_code] = [(actual_filepath, page_number)]
                    else:
                        self.pdf_paths[doc_code].append((actual_filepath, page_number))

            for this_doc_code, actual_filepath_list in self.pdf_paths.items():
                pages = []
                new_pages = []
                cancel_pages = []
                this_set = list(filter(lambda x: this_doc_code in x[1].keys(), changes.items()))[0][0]
                total_pages = int(changes[this_set][this_doc_code]["number_of_sheets"])
                for change in changes[this_set][this_doc_code]["changes"]:
                    if change["change_type"] == "new":
                        new_pages += change["pages"]
                    elif change["change_type"] == "cancel":
                        cancel_pages += change["pages"]
                    else:
                        pages += change["pages"]
                if len(actual_filepath_list) == 1:
                    doc = PdfReader(actual_filepath_list[0][0])
                    if len(doc.pages) > total_pages + len(new_pages) + len(cancel_pages):
                        if this_doc_code not in more_sheets:
                            more_sheets.append(this_doc_code)
                        ok = False
                    elif len(doc.pages) < total_pages + len(new_pages) + len(cancel_pages) \
                            and not len(doc.pages) == len(pages) + len(new_pages) + len(cancel_pages):
                        if this_doc_code not in less_sheets:
                            less_sheets.append(this_doc_code)
                        ok = False
                else:
                    if len(actual_filepath_list) > total_pages + len(new_pages) + len(cancel_pages):
                        if this_doc_code not in more_sheets:
                            more_sheets.append(this_doc_code)
                        ok = False
                    elif len(actual_filepath_list) < total_pages + len(new_pages) + len(cancel_pages)\
                            and not len(actual_filepath_list) == len(pages) + len(new_pages) + len(cancel_pages):
                        if this_doc_code not in less_sheets:
                            less_sheets.append(this_doc_code)
                        ok = False

        return ok, not_in_directory, more_sheets, less_sheets

    def _create_stamps(self, changes, cn_date, cn_number, author, directory_path):
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

    def _do_stamping(self, directory_path, changes, author_signature):
        for doc_code, pdf_path_list in self.pdf_paths.items():
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
                        this_stamp = PdfReader(this_stamp_path)
                        this_stamp.pages[0].scale(scale, scale)

                        if len(doc.pages) >= int(doc_info["number_of_sheets"]):
                            stamped_page = doc.pages[page_number - 1]
                        else:
                            stamped_page = doc.pages[this_doc_page]
                            this_doc_page += 1
                        stamped_page.transfer_rotation_to_content()
                        stamped_page.merge_translated_page(
                            this_stamp.pages[0],
                            stamped_page.cropbox.width - self._to_su(stamp_x),
                            self._to_su(stamp_y))
                        stamped_page.merge_translated_page(
                            self._make_note(f"{set_code}/{doc_info['set_position']}.{page_number}", 3.5).pages[0],
                            stamped_page.cropbox.width - self._to_su(note_x),
                            self._to_su(note_y),
                            over=False
                        )
                        self._sign(
                            stamped_page,
                            author_signature,
                            stamped_page.cropbox.width - self._to_su(stamp_x),
                            self._to_su(stamp_y),
                            scale)
                        output.add_page(doc.pages[page_number - 1])
            else:
                pass  # TODO make process multi-pdf docs
        result_path = os.path.join(directory_path, "result.pdf")
        output_stream = open(result_path, "wb")
        output.write(output_stream)
        output_stream.close()
        self.pdf_paths = {}

    def _make_note(self, text, size):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(self._to_su(70), self._to_su(10)))
        can.setFont("Times", self._to_su(size))
        can.drawString(0, 0, text)
        can.save()
        return PdfReader(packet)

    def _sign(self, page, signature, stamp_x, stamp_y, scale):
        img_temp = io.BytesIO()
        img_doc = canvas.Canvas(img_temp)
        img_doc.drawImage(signature,
                          stamp_x + self._to_su(52),
                          stamp_y + self._to_su(12),
                          self._to_su(7),
                          self._to_su(7),
                          [0, 50, 0, 50, 0, 50]
                          )
        img_doc.save()
        overlay = PdfReader(img_temp).get_page(0)
        page.transfer_rotation_to_content()
        page.merge_page(overlay)

    @staticmethod
    def _to_su(number):
        return number * 72 / 25.4

    def _delete_stamps(self):
        for path in self.stamp_paths:
            os.remove(path)
        self.stamp_paths = []
