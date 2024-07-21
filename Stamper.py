import os

from pypdf import PdfReader, PdfWriter
from StampDrawer import StampDrawer


# directory = r"C:\Users\n.nikiforov\PycharmProjects\IIhelper\3_ИИ"
# pdf_list = []

# for file in os.listdir(directory):
#     filename = os.fsdecode(file)
#     if filename.endswith(".pdf"):
#         pdf_list.append(file)
#     print(file)
# print(pdf_list)

# move to the beginning of the StringIO buffer
# packet.seek(0)

# create a new PDF with Reportlab
# new_pdf = PdfReader(packet)
# stamp = PdfReader(open("stampA4.pdf", "rb"))
# read your existing PDF
# output = PdfWriter()
# add the "watermark" (which is the new pdf) on the existing page
# for pdf_name in pdf_list:
#     pdf_path = os.path.join(directory, pdf_name)
#     existing_pdf = PdfReader(open(pdf_path, "rb"))
#     for page in existing_pdf.pages:
#         if not pdf_name.startswith("!"):
#             page.merge_translated_page(stamp.pages[0], 0, 30)
#         # page.merge_page(new_pdf.pages[0])
#         output.add_page(page)
# # finally, write "output" to a real file
# output_stream = open("destination.pdf", "wb")
# output.write(output_stream)
# output_stream.close()
output = PdfWriter()


class Stamper:
    def __init__(self):
        self.stamp_drawer = StampDrawer()
        self.stamp_paths = []
        self.pdf_paths = []
        self.writer = PdfWriter

    def stamp_directory(self, directory_path, changes, cn_number, cn_date, author, error_callback):
        all_is_ok, absent_files = self._check_consistency(directory_path, changes)
        if all_is_ok:
            self._create_stamps(changes, cn_date, cn_number, author, directory_path)
            self._do_stamping(directory_path, changes)
        else:
            absent_files_string = ", ".join(absent_files)
            error_callback("Нет PDF", f"Отсутствуют PDF для файлов: {absent_files_string}")

    def _check_consistency(self, directory_path, changes):  # TODO check pages
        doc_list = []
        not_in_directory = []
        ok = True
        for set_ in changes.values():
            doc_list += list(set_.keys())
        for doc_code in doc_list:
            filename = doc_code + ".pdf"
            another_filename = doc_code + "-Модель.pdf"
            file_path = os.path.join(directory_path, filename)
            another_filepath = os.path.join(directory_path, another_filename)
            if os.path.exists(file_path):
                self.pdf_paths.append(file_path)
            elif os.path.exists(another_filepath):
                self.pdf_paths.append(another_filepath)
            else:
                not_in_directory.append(doc_code)
                ok = False
        return ok, not_in_directory

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

    def _do_stamping(self, directory_path, changes):
        for pdf_path in self.pdf_paths:
            doc_code = pdf_path.split("\\")[-1].strip(".pdf").strip("-Модель")
            set_code = list(filter(lambda x: doc_code in x[1], changes.items()))[0][0]
            doc_changes = changes[set_code][doc_code]["changes"]
            doc_info = changes[set_code][doc_code]
            geometry = doc_info["geometry"]
            doc = PdfReader(pdf_path)
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
                    print(this_geometry)
                    stamp_x = this_geometry[0]
                    stamp_y = this_geometry[1]
                    # note_x = this_geometry[2]
                    # note_y = this_geometry[3]
                    scale = this_geometry[4]
                    if change_type == "patch":
                        this_stamp_filename = this_stamp_name + str(this_section) + ".pdf"
                    else:
                        this_stamp_filename = this_stamp_name + ".pdf"
                    this_stamp_path = os.path.join(directory_path, this_stamp_filename)
                    this_stamp = PdfReader(this_stamp_path)
                    this_stamp.pages[0].scale(scale, scale)
                    try:
                        if len(doc.pages) == int(doc_info["number_of_sheets"]):
                            doc.pages[page_number - 1].merge_translated_page(
                                this_stamp.pages[0],
                                self._to_su(stamp_x),
                                self._to_su(stamp_y))
                            output.add_page(doc.pages[page_number - 1])
                        else:
                            pass  # TODO
                    except IndexError:
                        pass
        result_path = os.path.join(directory_path, "result.pdf")
        output_stream = open(result_path, "wb")
        output.write(output_stream)
        output_stream.close()

    @staticmethod
    def _to_su(number):
        return number * 27 / 25.4

    def _delete_stamps(self):
        for path in self.stamp_paths:
            os.remove(path)
        self.stamp_paths = []
