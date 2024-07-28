import os
import re
from datetime import datetime

from docx2pdf import convert
from pypdf import PdfReader, PdfWriter

from ChangeTextCreator import ChangeTextCreator
from Stamper import Stamper


class MainManager:

    def __init__(self, target):
        self.t = target
        self.creator = ChangeTextCreator()
        self.stamper = Stamper()

    def stamp_directory(self, result_path, error_callback, info_callback, check_only=False):
        all_is_ok, absent_files, more_sheets, less_sheets, pdf_paths = self.check_consistency(
            self.t.directory_path_var.get(),
            self.t.changes
        )
        if all_is_ok:
            if check_only:
                info_callback("OK", "Ошибок не обнаружено")
            else:
                author = self.t.last_name_ru_var.get() + "\n" + self.t.last_name_en_var.get()
                self.stamper.create_stamps(
                    self.t.changes,
                    self.t.change_notice_date_var.get(),
                    self.t.change_notice_number_var.get(),
                    author,
                    self.t.directory_path_var.get()
                )
                self.stamper.build_change_notice(
                    self.t.directory_path_var.get(),
                    pdf_paths,
                    self.t.changes,
                    self.t.signature_path_var.get(),
                    result_path,
                    bool(self.t.do_stamps_var.get()),
                    bool(self.t.do_notes_var.get())
                )
                self.stamper.delete_stamps()
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

    def check_directory(self, error_callback, info_callback):
        stamped_sheets_pdf_path = os.path.join(self.t.directory_path_var.get(), "result.pdf")
        self.stamp_directory(
            stamped_sheets_pdf_path,
            error_callback,
            info_callback,
            check_only=True
        )

    def create_change_notice(self, error_callback, info_callback):
        stamped_sheets_pdf_path = os.path.join(self.t.directory_path_var.get(), "result.pdf")
        self.stamp_directory(
            stamped_sheets_pdf_path,
            error_callback,
            info_callback,
        )
        date_ = self.t.change_notice_date_var.get()
        title_path = os.path.join(self.t.directory_path_var.get(), "title.docx")
        self.creator.create_title(
            {
                "change_notice_date": date_,
                "change_due_date": self._add_months(datetime.strptime(date_, "%d.%m.%Y"), 1).strftime("%d.%m.%Y"),
                "author": self._get_author_string()
            },
            os.path.join(self.t.directory_path_var.get(), "template.docx"),
            title_path
        )
        title_pdf_path = os.path.join(self.t.directory_path_var.get(), "title.pdf")
        convert(title_path, title_pdf_path)

        signed_title = self.stamper.sign_title(title_pdf_path, self.t.signature_path_var.get())

        cn_path = os.path.join(self.t.directory_path_var.get(), f"ИИ {self.t.change_notice_number_var.get()}.pdf")
        output_stream = open(cn_path, "wb")
        merger = PdfWriter()
        merger.append(signed_title)
        merger.append(stamped_sheets_pdf_path)
        merger.write(output_stream)
        output_stream.close()

    def create_title_template(self):
        template_path = os.path.join(self.t.directory_path_var.get(), "template.docx")
        pdf_path = os.path.join(self.t.directory_path_var.get(), "dummy.pdf")
        sets_plus_revisions = list(map(lambda x: x + "_C0" + getattr(self.t, x + "_rev_var").get(), self.t.changes.keys()))
        change_notice_sets = "\n".join(sets_plus_revisions)
        change_notice_info = {"change_notice_number": self.t.change_notice_number_var.get(),
                              "change_notice_sets": change_notice_sets,
                              "set_name": self.t.set_name_var.get(),
                              "attachment_sheets_quantity": self._count_attachments(),
                              "change_notice_date": "{{ change_notice_date }}",
                              "change_due_date": "{{ change_due_date }}",
                              "sheets_total": "{{ sheets_total }}",
                              "author": "{{ author }}"}

        revisions_list = ["_C0" + getattr(self.t, x + "_rev_var").get() for x in self.t.changes.keys()]
        self.creator.create(change_notice_info, self.t.changes, revisions_list, template_path)
        convert(template_path, pdf_path)
        with open(pdf_path, "rb") as f:
            doc = PdfReader(f)
            sheets_total = len(doc.pages)
        os.remove(pdf_path)
        change_notice_info["sheets_total"] = str(sheets_total)
        self.creator.create(change_notice_info, self.t.changes, revisions_list, template_path)

    @staticmethod
    def check_consistency(directory_path, changes):
        not_in_directory = []
        more_sheets = []
        less_sheets = []
        pdf_paths = {}
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
                    if doc_code not in pdf_paths.keys():
                        pdf_paths[doc_code] = [(actual_filepath, page_number)]
                    else:
                        pdf_paths[doc_code].append((actual_filepath, page_number))

            for this_doc_code, actual_filepath_list in pdf_paths.items():
                pages = []
                cancel_pages = []
                this_set = list(filter(lambda x: this_doc_code in x[1].keys(), changes.items()))[0][0]
                total_pages = int(changes[this_set][this_doc_code]["number_of_sheets"])
                for change in changes[this_set][this_doc_code]["changes"]:
                    if change["change_type"] == "cancel":
                        cancel_pages += change["pages"]
                    else:
                        pages += change["pages"]
                if len(actual_filepath_list) == 1:
                    doc = PdfReader(actual_filepath_list[0][0])
                    if len(doc.pages) > total_pages + len(cancel_pages):
                        if this_doc_code not in more_sheets:
                            more_sheets.append(this_doc_code)
                        ok = False
                    elif len(doc.pages) < total_pages + len(cancel_pages)\
                            and not len(doc.pages) == len(pages) + len(cancel_pages):
                        if this_doc_code not in less_sheets:
                            less_sheets.append(this_doc_code)
                        ok = False
                else:
                    if len(actual_filepath_list) > total_pages + len(cancel_pages):
                        if this_doc_code not in more_sheets:
                            more_sheets.append(this_doc_code)
                        ok = False
                    elif len(actual_filepath_list) < total_pages + len(cancel_pages) \
                            and not len(actual_filepath_list) == len(pages) + len(cancel_pages):
                        if this_doc_code not in less_sheets:
                            less_sheets.append(this_doc_code)
                        ok = False
        return ok, not_in_directory, more_sheets, less_sheets, pdf_paths

    def _get_author_string(self):
        return self.t.last_name_ru_var.get() + " " + self.t.name_ru_var.get()[0] + "." \
               + self.t.surname_ru_var.get()[0] \
               + ". /\n" + self.t.last_name_en_var.get() + " " + self.t.name_en_var.get()[0] + "." \
               + self.t.surname_en_var.get()[0] + "."

    def _count_attachments(self):
        num = 0
        for _, set_info in self.t.changes.items():
            for _, document_info in set_info.items():
                for change in document_info["changes"]:
                    num += len(change["pages"])
        return num

    def _print_directory(self):
        pass  # TODO

    @staticmethod
    def _add_months(current_date, months_to_add):
        new_date = datetime(current_date.year + (current_date.month + months_to_add - 1) // 12,
                            (current_date.month + months_to_add - 1) % 12 + 1,
                            current_date.day, current_date.hour, current_date.minute, current_date.second)
        return new_date
