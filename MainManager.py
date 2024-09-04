import os
import re
import time
from datetime import datetime

import patoolib
from docx2pdf import convert
from pypdf import PdfReader, PdfWriter

import config
from ChangeTextCreator import ChangeTextCreator
from Stamper import Stamper
from AcadPrinter import AcadPrinter
from ExcelPrinter import ExcelPrinter
from WordPrinter import WordPrinter
from JournalHandler import JournalHandler


class MainManager:

    def __init__(self, target):
        self.t = target
        self.creator = ChangeTextCreator()
        self.stamper = Stamper()
        self.journal = JournalHandler()
        self.excel_printer = ExcelPrinter()
        self.word_printer = WordPrinter()
        self.acad_printer = AcadPrinter()

    def stamp_directory(self, result_path, error_callback, info_callback, check_only=False):
        all_is_ok, absent_files, more_sheets, less_sheets, pdf_paths = self.check_consistency(
            self.t.directory_path_var.get(),
            self.t.changes
        )
        if all_is_ok:
            if check_only:
                info_callback("OK", "Ошибок не обнаружено")
            else:
                change_info = self._get_change_info()
                self.stamper.create_stamps(
                    self.t.changes,
                    self.t.full_changes,
                    change_info,
                    self.t.directory_path_var.get()
                )
                self.stamper.build_change_notice(
                    self.t.directory_path_var.get(),
                    pdf_paths,
                    self.t.changes,
                    self.t.full_changes,
                    self.t.signature_path_var.get(),
                    result_path,
                    bool(self.t.do_stamps_var.get()),
                    bool(self.t.do_notes_var.get()),
                    bool(self.t.do_archive_notes_var.get()),
                    self.t.archive_number_var.get(),
                    self.t.archive_date_var.get()
                )
                self.stamper.delete_stamps()
        else:
            text = self._get_inconsistency_text(absent_files, more_sheets, less_sheets)
            error_callback("Проблемы с файлами PDF", text)

    def check_directory(self, error_callback, info_callback):
        stamped_sheets_pdf_path = os.path.join(self.t.directory_path_var.get(), "result.pdf")
        self.stamp_directory(
            stamped_sheets_pdf_path,
            error_callback,
            info_callback,
            check_only=True
        )

    def create_change_notice(self, error_callback, info_callback, dialog_callback):
        for set_code, set_docs in self.t.changes.items():
            for doc_code, doc_info in set_docs.items():
                if doc_info["page_size"] is None:
                    error_callback("Ошибка", "Есть документы с незаданными форматами. Откройте настройки комплекта и "
                                             "задайте форматы для всех документов")
                    return
        ok, not_in_directory, more_sheets, less_sheets, _ = self.check_consistency(
            self.t.directory_path_var.get(), self.t.changes
        )
        if not ok:
            text = self._get_inconsistency_text(not_in_directory, more_sheets, less_sheets)
            error_callback("Ошибка", text)
            return
        old_pdfs, no_originals = self._directory_has_old_pdfs(self.t.directory_path_var.get())
        if len(old_pdfs) > 0 or len(no_originals) > 0:
            text = self._get_old_pdfs_text(old_pdfs, no_originals)
            if not dialog_callback("Ошибка", text):
                return
        start_time = time.time()
        stamped_sheets_pdf_path = os.path.join(self.t.directory_path_var.get(), "result.pdf")
        self.stamp_directory(
            stamped_sheets_pdf_path,
            error_callback,
            info_callback,
        )
        date_ = self.t.change_notice_date_var.get()
        title_path = os.path.join(self.t.directory_path_var.get(),
                                  f"ИИ {self.t.change_notice_number_var.get()} титул.docx")
        self.creator.create_title(
            {
                "change_notice_date": date_,
                "change_due_date": self._add_months(datetime.strptime(date_, "%d.%m.%Y"), 1).strftime("%d.%m.%Y"),
                "author": self._get_author_string()
            },
            os.path.join(self.t.directory_path_var.get(), "template.docx"),
            title_path
        )
        title_pdf_path = os.path.join(self.t.directory_path_var.get(),
                                      f"ИИ {self.t.change_notice_number_var.get()} титул.pdf")
        convert(title_path, title_pdf_path)
        gnelitskiy = False
        if self.t.approved_var.get() == config.APPROVED_LIST[1]:
            gnelitskiy = True
        signed_title = self.stamper.sign_title(title_pdf_path, self.t.signature_path_var.get(), set_130=gnelitskiy)

        upper_folder_path = "\\".join(self.t.directory_path_var.get().split("\\")[:-1])
        cn_path = os.path.join(upper_folder_path, f"ИИ {self.t.change_notice_number_var.get()}.pdf")
        output_stream = open(cn_path, "wb")
        merger = PdfWriter()
        merger.append(signed_title)
        merger.append(stamped_sheets_pdf_path)
        merger.write(output_stream)
        output_stream.close()
        os.remove(stamped_sheets_pdf_path)
        # os.remove(title_path)
        # os.remove(title_pdf_path)
        self._reduce_pdf_size(cn_path)
        # self._create_originals_archive() TODO error
        end_time = time.time()
        text = self._get_output_text(start_time, end_time)
        info_callback("Успешно!", f"{text}")

    def create_title_template(self):
        template_path = os.path.join(self.t.directory_path_var.get(), "template.docx")
        pdf_path = os.path.join(self.t.directory_path_var.get(), "dummy.pdf")
        sets_plus_revisions = list(
            map(lambda x: x + "_C0" + getattr(self.t, x + "_rev_var").get(), self.t.changes.keys()))
        change_notice_sets = "\n".join(sets_plus_revisions)
        change_notice_info = {"change_notice_number": self.t.change_notice_number_var.get(),
                              "cn": self.t.change_number,
                              "change_notice_sets": change_notice_sets,
                              "set_name": self.t.set_name_var.get(),
                              "attachment_sheets_quantity": self._count_attachments(),
                              "change_notice_date": "{{ change_notice_date }}",
                              "change_due_date": "{{ change_due_date }}",
                              "sheets_total": "{{ sheets_total }}",
                              "author": "{{ author }}",
                              "agreed": self.t.agreed_var.get(),
                              "checked": self.t.checked_var.get(),
                              "examined": self.t.examined_var.get(),
                              "approved": self.t.approved_var.get(),
                              "estimates_ru": self._get_estimates_text()[0],
                              "safety_ru": self._get_safety_text()[0],
                              "estimates_en": self._get_estimates_text()[1],
                              "safety_en": self._get_safety_text()[1],
                              }

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
                    elif len(doc.pages) < total_pages + len(cancel_pages) \
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

    def update_directory_pdfs(self, directory_path):
        pdf_paths = []
        word_paths = []
        excel_paths = []
        dwg_paths = []
        for file in os.listdir(directory_path):
            if file.endswith(".docx") or file.endswith(".doc"):
                word_paths.append(file)
            elif file.endswith(".xlsx") or file.endswith(".xls") or file.endswith(".XLS"):
                excel_paths.append(file)
            elif file.endswith(".dwg"):
                dwg_paths.append(file)
            elif file.endswith(".pdf"):
                pdf_paths.append(file)
        for word_file in word_paths:
            filename = word_file.replace(".docx", "").replace(".doc", "")
            if (not filename.startswith("ИИ")
                    and not filename.startswith("template")
                    and not filename.startswith("title")):
                pdf = [x for x in pdf_paths if x.startswith(filename)]
                if len(pdf) < 1 or len(pdf) == 1 and self._compare_file_mod_times(directory_path, word_file, pdf[0]):
                    self.word_printer.convert(os.path.join(directory_path, word_file))
        for excel_file in excel_paths:
            filename = excel_file.replace(".xlsx", "").replace(".xls", "").replace(".XLS", "")
            pdf = [x for x in pdf_paths if x.startswith(filename)]
            if len(pdf) < 1 or len(pdf) == 1 and self._compare_file_mod_times(directory_path, excel_file, pdf[0]):
                self.excel_printer.convert(os.path.join(directory_path, excel_file))
        for dwg_file in dwg_paths:
            filename = dwg_file.replace(".dwg", "")
            pdf = [x for x in pdf_paths if x.startswith(filename)]
            page_size = self._get_doc_size(filename)
            if len(pdf) == 1 and self._compare_file_mod_times(directory_path, dwg_file, pdf[0]):
                output_path = pdf[0]
                self.acad_printer.convert(os.path.join(directory_path, dwg_file),
                                          os.path.join(directory_path, output_path), page_size)
            elif len(pdf) < 1:
                output_path = filename + ".pdf"
                self.acad_printer.convert(os.path.join(directory_path, dwg_file),
                                          os.path.join(directory_path, output_path), page_size)
        self.word_printer.close()
        self.excel_printer.close()
        self.acad_printer.close_acad()

    def insert_change_notice(self, originals_directory, change_notice_path, error_callback):
        ok, absent_files, files_list = self._check_originals_consistency(originals_directory)
        if not ok:
            error_callback("Ошибка", self._get_absent_originals_text(absent_files))
        else:
            set_start_page = 0
            for set_code, file in files_list:
                this_set_pages = self._insert_change_notice_pages(self.t.changes[set_code],
                                                                  os.path.join(originals_directory, file),
                                                                  change_notice_path, set_start_page)
                set_start_page += this_set_pages

    @staticmethod
    def _compare_file_mod_times(directory_path, file1, file2):
        return os.path.getmtime(os.path.join(directory_path, file1)) \
               > os.path.getmtime(os.path.join(directory_path, file2))

    @staticmethod
    def _get_inconsistency_text(not_in_directory, more_sheets, less_sheets):
        text = ""
        if not_in_directory:
            absent_files_string = ", ".join(not_in_directory)
            text += f"Отсутствуют PDF для документов: {absent_files_string}\n"
        if more_sheets:
            more_sheets_string = ", ".join(more_sheets)
            text += f"Для следующих документов листов в PDF больше, чем нужно: {more_sheets_string}\n"
        if less_sheets:
            less_sheets_string = ", ".join(less_sheets)
            text += f"Для следующих документов листов в PDF меньше, чем нужно: {less_sheets_string}\n"
        return text

    @staticmethod
    def _get_old_pdfs_text(old_pdfs, no_originals):
        text = ""
        if len(old_pdfs) > 0:
            text += "Для следующих файлов файлы PDF созданы позже времени изменения оригинального файла:\n"
            for file in old_pdfs:
                text += file + "\n"
        if len(no_originals) > 0:
            text += "Для следующих файлов PDF нет оригиналов:\n"
            for file in no_originals:
                text += file + "\n"
        text += "Продолжить сборку ИИ?"
        return text

    def _get_estimates_text(self):
        affect_estimates = self.t.estimates_var.get()
        if affect_estimates:
            return ["влияют", "affect"]
        else:
            return ["не влияют", "do not affect"]

    def _get_safety_text(self):
        affect_safety = self.t.safety_var.get()
        if affect_safety:
            return ["влияют", "affect"]
        else:
            return ["не влияют", "do not affect"]

    def _get_author_string(self):
        return self.t.last_name_ru_var.get() + " " + self.t.name_ru_var.get()[0] + "." \
               + self.t.surname_ru_var.get()[0] \
               + ". /\n" + self.t.last_name_en_var.get() + " " + self.t.name_en_var.get()[0] + "." \
               + self.t.surname_en_var.get()[0] + "."

    @staticmethod
    def _get_output_text(start_time, end_time):
        cn_time_minutes = int((end_time - start_time) // 60)
        cn_time_seconds = int((end_time - start_time) % 60)
        output_text = ""
        output_text += "ИИ успешно собрана.\n Время сборки: "
        if cn_time_minutes:
            minutes_remainder = cn_time_minutes % 10
            if minutes_remainder % 10 == 1:
                output_text += f"{cn_time_minutes} минута"
            elif minutes_remainder == 2 or minutes_remainder == 3 or minutes_remainder == 4:
                output_text += f"{cn_time_minutes} минуты"
            else:
                output_text += f"{cn_time_minutes} минут"
        if cn_time_seconds:
            output_text += " "
            seconds_remainder = cn_time_seconds % 10
            if seconds_remainder == 1:
                output_text += f"{cn_time_minutes} секунда"
            elif seconds_remainder == 2 or seconds_remainder == 3 or seconds_remainder == 4:
                output_text += f"{cn_time_minutes} секунды"
            else:
                output_text += f"{cn_time_seconds} секунд"
        return output_text

    def _count_attachments(self):
        num = 0
        for _, set_info in self.t.changes.items():
            for _, document_info in set_info.items():
                for change in document_info["changes"]:
                    num += len(change["pages"])
        return num

    @staticmethod
    def _reduce_pdf_size(path):
        reader = PdfReader(path)
        writer = PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        if reader.metadata is not None:
            writer.add_metadata(reader.metadata)

        for page in writer.pages:
            page.compress_content_streams()

        with open(path, "wb") as fp:
            writer.write(fp)

    @staticmethod
    def _add_months(current_date, months_to_add):
        new_date = datetime(current_date.year + (current_date.month + months_to_add - 1) // 12,
                            (current_date.month + months_to_add - 1) % 12 + 1,
                            current_date.day, current_date.hour, current_date.minute, current_date.second)
        return new_date

    def _directory_has_old_pdfs(self, directory_path):
        originals = []
        pdfs = []
        old_pdfs = []
        for filename in os.listdir(directory_path):
            if (filename.startswith("AKU")
                    and (filename.endswith(".dwg") or filename.endswith(".docx") or filename.endswith(
                        ".xlsx") or filename.endswith(".xls"))
                    and not filename.startswith("title") and not filename.startswith("template")
                    and not filename.startswith("ИИ")):
                originals.append(filename)
            elif filename.endswith(".pdf"):
                pdfs.append(filename)
        for original in originals:
            name = ".".join(original.split(".")[:-1])  # hack
            pdf = list(filter(lambda x: x.startswith(name), pdfs))[0]
            if self._compare_file_mod_times(directory_path, original, pdf):
                old_pdfs.append(original)
            pdfs.remove(pdf)
        filtered_pdfs = list(filter(lambda x: len(x.split("-")) == 2
                                              and not x.startswith("title")
                                              and not x.startswith("template")
                                              and not x.startswith("ИИ"), pdfs))
        return old_pdfs, filtered_pdfs

    def _check_originals_consistency(self, originals_dir):
        absent = []
        ok = True
        files_list = []
        for set_code in self.t.changes.keys():
            exists = False
            filename = ""
            for file in os.listdir(originals_dir):
                if file.split("_")[0] == set_code:
                    exists = True
                    filename = file
                    files_list.append((set_code, filename))
                    break
            if not exists:
                absent.append(filename)
                ok = False
        return ok, absent, files_list

    @staticmethod
    def _get_absent_originals_text(absent_list):
        text = "В указанной папке со сканами отсутствуют сканы следующих комплектов:\n"
        for filename in absent_list:
            text += filename + "\n"
        return text

    def _get_change_info(self):
        change_info = dict()
        change_info[self.t.change_number] = {
            "author": self.t.last_name_ru_var.get() + "\n" + self.t.last_name_en_var.get(),
            "change_notice_date": self.t.change_notice_date_var.get(),
            "change_notice_number": self.t.change_notice_number_var.get(),
        }
        if self.t.change_number > 1:
            for number in range(1, self.t.change_number):
                last_name_ru = getattr(self.t, str(number) + "_last_name_ru_var").get()
                last_name_en = getattr(self.t, str(number) + "_last_name_en_var").get()
                author = last_name_ru + "\n" + last_name_en
                change_info[number] = {
                    "author": author,
                    "change_notice_date": getattr(self.t, str(number) + "_change_notice_date_var").get(),
                    "change_notice_number": getattr(self.t, str(number) + "_change_notice_number_var").get()
                }
        return change_info

    def _prepare_patch(self, set_changes, change_notice_pages_number, start_ch_page):
        start_page_index = change_notice_pages_number - self._count_attachments()
        prepared_patch = {}

        current_cn_page = start_ch_page
        if start_ch_page == 0:
            current_cn_page = start_page_index

        total_correction = 0
        for doc_info in set_changes.values():
            doc_start_index = int(doc_info["set_start_page"]) - 1
            for change in doc_info["changes"]:
                correction = 0
                if change["change_type"] == "new" or change["change_type"] == "cancel":
                    correction = len(change["pages"])
                for page in change["pages"]:
                    prepared_patch[doc_start_index + page + total_correction] = current_cn_page
                    current_cn_page += 1
                total_correction += correction

        return prepared_patch, current_cn_page

    def _insert_change_notice_pages(self, set_changes, original_path, change_notice_path, cn_start_page):
        output = PdfWriter()
        original_doc = PdfReader(original_path)
        change_notice_doc = PdfReader(change_notice_path)

        set_patch, current_cn_page = self._prepare_patch(set_changes, len(change_notice_doc.pages), cn_start_page)

        i = 0
        this_cn_page = cn_start_page
        current_set_page = 0
        while this_cn_page < len(change_notice_doc.pages) and current_set_page < len(original_doc.pages):
            if i in set_patch.keys():
                output.add_page(change_notice_doc.pages[set_patch[i]])
                this_cn_page += 1
                current_set_page += 1
                i += 1
            else:
                output.add_page(original_doc.pages[current_set_page])
                current_set_page += 1
                i += 1

        result_filename = original_path.strip(".pdf") + " + " + change_notice_path.split("/")[-1]
        with open(result_filename, "wb") as f:
            output.write(f)
        return current_cn_page

    def _create_originals_archive(self):
        dir_path = self.t.directory_path_var.get()
        originals = []
        for filename in os.listdir(dir_path):
            if not (filename.endswith(".pdf")
                    or filename.endswith(".bak")
                    or filename.endswith(".dwl")
                    or filename.endswith(".dwl2")
                    or filename.endswith(".db")
                    or filename == "config"
                    or filename == "template.docx"):
                originals.append(filename)
        upper_folder_path = "\\".join(dir_path.split("\\")[:-1])
        archive_path = os.path.join(upper_folder_path, "ИИ " + self.t.change_notice_number_var.get() + ".rar")
        originals_paths = []
        for original in originals:
            originals_paths.append(os.path.join(dir_path, original))
        if os.path.exists(archive_path):
            os.remove(archive_path)
        patoolib.create_archive(archive_path, tuple(originals_paths), verbosity=-1)  # TODO wrong usage

    def _get_doc_size(self, filename):
        doc_size = None
        for set_code, set_info in self.t.changes.items():
            for doc_code, doc_info in set_info.items():
                if doc_code in filename:
                    doc_size = doc_info["page_size"]
                    break
            break
        return doc_size

    def get_journal_info(self):
        change_notice_number = int(self.t.change_notice_number_var.get())
        if change_notice_number:
            info, row = self.journal.get_change_notice_info(change_notice_number)
            if info:
                self.t.estimates_var.set(int(info["estimates"]))
                if int(info["change_number"]) > 1:
                    previous_list = self.journal.get_previous_change_notices_info(info["full_set_code"], row)
                    for i, change_info in enumerate(previous_list):
                        getattr(self.t, f"{i + 1}_last_name_ru_var").set(change_info["last_author"])
                        getattr(self.t, f"{i + 1}_change_notice_date_var").set(change_info["release_date"].strftime("%d.%m.%Y"))
                        getattr(self.t, f"{i + 1}_change_notice_number_var").set(change_info["change_notice_number"])
