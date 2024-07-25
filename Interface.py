import os
import sys
import re
import tkinter as tk
import traceback
from tkinter import scrolledtext, filedialog, messagebox, simpledialog, INSERT
from datetime import datetime, timedelta
import pickle
import tempfile

from transliterate import translit
from docx2pdf import convert
from pypdf import PdfReader, PdfWriter

from ChangesExtractor import ChangesExtractor
from ChangeTextCreator import ChangeTextCreator
from SettingsWindow import SettingsWindow
from Stamper import Stamper


class Interface:

    def __init__(self, settings):
        self.settings = settings
        self.window = tk.Tk()
        self.window.report_callback_exception = self.report_an_error
        if hasattr(sys, "_MEIPASS"):
            self.window.iconbitmap(os.path.join(sys._MEIPASS, r"icon.ico"))
        else:
            self.window.iconbitmap(r"materials\icon.ico")
        self.window.title("ИИшница")
        self.row_for_interface = 9

        self.name_ru_var = tk.StringVar(self.window)
        self.last_name_ru_var = tk.StringVar(self.window)
        self.surname_ru_var = tk.StringVar(self.window)
        self.name_en_var = tk.StringVar(self.window)
        self.last_name_en_var = tk.StringVar(self.window)
        self.surname_en_var = tk.StringVar(self.window)
        self.signature_path_var = tk.StringVar(self.window)

        self.directory_path_var = tk.StringVar(self.window)
        self.change_notice_number_var = tk.StringVar(self.window)
        self.change_notice_date_var = tk.StringVar(self.window)
        self.set_name_var = tk.StringVar(self.window)

        self.extractor = ChangesExtractor()
        self.creator = ChangeTextCreator()
        self.stamper = Stamper()
        self.changes = {}

        self.about_label = tk.Label(self.window, text="Сведения о составителе:")
        self.about_label.grid(sticky="W", row=0, column=0, columnspan=3, padx=7)

        self.last_name_ru_label = tk.Label(self.window, text="Фамилия на русском")
        self.last_name_ru_label.grid(sticky="W", row=1, column=0, padx=7)

        self.name_ru_label = tk.Label(self.window, text="Имя на русском")
        self.name_ru_label.grid(sticky="W", row=1, column=1, padx=7)

        self.surname_ru_label = tk.Label(self.window, text="Отчество на русском")
        self.surname_ru_label.grid(sticky="W", row=1, column=2, padx=7)

        self.last_name_ru_entry = tk.Entry(self.window, width=20, justify='right', textvariable=self.last_name_ru_var)
        self.last_name_ru_entry.grid(sticky="W", row=2, column=0, padx=7)

        self.name_ru_entry = tk.Entry(self.window, width=20, justify='right', textvariable=self.name_ru_var)
        self.name_ru_entry.grid(sticky="W", row=2, column=1, padx=7)

        self.surname_ru_entry = tk.Entry(self.window, width=20, justify='right', textvariable=self.surname_ru_var)
        self.surname_ru_entry.grid(sticky="W", row=2, column=2, padx=7)

        self.last_name_en_label = tk.Label(self.window, text="Фамилия на латинице")
        self.last_name_en_label.grid(sticky="W", row=3, column=0, padx=7)

        self.name_en_label = tk.Label(self.window, text="Имя на латинице")
        self.name_en_label.grid(sticky="W", row=3, column=1, padx=7)

        self.surname_en_label = tk.Label(self.window, text="Отчество на латинице")
        self.surname_en_label.grid(sticky="W", row=3, column=2, padx=7)

        self.last_name_en_entry = tk.Entry(self.window, width=20, justify='right', textvariable=self.last_name_en_var)
        self.last_name_en_entry.grid(sticky="W", row=4, column=0, padx=7)

        self.name_en_entry = tk.Entry(self.window, width=20, justify='right', textvariable=self.name_en_var)
        self.name_en_entry.grid(sticky="W", row=4, column=1, padx=7)

        self.surname_en_entry = tk.Entry(self.window, width=20, justify='right', textvariable=self.surname_en_var)
        self.surname_en_entry.grid(sticky="W", row=4, column=2, padx=7)

        self.signature_path_label = tk.Label(self.window, text="Путь до файла с подписью .png")
        self.signature_path_label.grid(sticky="W", row=5, column=0, columnspan=3, padx=7)

        self.signature_path_entry = tk.Entry(self.window, width=50, justify='right',
                                             textvariable=self.signature_path_var)
        self.signature_path_entry.grid(sticky="W", row=6, column=0, columnspan=2, padx=7)

        self.signature_path_button = tk.Button(self.window, text="Указать", command=self.get_signature)
        self.signature_path_button.grid(sticky="S", row=6, column=2)

        self.directory_path_label = tk.Label(self.window, text="Путь до папки с файлами ИИ:")
        self.directory_path_label.grid(sticky="W", row=7, column=0, columnspan=3, padx=7)

        self.directory_path_entry = tk.Entry(self.window, width=50, justify='right',
                                             textvariable=self.directory_path_var)
        self.directory_path_entry.grid(sticky="W", row=8, column=0, columnspan=2, padx=7)

        self.open_button = tk.Button(self.window, text="Указать", command=self.get_directory_path)
        self.open_button.grid(sticky="S", row=8, column=2)

        self.generate_title_button = tk.Button(self.window, text="Собрать титул", command=self._create_title_template,
                                               state='disabled')
        self.check_consistency_button = tk.Button(self.window, text="Проверить PDF", command=self._check_pdf,
                                                  state='disabled')
        self.generate_change_notice_button = tk.Button(self.window, text="Собрать ИИ",
                                                       command=self._create_change_notice,
                                                       )
        self.result_field = tk.scrolledtext.ScrolledText(self.window, width=45, height=8)
        self.settings_button = tk.Button(self.window, text="Настройки комплекта", command=self._open_settings)

        self.place_rest_of_interface(self.row_for_interface)

        self.directory_path_var.trace("w", self._handle_entry)
        self.signature_path_var.trace("w", self._handle_entry)

        self.window.protocol("WM_DELETE_WINDOW", self.on_exit)
        self._restore_settings()

    def get_signature(self):
        filetypes = [('PNG files', '*.png')]
        f = filedialog.askopenfiles(mode="r", filetypes=filetypes)
        paths = [i.name for i in f]
        self.signature_path_var.set(", ".join(paths))

    def _check_pdf(self):
        stamped_sheets_pdf_path = os.path.join(self.directory_path_var.get(), "result.pdf")
        self.stamper.stamp_directory(
            self.directory_path_var.get(),
            self.changes,
            self.change_notice_number_var.get(),
            self.change_notice_date_var.get(),
            self.last_name_ru_var.get() + "\n" + self.last_name_en_var.get(),
            self.signature_path_var.get(),
            stamped_sheets_pdf_path,
            messagebox.showerror,
            messagebox.showinfo,
            check_only=True
        )

    def _place_set_entries(self):
        row = self.row_for_interface
        self.main_set_name_label = tk.Label(self.window,
                                            text=f"Название основного комплекта\n {list(self.changes.keys())[0]}:")
        self.main_set_name_label.grid(sticky="W", row=row, column=0, columnspan=3, padx=7)
        row += 1
        self.main_set_name_entry = tk.Entry(self.window, width=50, justify='right',
                                            textvariable=self.set_name_var)
        self.main_set_name_entry.grid(sticky="W", row=row, column=0, columnspan=3, padx=7)
        row += 1
        for set_code in self.changes.keys():
            if not hasattr(self, set_code + "_rev_var"):
                setattr(self, set_code + "_rev_var", tk.StringVar())
            setattr(self, set_code + "_rev_label", tk.Label(self.window,
                                                            text=f"Ревизия комплекта {set_code}:"))
            getattr(self, set_code + "_rev_label").grid(sticky="W", row=row, column=0, columnspan=3, padx=7)
            row += 1
            setattr(self, set_code + "_rev_entry", tk.Entry(self.window, width=20, justify='right',
                                                            textvariable=getattr(self, set_code + "_rev_var")))
            getattr(self, set_code + "_rev_entry").grid(sticky="W", row=row, column=1, columnspan=3, padx=7)
            row += 1
        self.change_notice_number_label = tk.Label(self.window, text="Номер ИИ:")
        self.change_notice_number_label.grid(sticky="W", row=row, column=0, padx=7)

        self.change_notice_date_label = tk.Label(self.window, text="Дата ИИ:")
        self.change_notice_date_label.grid(sticky="W", row=row, column=1, padx=7)
        row += 1

        self.change_notice_number_entry = tk.Entry(self.window, width=20, justify='right',
                                                   textvariable=self.change_notice_number_var)
        self.change_notice_number_entry.grid(sticky="W", row=row, column=0, padx=7)

        self.change_notice_date_entry = tk.Entry(self.window, width=20, justify='right',
                                                 textvariable=self.change_notice_date_var)
        self.change_notice_date_entry.grid(sticky="W", row=row, column=1, padx=7)
        row += 1

        self.place_rest_of_interface(row)

    def get_directory_path(self):
        path = os.path.abspath(filedialog.askdirectory())
        self.directory_path_var.set(path)
        if self._restore_set_changes():
            self.changes = self.extractor.extract(self.directory_path_var.get())
            self._place_set_entries()
            self.change_notice_number_var.set(re.findall(r"\d+", path.split("\\")[-2])[0])
            tomorrow = (datetime.now() + timedelta(1)).strftime("%d.%m.%Y")
            self.change_notice_date_var.set(tomorrow)
            self._ask_for_number_of_changes()
        else:
            self._place_set_entries()
        self.print_message(self.changes)

    def place_rest_of_interface(self, row):
        self.generate_title_button.grid(row=row, column=0, pady=10)
        self.check_consistency_button.grid(row=row, column=1, pady=10)
        self.generate_change_notice_button.grid(row=row, column=2, pady=10)
        row += 1
        self.result_field.grid(row=row, columnspan=3, padx=7)
        row += 1
        self.settings_button.grid(row=row, column=2, pady=10, padx=7)

    def _handle_entry(self, *args):
        path = self.directory_path_var.get()
        save_path = self.signature_path_var.get()
        if path and save_path:
            self.generate_title_button.config(state='normal')
            self.generate_change_notice_button.config(state='normal')
            self.settings_button.config(state='normal')
            self.check_consistency_button.config(state='normal')
        else:
            self.generate_title_button.config(state='disabled')
            self.generate_change_notice_button.config(state='disabled')
            self.settings_button.config(state='disabled')
            self.check_consistency_button.config(state='disabled')

    @staticmethod
    def _transliterate(text):
        return translit(text, "ru", reversed=True)

    def print_message(self, text):
        self.result_field.insert(
            INSERT,
            text
        )

    def _save_settings(self):
        settings = {
            "name_ru": self.name_ru_var.get(),
            "last_name_ru": self.last_name_ru_var.get(),
            "surname_ru": self.surname_ru_var.get(),
            "name_en": self.name_en_var.get(),
            "last_name_en": self.last_name_en_var.get(),
            "surname_en": self.surname_en_var.get(),
            "signature_path": self.signature_path_var.get()
        }
        temp_dir = tempfile.gettempdir()
        with open(os.path.join(temp_dir, "scramble.cfg"), "wb") as f:
            pickle.dump(settings, f)

    def _restore_settings(self):
        temp_dir = tempfile.gettempdir()
        settings_path = os.path.join(temp_dir, "scramble.cfg")
        if os.path.exists(settings_path):
            with open(settings_path, "rb") as f:
                settings = pickle.load(f)
                self.name_ru_var.set(settings["name_ru"])
                self.last_name_ru_var.set(settings["last_name_ru"])
                self.surname_ru_var.set(settings["surname_ru"])
                self.name_en_var.set(settings["name_en"])
                self.last_name_en_var.set(settings["last_name_en"])
                self.surname_en_var.set(settings["surname_en"])
                self.signature_path_var.set(settings["signature_path"])

    def _restore_set_changes(self):
        set_settings_path = os.path.join(self.directory_path_var.get(), "config")
        if os.path.exists(set_settings_path):
            decision = messagebox.askyesno(
                "Пересборка настроек",
                "Для данного комплекта уже существует файл настроек. Пересобрать изменения комплекта?"
            )
            if not decision:
                with open(set_settings_path, "rb") as f:
                    restored_info = pickle.load(f)
                    self.changes = restored_info["changes"]
                    self.set_name_var.set(restored_info["set_name"])
                    self.change_notice_number_var.set(restored_info["change_notice_number"])
                    self.change_notice_date_var.set(restored_info["change_notice_date"])
                    for set_code in self.changes.keys():
                        setattr(self, set_code + "_rev_var", tk.StringVar())
                        getattr(self, set_code + "_rev_var").set(restored_info[set_code + "_rev"])
            return decision
        else:
            return True

    def save_set_changes(self):
        if self.directory_path_var.get() and self.changes:
            set_settings_path = os.path.join(self.directory_path_var.get(), "config")
            with open(set_settings_path, "wb") as f:
                saved_info = {
                    "set_name": self.set_name_var.get(),
                    "change_notice_number": self.change_notice_number_var.get(),
                    "change_notice_date": self.change_notice_date_var.get(),
                    "changes": self.changes
                }
                for set_code in self.changes.keys():
                    saved_info[set_code + "_rev"] = getattr(self, set_code + "_rev_var").get()
                pickle.dump(saved_info, f)

    def _ask_for_number_of_changes(self):
        for set_, set_documents in self.changes.items():
            for document_code, document_info in set_documents.items():
                for change in document_info["changes"]:
                    if (change["change_type"] == "patch"
                            and "sections_number" in change.keys()
                            and len(change["sections_number"]) < 1):
                        sections_number = []
                        for page in change["pages"]:
                            number = simpledialog.askinteger(
                                "Число измененных участков",
                                f"Для документа {document_code}, страница {page}, укажите число изменяемых участков:"
                            )
                            sections_number.append((page, number))
                        change["sections_number"] = sections_number

    def _open_settings(self):
        settings_window = SettingsWindow(self)
        self.window.wait_window(settings_window.window)

    def _create_change_notice(self):
        stamped_sheets_pdf_path = os.path.join(self.directory_path_var.get(), "result.pdf")
        self.stamper.stamp_directory(
            self.directory_path_var.get(),
            self.changes,
            self.change_notice_number_var.get(),
            self.change_notice_date_var.get(),
            self.last_name_ru_var.get() + "\n" + self.last_name_en_var.get(),
            self.signature_path_var.get(),
            stamped_sheets_pdf_path,
            messagebox.showerror,
            messagebox.showinfo
        )
        date_ = self.change_notice_date_var.get()
        title_path = os.path.join(self.directory_path_var.get(), "title.docx")
        self.creator.create_title(
            {
                "change_notice_date": date_,
                "change_due_date": self._add_months(datetime.strptime(date_, "%d.%m.%Y"), 1).strftime("%d.%m.%Y"),
                "author": self._get_author_string()
            },
            os.path.join(self.directory_path_var.get(), "template.docx"),
            title_path
        )
        title_pdf_path = os.path.join(self.directory_path_var.get(), "title.pdf")
        convert(title_path, title_pdf_path)
        cn_path = os.path.join(self.directory_path_var.get(), f"ИИ {self.change_notice_number_var.get()}.pdf")
        output_stream = open(cn_path, "wb")
        merger = PdfWriter()
        merger.append(title_pdf_path)
        merger.append(stamped_sheets_pdf_path)
        merger.write(output_stream)
        output_stream.close()

    def _create_title_template(self):
        template_path = os.path.join(self.directory_path_var.get(), "template.docx")
        pdf_path = os.path.join(self.directory_path_var.get(), "dummy.pdf")
        sets_plus_revisions = list(map(lambda x: x + "_C0" + getattr(self, x + "_rev_var").get(), self.changes.keys()))
        change_notice_sets = "\n".join(sets_plus_revisions)
        change_notice_info = {"change_notice_number": self.change_notice_number_var.get(),
                              "change_notice_sets": change_notice_sets,
                              "set_name": self.set_name_var.get(),
                              "attachment_sheets_quantity": self._count_attachments(),
                              "change_notice_date": "{{ change_notice_date }}",
                              "change_due_date": "{{ change_due_date }}",
                              "sheets_total": "{{ sheets_total }}",
                              "author": "{{ author }}"}

        revisions_list = ["_C0" + getattr(self, x + "_rev_var").get() for x in self.changes.keys()]
        self.creator.create(change_notice_info, self.changes, revisions_list, template_path)
        convert(template_path, pdf_path)
        with open(pdf_path, "rb") as f:
            doc = PdfReader(f)
            sheets_total = len(doc.pages)
        os.remove(pdf_path)
        change_notice_info["sheets_total"] = str(sheets_total)
        self.creator.create(change_notice_info, self.changes, revisions_list, template_path)

    @staticmethod
    def _add_months(current_date, months_to_add):
        new_date = datetime(current_date.year + (current_date.month + months_to_add - 1) // 12,
                            (current_date.month + months_to_add - 1) % 12 + 1,
                            current_date.day, current_date.hour, current_date.minute, current_date.second)
        return new_date

    def _count_attachments(self):
        num = 0
        for _, set_info in self.changes.items():
            for _, document_info in set_info.items():
                for change in document_info["changes"]:
                    num += len(change["pages"])
        return num

    def _get_author_string(self):
        return self.last_name_ru_var.get() + " " + self.name_ru_var.get()[0] + "." + self.surname_ru_var.get()[0] \
               + ". /\n" + self.last_name_en_var.get() + " " + self.name_en_var.get()[0] + "." \
               + self.surname_en_var.get()[0] + "."

    @staticmethod
    def report_an_error(*args):
        err = "\n".join(traceback.format_exception(*args))
        messagebox.showerror("Ошибка", err)

    def on_exit(self):
        self._save_settings()
        self.save_set_changes()
        self.window.destroy()

    def run(self):
        self.window.mainloop()
