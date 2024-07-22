import os
import tkinter as tk
from tkinter import scrolledtext, filedialog, messagebox, simpledialog, INSERT
from datetime import datetime
import pickle
import tempfile

from transliterate import translit
from docx2pdf import convert
from pypdf import PdfReader

import config
from ChangesExtractor import ChangesExtractor
from ChangeTextCreator import ChangeTextCreator
from SettingsWindow import SettingsWindow
from Stamper import Stamper


class Interface:

    def __init__(self, settings):
        self.settings = settings
        self.window = tk.Tk()
        self.window.title("ИИшница")
        self.window.iconbitmap(config.ICON_PATH)

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

        self.set_name_label = tk.Label(self.window, textvariable=self.set_name_var, wraplength=300)
        self.set_name_label.grid(sticky="W", row=9, column=0, columnspan=3, padx=7)

        self.change_notice_number_label = tk.Label(self.window, text="Номер ИИ:")
        self.change_notice_number_label.grid(sticky="W", row=10, column=0, padx=7)

        self.change_notice_date_label = tk.Label(self.window, text="Дата ИИ:")
        self.change_notice_date_label.grid(sticky="W", row=10, column=1, padx=7)

        self.change_notice_number_entry = tk.Entry(self.window, width=20, justify='right',
                                                   textvariable=self.change_notice_number_var)
        self.change_notice_number_entry.grid(sticky="W", row=11, column=0, padx=7)

        self.change_notice_date_entry = tk.Entry(self.window, width=20, justify='right',
                                                 textvariable=self.change_notice_date_var)
        self.change_notice_date_entry.grid(sticky="W", row=11, column=1, padx=7)

        self.generate_title_button = tk.Button(self.window, text="Собрать титул", command=self._create_title_template,
                                               state='disabled')
        self.generate_title_button.grid(row=12, column=0, pady=10)

        self.generate_change_notice_button = tk.Button(self.window, text="Собрать ИИ",
                                                       command=self._create_change_notice,
                                                       )
        self.generate_change_notice_button.grid(row=12, column=2, pady=10)

        self.result_field = tk.scrolledtext.ScrolledText(self.window, width=45, height=8)
        self.result_field.grid(row=13, columnspan=3, padx=7)

        self.settings_button = tk.Button(self.window, text="Настройки", command=self._open_settings)
        self.settings_button.grid(row=14, column=2, pady=10, padx=7)

        self.directory_path_var.trace("w", self._handle_entry)
        self.signature_path_var.trace("w", self._handle_entry)
        self.last_name_ru_var.trace("w", self._transliterate_last_name)
        self.name_ru_var.trace("w", self._transliterate_name)
        self.surname_ru_var.trace("w", self._transliterate_surname)

        self.window.protocol("WM_DELETE_WINDOW", self.on_exit)
        self._restore_settings()

    def get_signature(self):
        filetypes = [('PNG files', '*.png')]
        f = filedialog.askopenfiles(mode="r", filetypes=filetypes)
        paths = [i.name for i in f]
        self.signature_path_var.set(", ".join(paths))

    def _get_set_name(self):
        set_name = simpledialog.askstring("Название комплекта", "Введите название основного комплекта:")
        self.set_name_var.set(set_name)

    def get_directory_path(self):
        path = os.path.abspath(filedialog.askdirectory())
        self.directory_path_var.set(path)
        if self._restore_set_changes():
            self.changes = self.extractor.extract(self.directory_path_var.get())
            self._get_set_name()
            self._ask_for_number_of_changes()
        self.print_message(self.changes)

    def _handle_entry(self, *args):
        path = self.directory_path_var.get()
        save_path = self.signature_path_var.get()
        if path and save_path:
            self.generate_title_button.config(state='normal')
        else:
            self.generate_title_button.config(state='disabled')

    def _transliterate_last_name(self, *args):
        self.last_name_en_var.set(self._transliterate(self.last_name_ru_var.get()))

    def _transliterate_name(self, *args):
        self.name_en_var.set(self._transliterate(self.name_ru_var.get()))

    def _transliterate_surname(self, *args):
        self.surname_en_var.set(self._transliterate(self.surname_ru_var.get()))

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
        self.stamper.stamp_directory(
            self.directory_path_var.get(),
            self.changes,
            self.change_notice_number_var.get(),
            self.change_notice_date_var.get(),
            self.last_name_ru_var.get() + "\n" + self.last_name_en_var.get(),
            self.signature_path_var.get(),
            messagebox.showerror
        )

    def _create_title_template(self):
        template_path = os.path.join(self.directory_path_var.get(), "template.docx")
        pdf_path = os.path.join(self.directory_path_var.get(), "dummy.pdf")
        change_notice_info = {"change_notice_number": self.change_notice_number_var.get(),
                              "change_notice_sets": "\n".join(self.changes.keys()), "set_name": self.set_name_var.get(),
                              "attachment_sheet_quantity": self._count_attachments(),
                              "change_notice_date": "{{ change_notice_date }}",
                              "change_due_date": "{{ change_due_date }}", "sheets_total": "{{ sheets_total }}",
                              "author": "{{ author }}"}

        self.creator.create(change_notice_info, self.changes, template_path)
        convert(template_path, pdf_path)
        with open(pdf_path, "rb") as f:
            doc = PdfReader(f)
            sheets_total = len(doc.pages)
        os.remove(pdf_path)
        change_notice_info["sheets_total"] = str(sheets_total)
        self.creator.create(change_notice_info, self.changes, template_path)

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

    def on_exit(self):
        self._save_settings()
        self.save_set_changes()
        self.window.destroy()

    def run(self):
        self.window.mainloop()
