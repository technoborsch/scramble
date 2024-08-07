import ctypes
import os
import sys
import tkinter as tk
from tkinter import ttk
import traceback
from tkinter import scrolledtext, filedialog, messagebox, simpledialog, INSERT

from SettingsWindow import SettingsWindow
from SaveManager import SaveManager
from MainManager import MainManager
import config


class Interface:

    def __init__(self):
        self.window = tk.Tk()
        self.window.report_callback_exception = self.report_an_error
        if hasattr(sys, "_MEIPASS"):
            self.window.iconbitmap(os.path.join(sys._MEIPASS, r"icon.ico"))
        else:
            self.window.iconbitmap(r"materials\icon.ico")
        self.window.title(f"ИИшница v.{str(config.PROGRAM_VERSION)}")
        self.save_manager = SaveManager(self)
        self.main_manager = MainManager(self)
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
        self.archive_number_var = tk.StringVar(self.window)
        self.archive_date_var = tk.StringVar(self.window)
        self.set_name_var = tk.StringVar(self.window)
        self.agreed_var = tk.StringVar(self.window)
        self.checked_var = tk.StringVar(self.window)
        self.examined_var = tk.StringVar(self.window)
        self.approved_var = tk.StringVar(self.window)
        self.estimates_var = tk.IntVar(self.window)
        self.estimates_var.set(1)
        self.safety_var = tk.IntVar(self.window)
        self.safety_var.set(0)
        self.do_stamps_var = tk.IntVar(self.window)
        self.do_stamps_var.set(1)
        self.do_notes_var = tk.IntVar(self.window)
        self.do_notes_var.set(1)
        self.do_archive_notes_var = tk.IntVar(self.window)
        self.do_archive_notes_var.set(1)

        self.changes = {}

        # Постоянная часть интерфейса - зависит от автора
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
        # То, что относится к текущей ИИ
        self.directory_path_entry = tk.Entry(self.window, width=50, justify='right',
                                             textvariable=self.directory_path_var)
        self.directory_path_entry.grid(sticky="W", row=8, column=0, columnspan=2, padx=7)

        self.open_button = tk.Button(self.window, text="Указать", command=self.get_directory_path)
        self.open_button.grid(sticky="S", row=8, column=2)
        self.main_set_name_label = tk.Label(self.window)
        self.main_set_name_entry = tk.Entry(self.window, width=50, justify='right',
                                            textvariable=self.set_name_var)
        self.change_notice_number_label = tk.Label(self.window, text="Номер ИИ:")
        self.change_notice_date_label = tk.Label(self.window, text="Дата ИИ:")
        self.change_notice_number_entry = tk.Entry(self.window, width=20, justify='right',
                                                   textvariable=self.change_notice_number_var)
        self.change_notice_date_entry = tk.Entry(self.window, width=20, justify='right',
                                                 textvariable=self.change_notice_date_var)
        self.archive_number_label = tk.Label(self.window, text="Архивный номер комплекта:")
        self.archive_date_label = tk.Label(self.window, text="Дата сдачи в архив:")
        self.archive_number_entry = tk.Entry(self.window, width=20, justify='right',
                                                   textvariable=self.archive_number_var)
        self.archive_date_entry = tk.Entry(self.window, width=20, justify='right',
                                                 textvariable=self.archive_date_var)
        self.agreed_label = tk.Label(self.window, text="Согласовано:")
        self.agreed_combobox = ttk.Combobox(self.window, values=config.AGREED_LIST, width=15,
                                            justify='right', textvariable=self.agreed_var)
        self.checked_label = tk.Label(self.window, text="Проверил:")
        self.checked_combobox = ttk.Combobox(self.window, values=config.CHECKED_LIST, width=15,
                                             justify='right', textvariable=self.checked_var)
        self.examined_label = tk.Label(self.window, text="Нормоконтроль:")
        self.examined_combobox = ttk.Combobox(self.window, values=config.EXAMINED_LIST, width=15,
                                             justify='right', textvariable=self.examined_var)
        self.estimates_checkbox = tk.Checkbutton(self.window, text="Влияет на смету",
                                                 variable=self.estimates_var)
        self.safety_checkbox = tk.Checkbutton(self.window, text="Влияет на безопасность",
                                                 variable=self.safety_var)
        # Плавающая нижняя часть интерфейса
        self.generate_title_button = tk.Button(self.window, text="Собрать титул", command=self._create_title_template,
                                               state='disabled')
        self.check_consistency_button = tk.Button(self.window, text="Проверить PDF", command=self._check_directory,
                                                  state='disabled')
        self.update_pdfs_button = tk.Button(self.window, text="Обновить файлы PDF", command=self._update_pdfs,
                                            state='disabled')
        self.generate_change_notice_button = tk.Button(self.window, text="Собрать ИИ",
                                                       command=self._create_change_notice,
                                                       )
        self.insert_change_notice_button = tk.Button(self.window, text="Вставить ИИ в документ",
                                                     command=self._insert_change_notice)
        self.do_stamps_checkbox = tk.Checkbutton(self.window, text="Проставить штампы",
                                                 variable=self.do_stamps_var)
        self.do_notes_checkbox = tk.Checkbutton(self.window, text="Проставить пробивку",
                                                variable=self.do_notes_var)
        self.do_archive_notes_checkbox = tk.Checkbutton(self.window, text="Проставить архивные\nномера",
                                                        variable=self.do_archive_notes_var)
        self.result_field = tk.scrolledtext.ScrolledText(self.window, width=45, height=8)
        self.settings_button = tk.Button(self.window, text="Настройки комплекта", command=self._open_settings)

        self._place_rest_of_interface(self.row_for_interface)

        self.directory_path_var.trace("w", self._handle_buttons_state)
        self.signature_path_var.trace("w", self._handle_buttons_state)

        self.window.protocol("WM_DELETE_WINDOW", self.on_exit)
        self.window.bind_all("<Control-KeyPress>", self.keys)
        self._restore_settings()

    def get_signature(self):
        filetypes = [('PNG files', '*.png')]
        f = filedialog.askopenfiles(mode="r", filetypes=filetypes)
        paths = [i.name for i in f]
        self.signature_path_var.set(", ".join(paths))

    def _check_directory(self):
        self.main_manager.check_directory(messagebox.showerror, messagebox.showinfo)

    def _update_pdfs(self):
        self.main_manager.update_directory_pdfs(self.directory_path_var.get())
        messagebox.showinfo("Успешно", "Файлы PDF в папке ИИ обновлены")

    def _insert_change_notice(self, *args):
        messagebox.showinfo("Папка со сканами", "Укажите путь до папки со сканами комплектов")
        originals_directory = os.path.abspath(filedialog.askdirectory())
        messagebox.showinfo("Файл с ИИ", "Укажите путь до собранного ИИ")
        filetypes = [("Файлы PDF", "*.pdf")]
        f = filedialog.askopenfiles(mode="r", filetypes=filetypes)
        paths = [i.name for i in f]
        self.main_manager.insert_change_notice(originals_directory, paths[0], messagebox.showerror)
        messagebox.showinfo("Успешно", "Файлы с примененными ИИ находятся в папке с оригиналами")

    def place_set_entries(self):
        row = self.row_for_interface
        self.main_set_name_label.grid(sticky="W", row=row, column=0, columnspan=3, padx=7)
        self.main_set_name_label.config(text=f"Название основного комплекта\n {list(self.changes.keys())[0]}:")
        row += 1
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
        self.change_notice_number_label.grid(sticky="W", row=row, column=0, padx=7)

        self.change_notice_date_label.grid(sticky="W", row=row, column=1, padx=7)
        row += 1

        self.change_notice_number_entry.grid(sticky="W", row=row, column=0, padx=7)

        self.change_notice_date_entry.grid(sticky="W", row=row, column=1, padx=7)
        row += 1

        self.archive_number_label.grid(sticky="W", row=row, column=0, padx=7)
        self.archive_date_label.grid(sticky="W", row=row, column=1, padx=7)
        row += 1

        self.archive_number_entry.grid(sticky="W", row=row, column=0, padx=7)
        self.archive_date_entry.grid(sticky="W", row=row, column=1, padx=7)
        row += 1

        self.agreed_label.grid(sticky="W", row=row, column=0, padx=7)
        self.checked_label.grid(sticky="W", row=row, column=1, padx=7)
        self.examined_label.grid(sticky="W", row=row, column=2, padx=7)
        row += 1

        self.agreed_combobox.grid(sticky="W", row=row, column=0, padx=7)
        self.checked_combobox.grid(sticky="W", row=row, column=1, padx=7)
        self.examined_combobox.grid(sticky="W", row=row, column=2, padx=7)
        row += 1

        self.estimates_checkbox.grid(sticky="W", row=row, column=0, padx=7)
        self.safety_checkbox.grid(sticky="W", row=row, column=2, padx=7)
        row += 1

        self._place_rest_of_interface(row)

    def get_directory_path(self):
        path = os.path.abspath(filedialog.askdirectory())
        self.directory_path_var.set(path)
        self._restore_set_changes()
        self._set_approved_variable()
        self._print_message(self.changes)  # TODO rework - more info

    def _place_rest_of_interface(self, row):
        self.generate_title_button.grid(row=row, column=0, pady=10)
        self.update_pdfs_button.grid(row=row, column=1, pady=10)
        self.check_consistency_button.grid(row=row, column=2, pady=10)
        row += 1
        self.generate_change_notice_button.grid(row=row, column=0, pady=10)
        self.insert_change_notice_button.grid(row=row, column=1, pady=10)
        row += 1
        self.do_stamps_checkbox.grid(row=row, column=0, pady=10)
        self.do_notes_checkbox.grid(row=row, column=1, pady=10)
        self.do_archive_notes_checkbox.grid(row=row, column=2, pady=10)
        row += 1
        self.result_field.grid(row=row, columnspan=3, padx=7)
        row += 1
        self.settings_button.grid(row=row, column=2, pady=10, padx=7)

    def _handle_buttons_state(self, *args):
        path = self.directory_path_var.get()
        save_path = self.signature_path_var.get()
        if path and save_path:
            for button in [
                self.generate_title_button,
                self.generate_change_notice_button,
                self.settings_button,
                self.check_consistency_button,
                self.update_pdfs_button,
                self.insert_change_notice_button
            ]:
                button.config(state="normal")
        else:
            for button in [
                self.generate_title_button,
                self.generate_change_notice_button,
                self.settings_button,
                self.check_consistency_button,
                self.update_pdfs_button,
                self.insert_change_notice_button
            ]:
                button.config(state="disabled")

    def _print_message(self, text):
        self.result_field.insert(
            INSERT,
            text
        )

    def _save_settings(self):
        self.save_manager.save_settings()

    def _restore_settings(self):
        self.save_manager.restore_settings()

    def _restore_set_changes(self):
        self.save_manager.restore_set_settings(messagebox)

    def save_set_changes(self):
        self.save_manager.save_set_settings()

    def ask_for_number_of_changes(self):
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
        self.main_manager.create_change_notice(messagebox.showerror, messagebox.showinfo, messagebox.askyesno)

    def _create_title_template(self):
        self.main_manager.create_title_template()

    @staticmethod
    def report_an_error(*args):
        err = "\n".join(traceback.format_exception(*args))
        messagebox.showerror("Ошибка", err)

    def on_exit(self):
        self._save_settings()
        self.save_set_changes()
        self.window.destroy()

    def _set_approved_variable(self):
        code = int(list(self.changes.keys())[0].split(".")[1])
        if code == 120:
            self.approved_var.set(config.APPROVED_LIST[0])
        else:
            self.approved_var.set(config.APPROVED_LIST[1])

    @staticmethod
    def is_ru_lang_keyboard():
        u = ctypes.windll.LoadLibrary("user32.dll")
        pf = getattr(u, "GetKeyboardLayout")
        return hex(pf(0)) == '0x4190419'

    def keys(self, event):
        if self.is_ru_lang_keyboard():
            if event.keycode == 86:
                event.widget.event_generate("<<Paste>>")
            if event.keycode == 67:
                event.widget.event_generate("<<Copy>>")
            if event.keycode == 88:
                event.widget.event_generate("<<Cut>>")
            if event.keycode == 65535:
                event.widget.event_generate("<<Clear>>")
            if event.keycode == 65:
                event.widget.event_generate("<<SelectAll>>")

    def run(self):
        self.window.mainloop()
