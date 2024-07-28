import os
import tempfile
import pickle
import re
import tkinter as tk
from datetime import datetime, timedelta

from ChangesExtractor import ChangesExtractor


class SaveManager:

    def __init__(self, target):
        self.t = target
        self.extractor = ChangesExtractor()

    def save_settings(self):
        settings = {
            "name_ru": self.t.name_ru_var.get(),
            "last_name_ru": self.t.last_name_ru_var.get(),
            "surname_ru": self.t.surname_ru_var.get(),
            "name_en": self.t.name_en_var.get(),
            "last_name_en": self.t.last_name_en_var.get(),
            "surname_en": self.t.surname_en_var.get(),
            "signature_path": self.t.signature_path_var.get()
        }
        temp_dir = tempfile.gettempdir()
        with open(os.path.join(temp_dir, "scramble.cfg"), "wb") as f:
            pickle.dump(settings, f)

    def restore_settings(self):
        temp_dir = tempfile.gettempdir()
        settings_path = os.path.join(temp_dir, "scramble.cfg")
        if os.path.exists(settings_path):
            with open(settings_path, "rb") as f:
                settings = pickle.load(f)
                self.t.name_ru_var.set(settings["name_ru"])
                self.t.last_name_ru_var.set(settings["last_name_ru"])
                self.t.surname_ru_var.set(settings["surname_ru"])
                self.t.name_en_var.set(settings["name_en"])
                self.t.last_name_en_var.set(settings["last_name_en"])
                self.t.surname_en_var.set(settings["surname_en"])
                self.t.signature_path_var.set(settings["signature_path"])

    def restore_set_settings(self, message_callback):
        set_settings_path = os.path.join(self.t.directory_path_var.get(), "config")
        settings_exist = os.path.exists(set_settings_path)
        if not settings_exist:
            self._fill_variables(set_settings_path)
        else:
            decision = message_callback.askyesno(
                "Пересборка настроек",
                "Для данного комплекта уже существует файл настроек. Пересобрать изменения комплекта?"
            )
            if not decision:
                self._restore_variables(set_settings_path)
                self.t.place_set_entries()
            else:
                self._fill_variables(set_settings_path)

    def save_set_settings(self):
        if self.t.directory_path_var.get() and self.t.changes:
            set_settings_path = os.path.join(self.t.directory_path_var.get(), "config")
            with open(set_settings_path, "wb") as f:
                saved_info = {
                    "set_name": self.t.set_name_var.get(),
                    "change_notice_number": self.t.change_notice_number_var.get(),
                    "change_notice_date": self.t.change_notice_date_var.get(),
                    "changes": self.t.changes
                }
                for set_code in self.t.changes.keys():
                    saved_info[set_code + "_rev"] = getattr(self.t, set_code + "_rev_var").get()
                pickle.dump(saved_info, f)

    def _fill_variables(self, path):
        self.t.changes = self.extractor.extract(self.t.directory_path_var.get())
        self.t.place_set_entries()
        self.t.change_notice_number_var.set(re.findall(r"\d+", path.split("\\")[-3])[0])
        tomorrow = (datetime.now() + timedelta(1)).strftime("%d.%m.%Y")
        self.t.change_notice_date_var.set(tomorrow)
        self.t.ask_for_number_of_changes()

    def _restore_variables(self, path):
        with open(path, "rb") as f:
            restored_info = pickle.load(f)
            self.t.changes = restored_info["changes"]
            self.t.set_name_var.set(restored_info["set_name"])
            self.t.change_notice_number_var.set(restored_info["change_notice_number"])
            self.t.change_notice_date_var.set(restored_info["change_notice_date"])
            for set_code in self.t.changes.keys():
                setattr(self.t, set_code + "_rev_var", tk.StringVar())
                getattr(self.t, set_code + "_rev_var").set(restored_info[set_code + "_rev"])
