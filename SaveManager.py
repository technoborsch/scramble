import os
import tempfile
import pickle
import re
import tkinter as tk
from datetime import datetime, timedelta
from copy import deepcopy

from ChangesExtractor import ChangesExtractor
from tools import get_latest_changes, update_extracted_changes_with_saved_changes


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
            self._restore_variables(set_settings_path)
            self.t.place_set_entries()
            # decision = message_callback.askyesno(
            #     "Пересборка настроек",
            #     "Для данного комплекта уже существует файл настроек. Пересобрать изменения комплекта?"
            # )
            # if not decision:
            #     self._restore_variables(set_settings_path)
            #     self.t.place_set_entries()
            # else:
            #     self._fill_variables(set_settings_path)

    def save_set_settings(self):
        if self.t.directory_path_var.get() and self.t.changes:
            set_settings_path = os.path.join(self.t.directory_path_var.get(), "config")
            with open(set_settings_path, "wb") as f:
                saved_info = {
                    "set_name": self.t.set_name_var.get(),
                    "change_notice_number": self.t.change_notice_number_var.get(),
                    "change_notice_date": self.t.change_notice_date_var.get(),
                    "archive_number": self.t.archive_number_var.get(),
                    "archive_date": self.t.archive_date_var.get(),
                    "previous_inventory_number": self.t.previous_inventory_number_var.get(),
                    "agreed": self.t.agreed_var.get(),
                    "checked": self.t.checked_var.get(),
                    "examined": self.t.examined_var.get(),
                    "approved": self.t.approved_var.get(),
                    "estimates": self.t.estimates_var.get(),
                    "safety": self.t.safety_var.get(),
                    "changes": self.t.full_changes,
                    "previous_changes": self.get_previous_changes_info()
                }
                for set_code in self.t.full_changes.keys():
                    saved_info[set_code + "_rev"] = getattr(self.t, set_code + "_rev_var").get()
                pickle.dump(saved_info, f)

    def _fill_variables(self, path):
        self.t.full_changes = self.extractor.extract(self.t.directory_path_var.get())
        self.t.changes = get_latest_changes(self.t.full_changes)
        self.t.place_set_entries()
        self.t.change_notice_number_var.set(re.findall(r"\d+", path.split("\\")[-3])[0])
        tomorrow = (datetime.now() + timedelta(1)).strftime("%d.%m.%Y")
        self.t.change_notice_date_var.set(tomorrow)
        self.t.handle_change_number()
        self.t.set_change_info_vars()

        self.t.ask_for_number_of_changes()

    def _restore_variables(self, path):
        with open(path, "rb") as f:
            original_read_info = pickle.load(f)
            read_info = deepcopy(original_read_info)
            extracted_changes = get_latest_changes(self.extractor.extract(self.t.directory_path_var.get()))
            read_info["changes"] = update_extracted_changes_with_saved_changes(read_info["changes"], extracted_changes)
            restored_info = read_info
            self.t.full_changes = original_read_info["changes"]
            self.t.changes = get_latest_changes(self.t.full_changes)
            for parameter in [
                "set_name",
                "change_notice_number",
                "change_notice_date",
                "archive_number",
                "archive_date",
                "previous_inventory_number",
                "agreed",
                "checked",
                "examined",
                "approved",
                "estimates",
                "safety"
            ]:
                if parameter in restored_info.keys():
                    parameter_name = parameter + "_var"
                    getattr(self.t, parameter_name).set(restored_info[parameter])
            for set_code in self.t.changes.keys():
                setattr(self.t, set_code + "_rev_var", tk.StringVar())
                getattr(self.t, set_code + "_rev_var").set(restored_info[set_code + "_rev"])
            self.t.handle_change_number()
            self.t.set_change_info_vars()
            if "previous_changes" in original_read_info.keys():
                for i in range(self.t.change_number - 1, 0, -1):
                    change = original_read_info["previous_changes"][i]
                    getattr(self.t, f"{i}_change_notice_number_var").set(change["number"])
                    getattr(self.t, f"{i}_change_notice_date_var").set(change["date"])
                    getattr(self.t, f"{i}_last_name_ru_var").set(change["last_name_ru"])
                    getattr(self.t, f"{i}_last_name_en_var").set(change["last_name_en"])

    def get_previous_changes_info(self):
        info = {}
        for i in range(self.t.change_number - 1, 0, -1):
            number = getattr(self.t, f"{i}_change_notice_number_var").get()
            date = getattr(self.t, f"{i}_change_notice_date_var").get()
            last_name_ru = getattr(self.t, f"{i}_last_name_ru_var").get()
            last_name_en = getattr(self.t, f"{i}_last_name_en_var").get()
            info[i] = {
                "number": number,
                "date": date,
                "last_name_ru": last_name_ru,
                "last_name_en": last_name_en
            }
        return info
