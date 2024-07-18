import tkinter as tk
from tkinter import Frame
from tkinter.ttk import Combobox
from tkscrolledframe import ScrolledFrame


class SettingsWindow:

    def __init__(self, master):
        self.master = master
        self.window = tk.Toplevel(master.window)
        self.window.title("Настройки")
        self.sf = ScrolledFrame(self.window, width=1200, height=700)
        self.sf.pack(side="top", expand=1, fill="both")
        self.sf.bind_arrow_keys(self.window)
        self.sf.bind_scroll_wheel(self.window)
        self.in_f = self.sf.display_widget(Frame)

        this_row = 0
        ch = self.master.changes
        self.add_labels_line(this_row, [
            "Код комплекта",
            "Код документа",
            "Номер страницы",
            "Тип изменения",
            "Число изменяемых участков",
            "Координата X пробивки",
            "Координата Y пробивки",
            "Координата X штампа",
            "Координата Y штампа",
            "Масштаб штампа",
        ])
        this_row += 1

        for set_code, set_documents in ch.items():
            set_code_label = set_code + "_label"
            setattr(self, set_code_label, tk.Label(self.in_f, text=set_code))
            getattr(self, set_code_label).grid(row=this_row, column=0)
            this_row += 1
            for doc_code, doc_info in set_documents.items():
                doc_code_label = doc_code + "_label"
                setattr(self, doc_code_label, tk.Label(self.in_f, text=doc_code))
                getattr(self, doc_code_label).grid(row=this_row, column=1)
                this_row += 1
                for change in doc_info["changes"]:
                    for sheet in change["pages"]:
                        self.add_change_line(set_code, doc_code, int(sheet), this_row, 2, change)
                        this_row += 1

        self.ok_button = tk.Button(self.in_f, text="OK", command=self.save_settings)
        self.ok_button.grid(sticky="E", row=this_row, column=0, columnspan=10, padx=7, pady=7)

    def add_labels_line(self, row, labels):
        counter = 0
        for label in labels:
            attr_name = "label_" + str(row) + "_" + str(counter)
            setattr(self, attr_name, tk.Label(self.in_f, text=label))
            getattr(self, attr_name).grid(row=row, column=counter)
            counter += 1

    def add_change_line(self, set_code, document_code, sheet_number, row, column, change):
        this_column = column
        sheet_id = set_code + "%" + document_code + "%" + str(sheet_number)
        setattr(self, sheet_id + "%" + "label", tk.Label(self.in_f, text=str(sheet_number)))
        getattr(self, sheet_id + "%" + "label").grid(row=row, column=this_column)
        this_column += 1
        ch_type = change["change_type"]
        translations = {
            "patch": "Изм.",
            "replace": "Зам.",
            "new": "Нов.",
            "cancel": "Анн."
        }
        tr_ch_type = translations[ch_type]
        change_type_id = set_code + "%" + document_code + "%" + str(sheet_number) + "%" + "change_type"
        setattr(self, change_type_id + "%" + "var", tk.StringVar(self.in_f))
        getattr(self, change_type_id + "%" + "var").set(tr_ch_type)
        setattr(self, change_type_id + "%" + "entry",
                Combobox(self.in_f,
                         width=15,
                         justify="left",
                         textvariable=getattr(self, change_type_id + "%" + "var")))
        getattr(self, change_type_id + "%" + "entry")["values"] = tuple(translations.values())
        getattr(self, change_type_id + "%" + "entry").grid(sticky="W", row=row, column=this_column, padx=7)
        this_column += 1

        if ch_type == "patch":
            num_of_sections = str(list(filter(lambda x: x[0] == sheet_number, change["sections_number"]))[0][1])
        else:
            num_of_sections = "-"
        num_of_sections_id = set_code + "%" + document_code + "%" + str(sheet_number) + "%" + "num_of_sections"
        setattr(self, num_of_sections_id + "%" + "var", tk.StringVar(self.in_f))
        getattr(self, num_of_sections_id + "%" + "var").set(num_of_sections)
        setattr(self, num_of_sections_id + "%" + "entry",
                tk.Entry(self.in_f,
                         width=15,
                         justify="left",
                         textvariable=getattr(self, num_of_sections_id + "%" + "var")))
        getattr(self, num_of_sections_id + "%" + "entry").grid(sticky="W", row=row, column=this_column, padx=7)
        """
        setattr(self, attribute + "_name_var", tk.StringVar(self.window))
        getattr(self, attribute + "_name_var").set(attribute_parameters[0])
        setattr(self, attribute + "_name_entry",
                tk.Entry(self.window,
                         width=15,
                         justify="left",
                         textvariable=getattr(self, attribute + "_name_var")))
        getattr(self, attribute + "_name_entry").grid(sticky="W", row=row, column=1, padx=7)

        setattr(self, attribute + "_regex_var", tk.StringVar(self.window))
        getattr(self, attribute + "_regex_var").set(", ".join(attribute_parameters[2]))
        setattr(self, attribute + "_regex_entry",
                tk.Entry(self.window,
                         width=40,
                         justify="left",
                         textvariable=getattr(self, attribute + "_regex_var")))
        getattr(self, attribute + "_regex_entry").grid(sticky="W", row=row, column=2, padx=7)

        setattr(self, attribute + "_strip_var", tk.StringVar(self.window))
        getattr(self, attribute + "_strip_var").set(attribute_parameters[3])
        setattr(self, attribute + "_strip_entry",
                tk.Entry(self.window,
                         width=5,
                         justify="left",
                         textvariable=getattr(self, attribute + "_strip_var")))
        getattr(self, attribute + "_strip_entry").grid(sticky="W", row=row, column=3, padx=7)

        setattr(self, attribute + "_convert_var", tk.StringVar(self.window))
        getattr(self, attribute + "_convert_var").set(
            json.dumps(attribute_parameters[4], ensure_ascii=False) if attribute_parameters[4] else '')
        setattr(self, attribute + "_convert_entry",
                tk.Entry(self.window,
                         width=10,
                         justify="left",
                         textvariable=getattr(self, attribute + "_convert_var")))
        getattr(self, attribute + "_convert_entry").grid(sticky="W", row=row, column=4, padx=7)

        setattr(self, attribute + "_default_var", tk.StringVar(self.window))
        getattr(self, attribute + "_default_var").set(attribute_parameters[5] if attribute_parameters[5] else '')
        setattr(self, attribute + "_default_entry",
                tk.Entry(self.window,
                         width=10,
                         justify="left",
                         textvariable=getattr(self, attribute + "_default_var")))
        getattr(self, attribute + "_default_entry").grid(sticky="W", row=row, column=5, padx=7)

        setattr(self, attribute + "_capitalize_var", tk.StringVar(self.window))
        getattr(self, attribute + "_capitalize_var").set(attribute_parameters[6])
        setattr(self, attribute + "_capitalize_check",
                tk.Checkbutton(self.window,
                               width=10,
                               justify="left",
                               variable=getattr(self, attribute + "_capitalize_var")))
        getattr(self, attribute + "_capitalize_check").grid(sticky="W", row=row, column=6, padx=7)
        """

    def save_settings(self):
        """
        for attribute in dir(self):
            if attribute.endswith("_verbose_name_var"):
                setting_name = attribute[:-len("_verbose_name_var")]
                value = getattr(self, attribute).get()
                self.settings.parameters[setting_name][1] = value
            elif attribute.endswith("_name_var"):
                setting_name = attribute[:-len("_name_var")]
                value = getattr(self, attribute).get()
                self.settings.parameters[setting_name][0] = value
            elif attribute.endswith("_regex_var"):
                setting_name = attribute[:-len("_regex_var")]
                value = getattr(self, attribute).get()
                if value:
                    self.settings.parameters[setting_name][2] = value.split(", ")
                else:
                    self.settings.parameters[setting_name][2] = []
            elif attribute.endswith("_strip_var"):
                setting_name = attribute[:-len("_strip_var")]
                value = getattr(self, attribute).get()
                self.settings.parameters[setting_name][3] = value
            elif attribute.endswith("_convert_var"):
                setting_name = attribute[:-len("_convert_var")]
                value = getattr(self, attribute).get()
                if value:
                    self.settings.parameters[setting_name][4] = json.loads(value)
                else:
                    self.settings.parameters[setting_name][4] = {}
            elif attribute.endswith("_default_var"):
                setting_name = attribute[:-len("_default_var")]
                value = getattr(self, attribute).get()
                self.settings.parameters[setting_name][5] = value
            elif attribute.endswith("_capitalize_var"):
                setting_name = attribute[:-len("_capitalize_var")]
                value = getattr(self, attribute).get()
                self.settings.parameters[setting_name][3] = bool(value)
            elif attribute == "empty_lines_var":
                value = getattr(self, "empty_lines_var").get()
                self.settings.empty_lines = int(value)
            elif attribute == "ignore_mass_less_than_var":
                value = getattr(self, "ignore_mass_less_than_var").get()
                self.settings.ignore_with_mass_less_than = int(value)
            elif attribute == "ignore_list_var":
                value = getattr(self, "ignore_list_var").get().split(", ")
                self.settings.ignore_with_mass_less_than = value

        self.caller.apply_settings(self.settings)
        self.caller.extractor = Extractor(IOHandler(), self.settings)
        self.settings.save()
        """
        self.on_exit()

    def on_exit(self):
        self.window.destroy()
