import tkinter as tk
from tkinter.ttk import Combobox, Checkbutton
from tkinter import Frame
from tkscrolledframe import ScrolledFrame

from AdditionalWindow import AdditionalWindow
from config import SIZES_COORDINATES, DOC_SIZES_MAP


class SettingsWindow(AdditionalWindow):

    def __init__(self, master):
        super().__init__(master)
        self.window.title("Настройки комплекта")
        self.sf = ScrolledFrame(self.window, width=1470, height=700)
        self.sf.pack(side="top", expand=1, fill="both")
        self.sf.bind_arrow_keys(self.window)
        self.sf.bind_scroll_wheel(self.window)
        self.in_f = self.sf.display_widget(Frame)

        this_row = 0
        self.add_labels_line(this_row, [
            "Код комплекта",
            "Код документа",
            "Номер\nстраницы",
            "Тип\nизменения",
            "Число\nизменяемых\nучастков",
            "Описание\nна\nрусском",
            "Описание\nна\nанглийском",
            "Формат",
            "Координата\nX\nпробивки",
            "Координата\nY\nпробивки",
            "Координата\nX\nштампа",
            "Координата\nY\nштампа",
            "Масштаб\nштампа",
            "Архивник",
        ], self.in_f)
        this_row += 1

        for set_code, set_documents in self.master.changes.items():
            set_code_label = set_code + "_label"
            setattr(self, set_code_label, tk.Label(self.in_f, text=set_code))
            getattr(self, set_code_label).grid(row=this_row, column=0)
            this_row += 1
            for doc_code, doc_info in set_documents.items():
                doc_code_label = doc_code + "_label"
                setattr(self, doc_code_label, tk.Label(self.in_f, text=doc_code))
                getattr(self, doc_code_label).grid(row=this_row, column=1)

                doc_format_var = set_code + "%" + doc_code + "%format_var"
                setattr(self, doc_format_var, tk.StringVar(self.in_f, value=doc_info["page_size"]))
                doc_format_combobox = set_code + "%" + doc_code + "%format_combobox"
                setattr(self, doc_format_combobox, Combobox(self.in_f, values=list(SIZES_COORDINATES.keys()),
                                                            textvariable=getattr(self, doc_format_var)))
                getattr(self, doc_format_combobox).grid(row=this_row, column=7)

                doc_letters = doc_code.split("-")[-1][:3]

                doc_note_x = 0
                if doc_letters in DOC_SIZES_MAP.keys():
                    doc_note_x = SIZES_COORDINATES[DOC_SIZES_MAP[doc_letters]]["note_x"]
                doc_note_x_id = set_code + "%" + doc_code + "%" + "doc_note_x"
                setattr(self, doc_note_x_id + "%" + "var", tk.StringVar(self.in_f))
                getattr(self, doc_note_x_id + "%" + "var").set(doc_note_x)
                setattr(self, doc_note_x_id + "%" + "entry",
                        tk.Entry(self.in_f,
                                 width=10,
                                 justify="left",
                                 textvariable=getattr(self, doc_note_x_id + "%" + "var")))
                getattr(self, doc_note_x_id + "%" + "entry").grid(sticky="W", row=this_row, column=8, padx=7)

                doc_note_y = 0
                if doc_letters in DOC_SIZES_MAP.keys():
                    doc_note_y = SIZES_COORDINATES[DOC_SIZES_MAP[doc_letters]]["note_y"]
                doc_note_y_id = set_code + "%" + doc_code + "%" + "doc_note_y"
                setattr(self, doc_note_y_id + "%" + "var", tk.StringVar(self.in_f))
                getattr(self, doc_note_y_id + "%" + "var").set(doc_note_y)
                setattr(self, doc_note_y_id + "%" + "entry",
                        tk.Entry(self.in_f,
                                 width=10,
                                 justify="left",
                                 textvariable=getattr(self, doc_note_y_id + "%" + "var")))
                getattr(self, doc_note_y_id + "%" + "entry").grid(sticky="W", row=this_row, column=9, padx=7)

                doc_stamp_x = 0
                if doc_letters in DOC_SIZES_MAP.keys():
                    doc_stamp_x = SIZES_COORDINATES[DOC_SIZES_MAP[doc_letters]]["stamp_x"]
                doc_stamp_x_id = set_code + "%" + doc_code + "%" + "doc_stamp_x"
                setattr(self, doc_stamp_x_id + "%" + "var", tk.StringVar(self.in_f))
                getattr(self, doc_stamp_x_id + "%" + "var").set(doc_stamp_x)
                setattr(self, doc_stamp_x_id + "%" + "entry",
                        tk.Entry(self.in_f,
                                 width=10,
                                 justify="left",
                                 textvariable=getattr(self, doc_stamp_x_id + "%" + "var")))
                getattr(self, doc_stamp_x_id + "%" + "entry").grid(sticky="W", row=this_row, column=10, padx=7)

                doc_stamp_y = 0
                if doc_letters in DOC_SIZES_MAP.keys():
                    doc_stamp_y = SIZES_COORDINATES[DOC_SIZES_MAP[doc_letters]]["stamp_y"]
                doc_stamp_y_id = set_code + "%" + doc_code + "%" + "doc_stamp_y"
                setattr(self, doc_stamp_y_id + "%" + "var", tk.StringVar(self.in_f))
                getattr(self, doc_stamp_y_id + "%" + "var").set(doc_stamp_y)
                setattr(self, doc_stamp_y_id + "%" + "entry",
                        tk.Entry(self.in_f,
                                 width=10,
                                 justify="left",
                                 textvariable=getattr(self, doc_stamp_y_id + "%" + "var")))
                getattr(self, doc_stamp_y_id + "%" + "entry").grid(sticky="W", row=this_row, column=11, padx=7)

                doc_scale = 1.0
                doc_scale_id = set_code + "%" + doc_code + "%" + "doc_scale"
                setattr(self, doc_scale_id + "%" + "var", tk.StringVar(self.in_f))
                getattr(self, doc_scale_id + "%" + "var").set(doc_scale)
                setattr(self, doc_scale_id + "%" + "entry",
                        tk.Entry(self.in_f,
                                 width=10,
                                 justify="left",
                                 textvariable=getattr(self, doc_scale_id + "%" + "var")))
                getattr(self, doc_scale_id + "%" + "entry").grid(sticky="W", row=this_row, column=12, padx=7)

                do_archive_note_var = doc_code + "%do_archive_note_var"
                setattr(self, do_archive_note_var, tk.IntVar(self.in_f, value=int(doc_info["has_archive_number"])))
                do_archive_note_check = doc_code + "%do_archive_note_check"
                setattr(self, do_archive_note_check,
                        Checkbutton(self.in_f, variable=getattr(self, do_archive_note_var)))
                getattr(self, do_archive_note_check).grid(row=this_row, column=13)
                this_row += 1
                for i, change in enumerate(doc_info["changes"]):
                    change_description_ru = change["change_description_ru"]
                    change_description_ru_id = set_code + "%" + doc_code + "%" + str(i) + "%" + "change_description_ru"
                    setattr(self, change_description_ru_id + "%" + "var", tk.StringVar(self.in_f))
                    getattr(self, change_description_ru_id + "%" + "var").set(change_description_ru)
                    setattr(self, change_description_ru_id + "%" + "entry",
                            tk.Entry(self.in_f,
                                     width=20,
                                     justify="left",
                                     textvariable=getattr(self, change_description_ru_id + "%" + "var")))
                    getattr(self, change_description_ru_id + "%" + "entry").grid(sticky="W", row=this_row, column=5,
                                                                                 padx=7)

                    change_description_en = change["change_description_en"]
                    change_description_en_id = set_code + "%" + doc_code + "%" + str(i) + "%" + "change_description_en"
                    setattr(self, change_description_en_id + "%" + "var", tk.StringVar(self.in_f))
                    getattr(self, change_description_en_id + "%" + "var").set(change_description_en)
                    setattr(self, change_description_en_id + "%" + "entry",
                            tk.Entry(self.in_f,
                                     width=20,
                                     justify="left",
                                     textvariable=getattr(self, change_description_en_id + "%" + "var")))
                    getattr(self, change_description_en_id + "%" + "entry").grid(sticky="W", row=this_row, column=6,
                                                                                 padx=7)

                    for sheet in change["pages"]:
                        self.add_change_line(set_code, doc_code, int(sheet), this_row, 2, change, doc_info["geometry"])
                        this_row += 1

        self.ok_button = tk.Button(self.in_f, text="OK", command=self.save_settings)
        self.ok_button.grid(sticky="E", row=this_row, column=0, padx=7, pady=7)
        for argument in dir(self):
            for arg_letters in ["%doc_note_x", "%doc_note_y", "%doc_stamp_x", "%doc_stamp_y", "%doc_scale"]:
                if argument.endswith(arg_letters + "%var"):
                    split_argument = argument.split("%")
                    set_code = split_argument[0]
                    doc_code = split_argument[1]
                    getattr(self, argument).trace("w", self._get_geometry_trace_function(set_code, doc_code, arg_letters))
            if argument.endswith("%format_var"):
                split_argument = argument.split("%")
                set_code = split_argument[0]
                doc_code = split_argument[1]
                getattr(self, argument).trace("w", self._get_format_trace_function(set_code, doc_code))

    def add_change_line(self, set_code, document_code, sheet_number, row, column, change, geometry):
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
        setattr(self, change_type_id + "%" + "label", tk.Label(self.in_f, text=tr_ch_type))
        getattr(self, change_type_id + "%" + "label").grid(row=row, column=this_column)
        this_column += 1

        if ch_type == "patch":
            num_of_sections = str(list(filter(lambda x: x[0] == sheet_number, change["sections_number"]))[0][1])
            num_of_sections_id = set_code + "%" + document_code + "%" + str(sheet_number) + "%" + "num_of_sections"
            setattr(self, num_of_sections_id + "%" + "var", tk.StringVar(self.in_f))
            getattr(self, num_of_sections_id + "%" + "var").set(num_of_sections)
            setattr(self, num_of_sections_id + "%" + "entry",
                    tk.Entry(self.in_f,
                             width=5,
                             justify="left",
                             textvariable=getattr(self, num_of_sections_id + "%" + "var")))
            getattr(self, num_of_sections_id + "%" + "entry").grid(sticky="W", row=row, column=this_column, padx=7)
        else:
            num_of_sections = "-"
            num_of_sections_id = set_code + "%" + document_code + "%" + str(sheet_number) + "%" + "no_sections"
            setattr(self, num_of_sections_id + "%" + "label", tk.Label(self.in_f, text=num_of_sections))
            getattr(self, num_of_sections_id + "%" + "label").grid(row=row, column=this_column)
        this_column += 4

        note_x = str(list(filter(lambda x: x[0] == sheet_number, geometry))[0][1][2])
        note_x_id = set_code + "%" + document_code + "%" + str(sheet_number) + "%" + "note_x"
        setattr(self, note_x_id + "%" + "var", tk.DoubleVar(self.in_f))
        getattr(self, note_x_id + "%" + "var").set(note_x)
        setattr(self, note_x_id + "%" + "entry",
                tk.Entry(self.in_f,
                         width=10,
                         justify="left",
                         textvariable=getattr(self, note_x_id + "%" + "var")))
        getattr(self, note_x_id + "%" + "entry").grid(sticky="W", row=row, column=this_column, padx=7)
        this_column += 1

        note_y = str(list(filter(lambda x: x[0] == sheet_number, geometry))[0][1][3])
        note_y_id = set_code + "%" + document_code + "%" + str(sheet_number) + "%" + "note_y"
        setattr(self, note_y_id + "%" + "var", tk.DoubleVar(self.in_f))
        getattr(self, note_y_id + "%" + "var").set(note_y)
        setattr(self, note_y_id + "%" + "entry",
                tk.Entry(self.in_f,
                         width=10,
                         justify="left",
                         textvariable=getattr(self, note_y_id + "%" + "var")))
        getattr(self, note_y_id + "%" + "entry").grid(sticky="W", row=row, column=this_column, padx=7)
        this_column += 1

        stamp_x = str(list(filter(lambda x: x[0] == sheet_number, geometry))[0][1][0])
        stamp_x_id = set_code + "%" + document_code + "%" + str(sheet_number) + "%" + "stamp_x"
        setattr(self, stamp_x_id + "%" + "var", tk.StringVar(self.in_f))
        getattr(self, stamp_x_id + "%" + "var").set(stamp_x)
        setattr(self, stamp_x_id + "%" + "entry",
                tk.Entry(self.in_f,
                         width=10,
                         justify="left",
                         textvariable=getattr(self, stamp_x_id + "%" + "var")))
        getattr(self, stamp_x_id + "%" + "entry").grid(sticky="W", row=row, column=this_column, padx=7)
        this_column += 1

        stamp_y = str(list(filter(lambda x: x[0] == sheet_number, geometry))[0][1][1])
        stamp_y_id = set_code + "%" + document_code + "%" + str(sheet_number) + "%" + "stamp_y"
        setattr(self, stamp_y_id + "%" + "var", tk.StringVar(self.in_f))
        getattr(self, stamp_y_id + "%" + "var").set(stamp_y)
        setattr(self, stamp_y_id + "%" + "entry",
                tk.Entry(self.in_f,
                         width=10,
                         justify="left",
                         textvariable=getattr(self, stamp_y_id + "%" + "var")))
        getattr(self, stamp_y_id + "%" + "entry").grid(sticky="W", row=row, column=this_column, padx=7)
        this_column += 1

        scale = str(list(filter(lambda x: x[0] == sheet_number, geometry))[0][1][4])
        scale_id = set_code + "%" + document_code + "%" + str(sheet_number) + "%" + "scale"
        setattr(self, scale_id + "%" + "var", tk.StringVar(self.in_f))
        getattr(self, scale_id + "%" + "var").set(scale)
        setattr(self, scale_id + "%" + "entry",
                tk.Entry(self.in_f,
                         width=10,
                         justify="left",
                         textvariable=getattr(self, scale_id + "%" + "var")))
        getattr(self, scale_id + "%" + "entry").grid(sticky="W", row=row, column=this_column, padx=7)
        this_column += 1

    def save_settings(self):
        for attribute in dir(self):
            if attribute.endswith("%num_of_sections%var"):
                setting_name = attribute[:-len("%num_of_sections%var")]
                value = getattr(self, attribute).get()
                set_code, doc_code, sheet_number = setting_name.split("%")
                doc_changes = self.master.changes[set_code][doc_code]["changes"]
                for change in doc_changes:
                    if int(sheet_number) in change["pages"]:
                        for i, pair in enumerate(change["sections_number"]):
                            if pair[0] == int(sheet_number):
                                change["sections_number"][i] = (int(sheet_number), value)

            if attribute.endswith("%change_description_ru%var"):
                setting_name = attribute[:-len("%change_description_ru%var")]
                value = getattr(self, attribute).get()
                set_code, doc_code, i = setting_name.split("%")
                doc_changes = self.master.changes[set_code][doc_code]["changes"]
                doc_changes[int(i)]["change_description_ru"] = value

            if attribute.endswith("%change_description_en%var"):
                setting_name = attribute[:-len("%change_description_en%var")]
                value = getattr(self, attribute).get()
                set_code, doc_code, i = setting_name.split("%")
                doc_changes = self.master.changes[set_code][doc_code]["changes"]
                doc_changes[int(i)]["change_description_en"] = value

            if attribute.endswith("%note_x%var"):
                setting_name = attribute[:-len("%note_x%var")]
                value = float(getattr(self, attribute).get())
                set_code, doc_code, sheet_number = setting_name.split("%")
                doc_geometry = self.master.changes[set_code][doc_code]["geometry"]
                for i, page in enumerate(doc_geometry):
                    if int(sheet_number) == page[0]:
                        g = page[1]
                        doc_geometry[i] = (int(sheet_number), (g[0], g[1], value, g[3], g[4]))

            if attribute.endswith("%note_y%var"):
                setting_name = attribute[:-len("%note_y%var")]
                value = float(getattr(self, attribute).get())
                set_code, doc_code, sheet_number = setting_name.split("%")
                doc_geometry = self.master.changes[set_code][doc_code]["geometry"]
                for i, page in enumerate(doc_geometry):
                    if int(sheet_number) == page[0]:
                        g = page[1]
                        doc_geometry[i] = (int(sheet_number), (g[0], g[1], g[2], value, g[4]))

            if attribute.endswith("%stamp_x%var"):
                setting_name = attribute[:-len("%stamp_x%var")]
                value = float(getattr(self, attribute).get())
                set_code, doc_code, sheet_number = setting_name.split("%")
                doc_geometry = self.master.changes[set_code][doc_code]["geometry"]
                for i, page in enumerate(doc_geometry):
                    if int(sheet_number) == page[0]:
                        g = page[1]
                        doc_geometry[i] = (int(sheet_number), (value, g[1], g[2], g[3], g[4]))

            if attribute.endswith("%stamp_y%var"):
                setting_name = attribute[:-len("%stamp_y%var")]
                value = float(getattr(self, attribute).get())
                set_code, doc_code, sheet_number = setting_name.split("%")
                doc_geometry = self.master.changes[set_code][doc_code]["geometry"]
                for i, page in enumerate(doc_geometry):
                    if int(sheet_number) == page[0]:
                        g = page[1]
                        doc_geometry[i] = (int(sheet_number), (g[0], value, g[2], g[3], g[4]))

            if attribute.endswith("%scale%var"):
                setting_name = attribute[:-len("%scale%var")]
                value = float(getattr(self, attribute).get())
                set_code, doc_code, sheet_number = setting_name.split("%")
                doc_geometry = self.master.changes[set_code][doc_code]["geometry"]
                for i, page in enumerate(doc_geometry):
                    if int(sheet_number) == page[0]:
                        g = page[1]
                        doc_geometry[i] = (int(sheet_number), (g[0], g[1], g[2], g[3], value))

            if attribute.endswith("%format_var"):
                setting_name = attribute[:-len("%format_var")]
                value = getattr(self, attribute).get()
                set_code, doc_code = setting_name.split("%")
                doc_info = self.master.changes[set_code][doc_code]
                doc_info["page_size"] = value
        self.master.save_set_changes()
        self.on_exit()

    def _get_geometry_trace_function(self, set_code, doc_code, arg_letters, *args):
        def tracer(*args):
            main_attribute_code = set_code + "%" + doc_code + arg_letters + "%var"
            main_attribute_value = getattr(self, main_attribute_code).get()
            for attr in dir(self):
                if (attr is not main_attribute_code
                        and attr.startswith(set_code + "%" + doc_code)
                        and attr.endswith("%" + arg_letters[5:] + "%var")):
                    getattr(self, attr).set(main_attribute_value)

        return tracer

    def _get_format_trace_function(self, set_code, doc_code):
        def tracer(*args):
            main_attribute_code = set_code + "%" + doc_code + "%format_var"
            main_attribute_value = getattr(self, main_attribute_code).get()
            for name, value in [
                ("%doc_stamp_x%var", SIZES_COORDINATES[main_attribute_value]["stamp_x"]),
                ("%doc_stamp_y%var", SIZES_COORDINATES[main_attribute_value]["stamp_y"]),
                ("%doc_note_x%var", SIZES_COORDINATES[main_attribute_value]["note_x"]),
                ("%doc_note_y%var", SIZES_COORDINATES[main_attribute_value]["note_y"]),
                ("%doc_scale%var", SIZES_COORDINATES[main_attribute_value]["scale"])
            ]:
                getattr(self, set_code + "%" + doc_code + name).set(value)
        return tracer
