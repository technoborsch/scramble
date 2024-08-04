import tkinter as tk
from tkinter import Frame
from tkscrolledframe import ScrolledFrame


class SettingsWindow:

    def __init__(self, master):
        self.master = master
        self.window = tk.Toplevel(master.window)
        self.window.title("Настройки комплекта")
        self.sf = ScrolledFrame(self.window, width=1350, height=700)
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
        ])
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
                    getattr(self, change_description_ru_id + "%" + "entry").grid(sticky="W", row=this_row, column=5, padx=7)

                    change_description_en = change["change_description_en"]
                    change_description_en_id = set_code + "%" + doc_code + "%" + str(i) + "%" + "change_description_en"
                    setattr(self, change_description_en_id + "%" + "var", tk.StringVar(self.in_f))
                    getattr(self, change_description_en_id + "%" + "var").set(change_description_en)
                    setattr(self, change_description_en_id + "%" + "entry",
                            tk.Entry(self.in_f,
                                     width=20,
                                     justify="left",
                                     textvariable=getattr(self, change_description_en_id + "%" + "var")))
                    getattr(self, change_description_en_id + "%" + "entry").grid(sticky="W", row=this_row, column=6, padx=7)

                    for sheet in change["pages"]:
                        self.add_change_line(set_code, doc_code, int(sheet), this_row, 2, change, doc_info["geometry"])
                        this_row += 1

        self.ok_button = tk.Button(self.in_f, text="OK", command=self.save_settings)
        self.ok_button.grid(sticky="E", row=this_row, column=0, padx=7, pady=7)

    def add_labels_line(self, row, labels):
        counter = 0
        for label in labels:
            attr_name = "label_" + str(row) + "_" + str(counter)
            setattr(self, attr_name, tk.Label(self.in_f, text=label))
            getattr(self, attr_name).grid(row=row, column=counter)
            counter += 1

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
            if attribute.endswith("num_of_sections%var"):
                setting_name = attribute[:-len("%num_of_sections%var")]
                value = getattr(self, attribute).get()
                set_code, doc_code, sheet_number = setting_name.split("%")
                doc_changes = self.master.changes[set_code][doc_code]["changes"]
                for change in doc_changes:
                    if int(sheet_number) in change["pages"]:
                        for i, pair in enumerate(change["sections_number"]):
                            if pair[0] == int(sheet_number):
                                change["sections_number"][i] = (int(sheet_number), value)

            if attribute.endswith("change_description_ru%var"):
                setting_name = attribute[:-len("%change_description_ru%var")]
                value = getattr(self, attribute).get()
                set_code, doc_code, i = setting_name.split("%")
                doc_changes = self.master.changes[set_code][doc_code]["changes"]
                doc_changes[int(i)]["change_description_ru"] = value

            if attribute.endswith("change_description_en%var"):
                setting_name = attribute[:-len("%change_description_en%var")]
                value = getattr(self, attribute).get()
                set_code, doc_code, i = setting_name.split("%")
                doc_changes = self.master.changes[set_code][doc_code]["changes"]
                doc_changes[int(i)]["change_description_en"] = value

            if attribute.endswith("note_x%var"):
                setting_name = attribute[:-len("%note_x%var")]
                value = float(getattr(self, attribute).get())
                set_code, doc_code, sheet_number = setting_name.split("%")
                doc_geometry = self.master.changes[set_code][doc_code]["geometry"]
                for i, page in enumerate(doc_geometry):
                    if int(sheet_number) == page[0]:
                        g = page[1]
                        doc_geometry[i] = (int(sheet_number), (g[0], g[1], value, g[3], g[4]))

            if attribute.endswith("note_y%var"):
                setting_name = attribute[:-len("%note_y%var")]
                value = float(getattr(self, attribute).get())
                set_code, doc_code, sheet_number = setting_name.split("%")
                doc_geometry = self.master.changes[set_code][doc_code]["geometry"]
                for i, page in enumerate(doc_geometry):
                    if int(sheet_number) == page[0]:
                        g = page[1]
                        doc_geometry[i] = (int(sheet_number), (g[0], g[1], g[2], value, g[4]))

            if attribute.endswith("stamp_x%var"):
                setting_name = attribute[:-len("%stamp_x%var")]
                value = float(getattr(self, attribute).get())
                set_code, doc_code, sheet_number = setting_name.split("%")
                doc_geometry = self.master.changes[set_code][doc_code]["geometry"]
                for i, page in enumerate(doc_geometry):
                    if int(sheet_number) == page[0]:
                        g = page[1]
                        doc_geometry[i] = (int(sheet_number), (value, g[1], g[2], g[3], g[4]))

            if attribute.endswith("stamp_y%var"):
                setting_name = attribute[:-len("%stamp_y%var")]
                value = float(getattr(self, attribute).get())
                set_code, doc_code, sheet_number = setting_name.split("%")
                doc_geometry = self.master.changes[set_code][doc_code]["geometry"]
                for i, page in enumerate(doc_geometry):
                    if int(sheet_number) == page[0]:
                        g = page[1]
                        doc_geometry[i] = (int(sheet_number), (g[0], value, g[2], g[3], g[4]))

            if attribute.endswith("scale%var"):
                setting_name = attribute[:-len("%scale%var")]
                value = float(getattr(self, attribute).get())
                set_code, doc_code, sheet_number = setting_name.split("%")
                doc_geometry = self.master.changes[set_code][doc_code]["geometry"]
                for i, page in enumerate(doc_geometry):
                    if int(sheet_number) == page[0]:
                        g = page[1]
                        doc_geometry[i] = (int(sheet_number), (g[0], g[1], g[2], g[3], value))
        self.master.save_set_changes()
        self.on_exit()

    def on_exit(self):
        self.window.destroy()
