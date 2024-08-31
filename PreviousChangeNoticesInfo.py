import tkinter as tk

from transliterate import translit

from AdditionalWindow import AdditionalWindow


class PreviousChangeNoticesInfo(AdditionalWindow):

    def __init__(self, master):
        super().__init__(master)
        self.window.title("Информация по предыдущим ИИ")
        this_row = 0
        self.add_labels_line(this_row, [
            "Номер изменения",
            "Номер ИИ",
            "Дата ИИ",
            "Фамилия автора ИИ\nкириллица",
            "Фамилия автора ИИ\nлатиница",
        ], self.window)
        this_row += 1

        for i in range(self.master.change_number - 1, 0, -1):
            this_column = 0

            label_name = f"{i}_change_label"
            setattr(self, label_name, tk.Label(self.window, text=str(i)))
            getattr(self, label_name).grid(row=this_row, column=this_column)
            this_column += 1

            change_number_entry_name = f"{i}_change_notice_number_entry"
            change_number_var = getattr(self.master, f"{i}_change_notice_number_var")
            setattr(self, change_number_entry_name, tk.Entry(self.window, width=20, justify="right",
                                                             textvariable=change_number_var))
            getattr(self, change_number_entry_name).grid(row=this_row, column=this_column)
            this_column += 1

            change_date_entry_name = f"{i}_change_notice_date_entry"
            change_date_var = getattr(self.master, f"{i}_change_notice_date_var")
            setattr(self, change_date_entry_name, tk.Entry(self.window, width=20, justify="right",
                                                           textvariable=change_date_var))
            getattr(self, change_date_entry_name).grid(row=this_row, column=this_column)
            this_column += 1

            last_name_ru_entry_name = f"{i}_last_name_ru_entry"
            last_name_ru_var = getattr(self.master, f"{i}_last_name_ru_var")
            setattr(self, last_name_ru_entry_name, tk.Entry(self.window, width=20, justify="right",
                                                            textvariable=last_name_ru_var))
            getattr(self, last_name_ru_entry_name).grid(row=this_row, column=this_column)
            this_column += 1

            last_name_en_entry_name = f"{i}_last_name_en_entry"
            last_name_en_var = getattr(self.master, f"{i}_last_name_en_var")
            setattr(self, last_name_en_entry_name, tk.Entry(self.window, width=20, justify="right",
                                                            textvariable=last_name_en_var))
            getattr(self, last_name_en_entry_name).grid(row=this_row, column=this_column)
            this_column += 1

            this_row += 1

        self.ok_button = tk.Button(self.window, text="OK", command=self.save_info)
        self.ok_button.grid(sticky="E", row=this_row, column=0, padx=7, pady=7)

        for i in range(self.master.change_number - 1, 0, -1):
            last_name_ru_var = getattr(self.master, f"{i}_last_name_ru_var")
            last_name_en_var = getattr(self.master, f"{i}_last_name_en_var")
            last_name_ru_var.trace("w", self.get_transliterate_function(last_name_ru_var, last_name_en_var))

    def save_info(self):
        self.on_exit()

    @staticmethod
    def get_transliterate_function(last_name_ru_var, last_name_en_var):
        def tracer(*args):
            last_name_en_var.set(translit(last_name_ru_var.get(), "ru", reversed=True))
        return tracer

    # TODO complete
