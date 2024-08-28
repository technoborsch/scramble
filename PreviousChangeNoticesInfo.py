import tkinter as tk

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
            "Автор ИИ",
            "Проверил",
        ], self.window)
        this_row += 1

    # TODO complete
