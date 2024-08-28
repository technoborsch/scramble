import tkinter as tk


class AdditionalWindow:

    def __init__(self, master):
        self.master = master
        self.window = tk.Toplevel(master.window)

    def add_labels_line(self, row, labels, window):
        counter = 0
        for label in labels:
            attr_name = "label_" + str(row) + "_" + str(counter)
            setattr(self, attr_name, tk.Label(window, text=label))
            getattr(self, attr_name).grid(row=row, column=counter)
            counter += 1

    def on_exit(self):
        self.window.destroy()
