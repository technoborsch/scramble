import os
import sys
from time import sleep
from copy import deepcopy

import pyautocad
from pyautocad import Autocad


import config


plot_config_path = os.path.join(os.getcwd(), r"materials\printer_config")
plot_printer_desc_path = os.path.join(os.getcwd(), r"materials\printer_config\PMP Files")
plot_styles_path = os.path.join(os.getcwd(), r"materials\printer_config\Plot Styles")
if hasattr(sys, "_MEIPASS"):
    plot_config_path = os.path.join(sys._MEIPASS, r"printer_config")
    plot_printer_desc_path = os.path.join(sys._MEIPASS, r"printer_config\PMP Files")
    plot_styles_path = os.path.join(sys._MEIPASS, r"printer_config\Plot Styles")


class AcadPrinter:

    def __init__(self):
        self.acad = None
        self._setup()

    def convert(self, dwg_path, output_path, doc_size):
        self._setup()
        self.acad.Application.Documents.Open(dwg_path)
        sleep(5)
        with applied_plotter_settings(self.acad):

            dwg = self.acad.ActiveDocument
            plot = dwg.Plot
            this_config = dwg.ActiveLayout

            this_config.ConfigName = "DWG To PDF.pc3"
            this_config.RefreshPlotDeviceInfo()
            this_config.UseStandardScale = False
            this_config.StandardScale = pyautocad.ACAD.acScaleToFit
            this_config.PlotRotation = pyautocad.ACAD.ac0degrees
            this_config.PlotType = pyautocad.ACAD.acExtents
            this_config.RefreshPlotDeviceInfo()
            this_config.CanonicalMediaName = config.FORMAT_INFO[doc_size]["plotter"]
            this_config.CenterPlot = True
            this_config.PlotWithPlotStyles = True
            this_config.StyleSheet = "akku-standard.ctb"

            dwg.SetVariable("BACKGROUNDPLOT", 0)

            this_config.RefreshPlotDeviceInfo()

            result = plot.PlotToFile(
                output_path,
                this_config.ConfigName
            )
        return result

    def close_acad(self):
        if self.acad.Application.Documents.Count != 0:
            for doc in self.acad.Application.Documents:
                doc.Close(False)
                sleep(0.5)

    def _setup(self):
        self.acad = Autocad(True, False)


class applied_plotter_settings:

    def __init__(self, acad):
        self.acad = acad

    def __enter__(self):
        self.original_printer_config_path = deepcopy(self.acad.Application.Preferences.Files.PrinterConfigPath)
        self.original_print_desc_path = deepcopy(self.acad.Application.Preferences.Files.PrinterDescPath)
        self.original_print_styles_path = deepcopy(self.acad.Application.Preferences.Files.PrinterStyleSheetPath)
        self.acad.Application.Preferences.Files.PrinterConfigPath = plot_config_path
        self.acad.Application.Preferences.Files.PrinterDescPath = plot_printer_desc_path
        self.acad.Application.Preferences.Files.PrinterStyleSheetPath = plot_styles_path
        self.acad.ActiveDocument.ActiveLayout.RefreshPlotDeviceInfo()

    def __exit__(self, type_, value, traceback):
        self.acad.Application.Preferences.Files.PrinterConfigPath = self.original_printer_config_path
        self.acad.Application.Preferences.Files.PrinterDescPath = self.original_print_desc_path
        self.acad.Application.Preferences.Files.PrinterStyleSheetPath = self.original_print_styles_path


if __name__ == "__main__":
    app = Autocad(True, False)
    sleep(3)
    print(app.Application.Preferences.Files.PrinterConfigPath)
    print(app.Application.Preferences.Files.PrinterDescPath)
    print(app.Application.Preferences.Files.PrinterStyleSheetPath)
    names_list = app.Application.ActiveDocument.ActiveLayout.GetCanonicalMediaNames()
    for name in names_list:
        print(name)
