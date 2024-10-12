from time import sleep
from copy import deepcopy

import pyautocad
from pyautocad import Autocad

import config


class AcadPrinter:
    """
    Класс, отвечающий за работу с AutoCAD через COM API.
    """
    def __init__(self):
        self.acad = None
        self._setup()  # чтобы каждый раз при конвертации создавать новый экземпляр AutoCAD

    def convert(self, dwg_path: str, output_path: str, doc_size: str) -> None:
        """
        Печатает файл dwg в pdf
        :param str dwg_path: Путь до исходного файла dwg
        :param str output_path: Путь сохранения pdf-файла
        :param str doc_size: Размер листа для печати. Должен быть одним из перечисленных в файле config
        :return:
        """
        self._setup()
        self.acad.Application.Documents.Open(dwg_path)
        with applied_plotter_settings(self.acad):
            printed = False
            while not printed:
                try:
                    dwg = self.acad.ActiveDocument
                    plot = dwg.Plot
                    this_config = dwg.ActiveLayout

                    this_config.ConfigName = "DWG To PDF.pc3"
                    this_config.RefreshPlotDeviceInfo()
                    this_config.UseStandardScale = False
                    this_config.StandardScale = pyautocad.ACAD.acScaleToFit
                    this_config.PlotRotation = pyautocad.ACAD.ac90degrees
                    this_config.PlotType = pyautocad.ACAD.acExtents
                    this_config.RefreshPlotDeviceInfo()
                    this_config.CanonicalMediaName = config.FORMAT_INFO[doc_size]["plotter"]
                    this_config.CenterPlot = True
                    this_config.PlotWithPlotStyles = True
                    this_config.StyleSheet = "akku-standard.ctb"

                    dwg.SetVariable("BACKGROUNDPLOT", 0)

                    this_config.RefreshPlotDeviceInfo()

                    plot.PlotToFile(
                        output_path,
                        this_config.ConfigName
                    )
                    printed = True
                except Exception:
                    sleep(0.2)
            return

    def close_acad(self) -> None:
        """
        Закрывает управляемый AutoCAD
        :return:
        """
        if self.acad.Application.Documents.Count != 0:
            for doc in self.acad.Application.Documents:
                doc.Close(False)
                sleep(0.5)

    def _setup(self) -> None:
        """
        Настройка экземпляра AutoCAD
        :return:
        """
        self.acad = Autocad(True, False)


class applied_plotter_settings:
    """
    Контекстный менеджер, который устанавливает пути печати в AutoCAD на файлы, идущие вместе с этой программой
    """
    def __init__(self, acad):
        self.acad = acad

    def __enter__(self):
        """
        Сохраняет оригинальные пути до файлов конфигураций, устанавливает свои
        :return:
        """
        self.original_printer_config_path = deepcopy(self.acad.Application.Preferences.Files.PrinterConfigPath)
        self.original_print_desc_path = deepcopy(self.acad.Application.Preferences.Files.PrinterDescPath)
        self.original_print_styles_path = deepcopy(self.acad.Application.Preferences.Files.PrinterStyleSheetPath)
        self.acad.Application.Preferences.Files.PrinterConfigPath = config.PLOT_CONFIG_PATH
        self.acad.Application.Preferences.Files.PrinterDescPath = config.PLOT_PRINTER_DESC_PATH
        self.acad.Application.Preferences.Files.PrinterStyleSheetPath = config.PLOT_STYLES_PATH
        self.acad.ActiveDocument.ActiveLayout.RefreshPlotDeviceInfo()

    def __exit__(self, type_, value, traceback):
        """
        Возвращает исходные значения настроек AutoCAD
        :param type_:
        :param value:
        :param traceback:
        :return:
        """
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
