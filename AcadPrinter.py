import pyautocad
from pyautocad import Autocad
from time import sleep

import config


class AcadPrinter:

    def __init__(self):
        self.acad = None
        self._setup()

    def convert(self, dwg_path, output_path, doc_size):
        self._setup()
        self.acad.Application.Documents.Open(dwg_path)
        sleep(5)

        dwg = self.acad.ActiveDocument
        plot = dwg.Plot
        this_config = dwg.ActiveLayout
        if not this_config.ConfigName == "DWG To PDF.pc3":
            this_config.ConfigName = "DWG To PDF.pc3"
            this_config.UseStandardScale = False
            this_config.StandardScale = pyautocad.ACAD.acScaleToFit
            this_config.PlotRotation = pyautocad.ACAD.ac0degrees
            this_config.PlotType = pyautocad.ACAD.acExtents

            this_config.RefreshPlotDeviceInfo()
            this_config.CanonicalMediaName = config.PLOT_LIST_MAP[doc_size]
            this_config.CenterPlot = True
            this_config.PlotWithPlotStyles = True
            this_config.StyleSheet = "monochrome.ctb"

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
