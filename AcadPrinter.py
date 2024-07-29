import pyautocad
from pyautocad import Autocad
from time import sleep


class AcadPrinter:

    def __init__(self):
        self.acad = Autocad(True)

    def convert(self, dwg_path):
        self.acad.Application.Documents.Open(dwg_path)
        sleep(0.5)

        dwg = self.acad.ActiveDocument
        plot = dwg.Plot
        this_config = dwg.ActiveLayout

        this_config.ConfigName = "DWG To PDF.pc3"
        this_config.UseStandardScale = False
        this_config.StandardScale = pyautocad.ACAD.acScaleToFit
        this_config.PlotRotation = pyautocad.ACAD.ac0degrees
        this_config.PlotType = pyautocad.ACAD.acExtents
        sleep(0.5)

        this_config.RefreshPlotDeviceInfo()
        this_config.CanonicalMediaName = "ISO_full_bleed_A2_(594.00_x_420.00_MM)"  # TODO make choose sizes
        this_config.CenterPlot = True
        this_config.PlotWithPlotStyles = True
        this_config.StyleSheet = "monochrome.ctb"
        sleep(0.5)

        dwg.SetVariable("BACKGROUNDPLOT", 0)

        this_config.RefreshPlotDeviceInfo()

        return plot.PlotToFile(
            dwg_path.replace(".dwg", ".pdf"),
            this_config.ConfigName
        )

    def close_acad(self):
        self.acad.Application.Documents.Close(False)
        self.acad.Application.Quit()
