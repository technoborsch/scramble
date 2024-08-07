import pyautocad
from pyautocad import Autocad
from time import sleep

import config


class AcadPrinter:

    def __init__(self):
        self.acad = None
        self._setup()

    def convert(self, dwg_path, output_path):
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
            letters = dwg_path.split("-")[-1][:3].upper()
            size = config.DOC_SIZES_MAP[letters]  # TODO now it is in doc info
            this_config.CanonicalMediaName = config.PLOT_LIST_MAP[size]  # TODO make better chooses
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


if __name__ == "__main__":
    app = Autocad(True)
    dwg = app.Application.Documents.Open(
        r"\\aep-dc\so\Никифоров\ИИ\ИИ 03878\Старая версия\3 ИИ\AKU.0120.10UJA.JNB.TM.TB0107-MTB0007.dwg"
    )
    sleep(0.5)
    this_config = dwg.ActiveLayout
    print(this_config.ConfigName)
    this_config.ConfigName = "DWG To PDF.pc3"
    print(this_config.GetCanonicalMediaNames())
    print(this_config.PlotType)
    print(dwg.Limits)
    print(dwg.Width)
    print(dwg.Height)
    print(dwg.ModelSpace.Units)
    for object_ in app.iter_objects("Polyline"):
        print(object_.ObjectName)
    for object_ in app.iter_objects("rectangle"):
        print(object_.ObjectName)
    for object_ in app.iter_objects("line"):
        print(object_.ObjectName)

    dwg.Close(False)
