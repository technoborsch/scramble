from pyautocad import Autocad

acad = Autocad(True)
acad.Application.Documents.Open(r"C:\Users\n.nikiforov\PycharmProjects\IIhelper\stampA4.dwg")
# acad.Application.Documents.Item(1).Activate()
for text in acad.iter_objects("text"):
    if text.TextString == "II_Author":
        text.TextString = "Никифоров\nNikiforov"
    if text.TextString == "II_Date":
        text.TextString = "02.07.2024"
    if text.TextString == "II_Num":
        text.TextString = "04707"
acad.ActiveDocument.Export("stampA4", "pdf", acad.ActiveDocument.ActiveSelectionSet)
