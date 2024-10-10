ICON_PATH = r"materials\icon.ico"

SECRET = "ODMxNjc3ODQ4NzQyMzIwMjM0NzIx"
CHAR_OFFSET = "24"
PROGRAM_VERSION = 0.3

AGREED_LIST = [
    "Алексеев А.В./ Alekseev A.V.",
    "Вяткин С.С./ Vyatkin S.S.",
    "Гриценко О.Л./ Gritsenko O.L.",
    "Миронов Е.А./ Mironov E.A.",
    "Мулюкова Д.Ф./ Mulyukova D.F.",
    "Пермяков С.Е./ Permyakov S.E.",
    "Савицкий В.В./ Savitsky V.V.",
    "Феняк Н.И./ Fenyak N.I.",
    "Черноскулов В.С./ Chernoskulov V.S.",
    "Томлехт Г.В./ Tomlekht G.V.",
    "Буравкин А.В./ Buravkin A.V.",
    "Малышева А.Ю./ Malysheva A.Y.",
    "Нестеров Д.А./ Nesterov D.A.",
    "Латушкин А.В./ Latushkin A.V.",
    "Кушнерёв Д.Ю./ Kushneryov D.Y."
]

CHECKED_LIST = [
    "Алексеев А.В./ Alekseev A.V.",
    "Вяткин С.С./ Vyatkin S.S.",
    "Гриценко О.Л./ Gritsenko O.L.",
    "Миронов Е.А./ Mironov E.A.",
    "Мулюкова Д.Ф./ Mulyukova D.F.",
    "Пермяков С.Е./ Permyakov S.E.",
    "Савицкий В.В./ Savitsky V.V.",
    "Феняк Н.И./ Fenyak N.I.",
    "Черноскулов В.С./ Chernoskulov V.S.",
    "Томлехт Г.В./ Tomlekht G.V.",
    "Буравкин А.В./ Buravkin A.V.",
    "Малышева А.Ю./ Malysheva A.Y.",
    "Нестеров Д.А./ Nesterov D.A.",
    "Латушкин А.В./ Latushkin A.V.",
    "Кушнерёв Д.Ю./ Kushneryov D.Y."
]

EXAMINED_LIST = [
    "Полянская Е.Н./ Polianskaia E.N.",
    "Мязина Л.С./ Myazina L.S.",
    "Минаков С.А./ Minakov S.A.",
]

APPROVED_LIST = [
    "Гончарок А.А./ Goncharok A.A.",
    "Гнелицкий В.Г./ Gnelitskiy V.G."
]

DOC_SIZES_MAP = {
    "MAA": "A4",
    "MAB": "A4",
    "MDB": "A4",
    "MFS": "A2",
    "MPB": "A3",
    "MTB": "A2",
    "MTC": "A1",
    "MLH": "A1",
    "MAZ": "A4",
    "MPA": "A3",
    "MFA": "A1"
}

FORMAT_INFO = {
    "A4": {
        "stamp_x": 151,
        "stamp_y": 18,
        "note_x": 70,
        "note_y": 12,
        "scale": 1,
        "height": 297,
        "width": 210,
        "archive_stamp": False,
        "plotter": "ISO_full_bleed_A4_(210.00_x_297.00_MM)"
    },
    "A4x3": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 297,
        "width": 630,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (630.00 x 297.00мм)"
    },
    "A4x4": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 297,
        "width": 841,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (297.00 x 840.00мм)"
    },
    "A4x5": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 297,
        "width": 1051,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (297.00 x 1050.00мм)"
    },
    "A4x6": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 297,
        "width": 1261,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (297.00 x 1261.00мм)"
    },
    "A4x7": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 297,
        "width": 1471,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (297.00 x 1471.00мм)"
    },
    "A4x8": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 297,
        "width": 1682,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (297.00 x 1682.00мм)"
    },
    "A4x9": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 297,
        "width": 1892,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (297.00 x 1892.00мм)"
    },
    "A3": {
        "stamp_x": 256,
        "stamp_y": 3,
        "note_x": 85,
        "note_y": 12.5,
        "scale": 0.8,
        "height": 297,
        "width": 420,
        "archive_stamp": False,
        "plotter": "ISO_full_bleed_A3_(297.00_x_420.00_MM)"
    },
    "A3x3": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 420,
        "width": 891,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (420.00 x 891.00мм)_1"
    },
    "A3x4": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 420,
        "width": 1189,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (420.00 x 1188.00мм)_1"
    },
    "A3x5": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 420,
        "width": 1486,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (420.00 x 1485.00мм)"
    },
    "A3x6": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 420,
        "width": 1783,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (420.00 x 1783.00мм)"
    },
    "A3x7": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 420,
        "width": 2080,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (420.00 x 2080.00мм)"
    },
    "A2": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 420,
        "width": 594,
        "archive_stamp": True,
        "plotter": "ISO_full_bleed_A2_(420.00_x_594.00_MM)"
    },
    "A2x3": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 594,
        "width": 1261,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (594.00 x 1261.00мм)"
    },
    "A2x4": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 594,
        "width": 1682,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (594.00 x 1682.00мм)"
    },
    "A2x5": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 594,
        "width": 2102,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (594.00 x 2102.00мм)"
    },
    "A1": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 594,
        "width": 841,
        "archive_stamp": True,
        "plotter": "ISO_full_bleed_A1_(594.00_x_841.00_MM)"
    },
    "A1x3": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 841,
        "width": 1783,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (841.00 x 1783.00мм)"
    },
    "A1x4": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 841,
        "width": 2378,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (841.00 x 2378.00мм)"
    },
    "A0": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 841,
        "width": 1189,
        "archive_stamp": True,
        "plotter": "ISO_full_bleed_A0_(841.00_x_1189.00_MM)"
    },
    "A0x2": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 1189,
        "width": 1682,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (1189.00 x 1682.00мм)"
    },
    "A0x3": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1,
        "height": 1189,
        "width": 2523,
        "archive_stamp": True,
        "plotter": "UserDefinedMetric (1189.00 x 2523.00мм)"
    },
}

CHANGE_NAME_MAP = {
    "replace": "Зам.",
    "patch": "Изм.",
    "new": "Нов",
    "cancel": "Анн.",
}

JOURNAL_ADDRESS = r"\\aep-dc\so\Журнал регистрации ИИ\Журнал регистрации ИИ.xlsx"
