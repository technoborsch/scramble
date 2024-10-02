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
    "Мараховская О.Н./ Marakhovskaya O.N.",
    "Пермяков С.Е./ Permyakov S.E.",
    "Савицкий В.В./ Savitsky V.V.",
    "Феняк Н.И./ Fenyak N.I.",
    "Черноскулов В.С./ Chernoskulov V.S.",
    "Томлехт Г.В./ Tomlekht G.V.",
    "Буравкин А.В./ Buravkin A.V.",
    "Богословский А.Г./ Bogoslovskiy A.G",
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
    "Мараховская О.Н./ Marakhovskaya O.N.",
    "Пермяков С.Е./ Permyakov S.E.",
    "Савицкий В.В./ Savitsky V.V.",
    "Феняк Н.И./ Fenyak N.I.",
    "Черноскулов В.С./ Chernoskulov V.S.",
    "Томлехт Г.В./ Tomlekht G.V.",
    "Буравкин А.В./ Buravkin A.V.",
    "Богословский А.Г./ Bogoslovskiy A.G",
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
    "MFS": "A2_V",
    "MPB": "A3_V",
    "MTB": "A2_V",
    "MTC": "A1_V",
    "MLH": "A1_V",
    "MAZ": "A4",
    "MPA": "A3_V",
    "MFA": "A1_V"
}

ARCHIVE_MAP = {
    "A4": False,
    "A3_V": False,
    "A2_V": True,
    "A1_V": True,
}

SIZES_COORDINATES = {
    "A4": {
        "stamp_x": 151,
        "stamp_y": 18,
        "note_x": 70,
        "note_y": 12,
        "scale": 1.0
    },
    "A4_V": {
        "stamp_x": 151,
        "stamp_y": 18,
        "note_x": 70,
        "note_y": 12,
        "scale": 1.0
    },
    "A3_V": {
        "stamp_x": 256,
        "stamp_y": 3,
        "note_x": 85,
        "note_y": 12.5,
        "scale": 0.8
    },
    "A2_V": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1.0
    },
    "A1_V": {
        "stamp_x": 341,
        "stamp_y": 5,
        "note_x": 100,
        "note_y": 90,
        "scale": 1.0
    },
}

INITIAL_SIZES = {
    "A4": {
        "height": 297,
        "width": 210,
    },
    "A3_V": {
        "height": 297,
        "width": 420,
    },
    "A2_V": {
        "height": 420,
        "width": 594,
    },
    "A1_V": {
        "height": 594,
        "width": 841,
    },
}

PLOT_LIST_MAP = {
    "A4": "ISO_full_bleed_A4_(210.00_x_297.00_MM)",
    "A3_V": "ISO_full_bleed_A3_(420.00_x_297.00_MM)",
    "A2_V": "ISO_full_bleed_A2_(594.00_x_420.00_MM)",
    "A1_V": "ISO_full_bleed_A1_(841.00_x_594.00_MM)",
}

CHANGE_NAME_MAP = {
    "replace": "Зам.",
    "patch": "Изм.",
    "new": "Нов",
    "cancel": "Анн.",
}

JOURNAL_ADDRESS = r"\\aep-dc\so\Журнал регистрации ИИ\Журнал регистрации ИИ.xlsx"
