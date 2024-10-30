import sys
import os
from docxtpl import DocxTemplate

from ChangesExtractor import ChangesExtractor
from tools import zip_pages


class ChangeTextCreator:

    def __init__(self):
        if hasattr(sys, "_MEIPASS"):
            self.template_path = os.path.join(sys._MEIPASS, r"template.docx")
        else:
            self.template_path = r"materials\template.docx"

    def create(self, change_notice_info, changes, revisions_list, output_path):
        doc = DocxTemplate(self.template_path)
        change_notice_info["changes"] = self._generate_changes_text(changes, revisions_list)
        doc.render(change_notice_info)
        doc.save(output_path)

    @staticmethod
    def create_title(change_notice_info, template_path, output_path):
        doc = DocxTemplate(template_path)
        doc.render(change_notice_info)
        doc.save(output_path)

    def _generate_changes_text(self, changes, revisions_list):
        changes_text = {}
        i = 0
        for set_code, set_changes in changes.items():
            russian_text = ""
            english_text = ""
            for document_code, document_info in set_changes.items():
                patch_pages = []
                patch_descriptions_ru = []
                patch_descriptions_en = []
                replace_pages = []
                replace_descriptions_ru = []
                replace_descriptions_en = []
                new_pages = []
                new_descriptions_ru = []
                new_descriptions_en = []
                cancel_pages = []
                cancel_descriptions_ru = []
                cancel_descriptions_en = []

                for change in document_info["changes"]:
                    if change["change_type"] == "patch":
                        patch_pages = change["sections_number"]
                        patch_descriptions_ru.append(change["change_description_ru"])
                        patch_descriptions_en.append(change["change_description_en"])
                    elif change["change_type"] == "replace":
                        replace_pages += change["pages"]
                        replace_descriptions_ru.append(change["change_description_ru"])
                        replace_descriptions_en.append(change["change_description_en"])
                    elif change["change_type"] == "new":
                        new_pages += change["pages"]
                        new_descriptions_ru.append(change["change_description_ru"])
                        new_descriptions_en.append(change["change_description_en"])
                    elif change["change_type"] == "cancel":
                        cancel_pages += change["pages"]
                        cancel_descriptions_ru.append(change["change_description_ru"])
                        cancel_descriptions_en.append(change["change_description_en"])

                if patch_pages:
                    for page in patch_pages:
                        description_ru = ""
                        description_en = ""
                        if len(patch_descriptions_ru) > 0:
                            description_ru = ". " + ". ".join(patch_descriptions_ru)
                        if len(patch_descriptions_en) > 0:
                            description_en = ". " + ". ".join(patch_descriptions_en)

                        russian_text += "\t- лист " \
                                        + str(document_info['set_position']) + "." + str(page[0]) \
                                        + " – внесены изменения в документ " \
                                          f"«{document_info['doc_ru_name']}»{description_ru} (количество участков - {page[1]});\n"
                        english_text += "\t- sheet " \
                                        + str(document_info['set_position']) + "." + str(page[0]) \
                                        + " – changes have been made to the document " \
                                          f"«{document_info['doc_eng_name']}»{description_en} (number of sections - {page[1]});\n"

                if replace_pages:
                    description_ru = ""
                    description_en = ""
                    if len(replace_descriptions_ru) > 0:
                        description_ru = ". " + ". ".join(replace_descriptions_ru)
                    if len(replace_descriptions_en) > 0:
                        description_en = ". " + ". ".join(replace_descriptions_en)

                    if len(replace_pages) == 1:
                        russian_text += "\t- лист " \
                                        + str(document_info['set_position']) + "." + str(replace_pages[0]) \
                                        + " – внесены изменения в документ " \
                                          f"«{document_info['doc_ru_name']}»{description_ru} (лист заменен);\n"
                        english_text += "\t- sheet " \
                                        + str(document_info['set_position']) + "." + str(replace_pages[0]) \
                                        + " – changes have been made to the document " \
                                          f"«{document_info['doc_eng_name']}»{description_en} (sheet has been replaced);\n"
                    else:
                        zipped_pages = zip_pages(document_info["set_position"], replace_pages)
                        russian_text += f"\t- листы {zipped_pages} – внесены изменения в документ " \
                                        f"«{document_info['doc_ru_name']}»{description_ru} (листы заменены);\n"
                        english_text += f"\t- sheets {zipped_pages} – changes have been made to the document " \
                                        f"«{document_info['doc_eng_name']}»{description_en} (sheets have been replaced);\n"

                if cancel_pages:
                    description_ru = ""
                    description_en = ""
                    if len(cancel_descriptions_ru) > 0:
                        description_ru = ". " + ". ".join(cancel_descriptions_ru)
                    if len(cancel_descriptions_en) > 0:
                        description_en = ". " + ". ".join(cancel_descriptions_en)
                    if len(cancel_pages) == 1:
                        russian_text += "\t- лист " \
                                        + str(document_info['set_position']) + "." + str(cancel_pages[0]) \
                                        + f" – аннулирован{description_ru}\n"
                        english_text += "\t- sheet " \
                                        + str(document_info['set_position']) + "." + str(cancel_pages[0]) \
                                        + f" – cancelled{description_en}\n"
                    else:
                        zipped_pages = zip_pages(document_info["set_position"], cancel_pages)
                        russian_text += f"\t- листы {zipped_pages} – листы аннулированы{description_ru}\n"
                        english_text += f"\t- sheets {zipped_pages} – sheets were cancelled{description_en}\n"

                if new_pages:
                    description_ru = ""
                    description_en = ""
                    if len(new_descriptions_ru) > 0:
                        description_ru = ". " + ". ".join(new_descriptions_ru)
                    if len(new_descriptions_en) > 0:
                        description_en = ". " + ". ".join(new_descriptions_en)

                    if len(new_pages) == 1:
                        russian_text += "\t- лист " \
                                        + str(document_info['set_position']) + "." + str(new_pages[0]) \
                                        + " – добавлен лист в документ " \
                                          f"«{document_info['doc_ru_name']}»{description_ru} (новый лист);\n"
                        english_text += "\t- sheet " \
                                        + str(document_info['set_position']) + "." + str(new_pages[0]) \
                                        + " – a sheet has been added to the document " \
                                          f"«{document_info['doc_eng_name']}»{description_en} (new sheet);\n"
                    else:
                        zipped_pages = zip_pages(document_info["set_position"], new_pages)
                        russian_text += f"\t- листы {zipped_pages} – добавлены листы в документ " \
                                        f"«{document_info['doc_ru_name']}»{description_ru} (листы новые);\n"
                        english_text += f"\t- sheets {zipped_pages} – sheets were added to the document " \
                                        f"«{document_info['doc_eng_name']}»{description_en} (sheets are new);\n"

            changes_text[set_code + revisions_list[i]] = russian_text.strip(";\n") + "." + "\n\n" + english_text.strip(";\n") + "."
            i += 1
        return changes_text


if __name__ == "__main__":
    creator = ChangeTextCreator()
    extractor = ChangesExtractor()
    changes_ = extractor.extract(r"D:\ИИ 04808\ИИ 04808\4. ИИ")
