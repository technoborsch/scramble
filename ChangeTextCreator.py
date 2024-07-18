from docxtpl import DocxTemplate
from ChangesExtractor import ChangesExtractor


class ChangeTextCreator:

    def __init__(self):
        pass

    def create(self, change_notice_info, changes):
        doc = DocxTemplate(r"materials\template.docx")
        change_notice_info["changes"] = self._generate_changes_text(changes)
        doc.render(change_notice_info)
        doc.save("result.docx")

    def _generate_changes_text(self, changes):
        changes_text = {}
        for set_code, set_changes in changes.items():
            russian_text = ""
            english_text = ""
            for document_code, document_info in set_changes.items():
                patch_pages = []
                replace_pages = []
                new_pages = []
                cancel_pages = []

                for change in document_info["changes"]:
                    if change["change_type"] == "patch":
                        patch_pages = change["sections_number"]
                    elif change["change_type"] == "replace":
                        replace_pages += change["pages"]
                    elif change["change_type"] == "new":
                        new_pages += change["pages"]
                    elif change["change_type"] == "cancel":
                        cancel_pages += change["pages"]

                if patch_pages:
                    for page in patch_pages:
                        russian_text += "\t- лист " \
                                        + str(document_info['set_position']) + "." + str(page[0]) \
                                        + " — внесены изменения в документ " \
                                          f"«{document_info['doc_ru_name']}» (количество участков - {page[1]});\n"
                        english_text += "\t- sheet " \
                                        + str(document_info['set_position']) + "." + str(page[0]) \
                                        + " — changes have been made to the document " \
                                          f"«{document_info['doc_eng_name']}» (number of sections - {page[1]});\n"

                if replace_pages:
                    if len(replace_pages) == 1:
                        russian_text += "\t- лист " \
                                        + str(document_info['set_position']) + "." + str(replace_pages[0]) \
                                        + " — внесены изменения в документ " \
                                          f"«{document_info['doc_ru_name']}» (лист заменен);\n"
                        english_text += "\t- sheet " \
                                        + str(document_info['set_position']) + "." + str(replace_pages[0]) \
                                        + " — changes have been made to the document " \
                                          f"«{document_info['doc_eng_name']}» (sheet has been replaced);\n"
                    else:
                        zipped_pages = self._zip_pages(document_info["set_position"], replace_pages)
                        russian_text += f"\t- листы {zipped_pages} — внесены изменения в документ " \
                                        f"«{document_info['doc_ru_name']}» (листы заменены);\n"
                        english_text += f"\t- sheets {zipped_pages} — changes have been made to the document " \
                                        f"«{document_info['doc_eng_name']}» (sheets have been replaced);\n"

                if cancel_pages:
                    if len(cancel_pages) == 1:
                        russian_text += "\t- лист " \
                                        + str(document_info['set_position']) + "." + str(cancel_pages[0]) \
                                        + " — аннулирован;\n"
                        english_text += "\t- sheet " \
                                        + str(document_info['set_position']) + "." + str(cancel_pages[0]) \
                                        + " — cancelled;\n"
                    else:
                        zipped_pages = self._zip_pages(document_info["set_position"], cancel_pages)
                        russian_text += f"\t- листы {zipped_pages} — листы аннулированы;\n"
                        english_text += f"\t- sheets {zipped_pages} — sheets were cancelled;\n"

                if new_pages:
                    if len(new_pages) == 1:
                        russian_text += "\t- лист " \
                                        + str(document_info['set_position']) + "." + str(new_pages[0]) \
                                        + " — добавлен лист в документ " \
                                          f"«{document_info['doc_ru_name']}» (новый лист);\n"
                        english_text += "\t- sheet " \
                                        + str(document_info['set_position']) + "." + str(new_pages[0]) \
                                        + " — a sheet has been added to the document " \
                                          f"«{document_info['doc_eng_name']}» (new sheet);\n"
                    else:
                        zipped_pages = self._zip_pages(document_info["set_position"], new_pages)
                        russian_text += f"\t- листы {zipped_pages} — добавлены листы в документ " \
                                        f"«{document_info['doc_ru_name']}» (листы новые);\n"
                        english_text += f"\t- sheets {zipped_pages} — sheets were added to the document " \
                                        f"«{document_info['doc_eng_name']}» (sheets are new);\n"

            changes_text[set_code] = russian_text.strip(";\n") + "." + "\n\n" + english_text.strip(";\n") + "."
        return changes_text

    @staticmethod
    def _zip_pages(set_position, pages_list):
        result = ""
        zipped_pages = []
        this_pack = []
        for i in range(1, len(pages_list)):
            if not this_pack:
                if pages_list[i] - pages_list[i - 1] == 1:
                    this_pack.append(pages_list[i - 1])
                else:
                    zipped_pages.append(pages_list[i - 1])
            elif not pages_list[i] - pages_list[i - 1] == 1:
                this_pack.append(pages_list[i - 1])
                zipped_pages.append(this_pack)
                this_pack = []
            if i == len(pages_list) - 1:
                if pages_list[i] - pages_list[i - 1] == 1 and this_pack:
                    this_pack.append(pages_list[i])
                    zipped_pages.append(this_pack)
                    this_pack = []
                else:
                    zipped_pages.append(pages_list[i])
        for item in zipped_pages:
            if isinstance(item, list):
                result += str(set_position) + "." + str(item[0]) + "-" + str(set_position) + "." + str(item[1]) + ","
            else:
                result += str(set_position) + "." + str(item) + ","
        return result.strip(",")


if __name__ == "__main__":
    creator = ChangeTextCreator()
    extractor = ChangesExtractor()
    changes_ = extractor.extract(r"D:\ИИ 04808\ИИ 04808\4. ИИ")
