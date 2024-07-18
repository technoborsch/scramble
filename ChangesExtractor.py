import os
import docx
import pprint

from config import DOC_SIZES_MAP, SIZES_COORDINATES


class ChangesExtractor:

    def __init__(self):
        pass

    def extract(self, folder_path):
        doc_paths = map(lambda x: os.path.join(folder_path, x), self._find_list_of_documents_of_the_set(folder_path))
        changes = {}
        for doc_path in doc_paths:
            doc = docx.Document(doc_path)
            set_code = doc.tables[0].rows[1].cells[0].text.split("-")[0].replace("\n", "")

            for table in doc.tables:
                for row in table.rows[1:-2]:
                    row = row.cells
                    revision = row[2].text.lower()
                    split_revision = revision.split()
                    if "изм" in revision:
                        if set_code not in changes.keys():
                            changes[set_code] = {}
                        doc_code = row[0].text.replace("\n", "")
                        number_of_sheets = split_revision[0].split("/")[1].split(".")[1]
                        if doc_code not in changes[set_code].keys():
                            doc_ru_name, doc_eng_name = row[1].text.split("/")
                            if "MFS" in doc_code.split("-")[1]:
                                doc_ru_name += row[1].text.split()[-1]
                            revision_number = split_revision[0].split("/")[0]
                            set_position = split_revision[0].split("/")[1].split(".")[0]
                            changes[set_code][doc_code] = {
                                "doc_ru_name": doc_ru_name.strip(),
                                "doc_eng_name": doc_eng_name.strip(),
                                "revision": revision_number,
                                "number_of_sheets": number_of_sheets,
                                "set_position": set_position,
                                "changes": [],
                                "geometry": self._fill_geometry(doc_code),
                            }
                        changes[set_code][doc_code]["changes"] = self._extract_changed_sheets(revision,
                                                                                              number_of_sheets)
            for set_code, set_changes in changes.items():
                changes[set_code] = dict(sorted(set_changes.items(), key=lambda item: int(item[1]["set_position"])))
            changes = dict(sorted(changes.items(), key=lambda item: item[0]))
        return changes

    def _extract_changed_sheets(self, revision, number_of_sheets):
        changes = [change.strip(". ") for change in revision.split("изм")]
        changes_list = []
        for change in changes[1:]:
            split_change = change.split()
            change_number = int(split_change[0])
            if len(split_change) == 1:
                changes_list.append({
                    "change_number": change_number,
                    "change_type": "patch",
                    "pages": [num + 1 for num in range(0, int(number_of_sheets))],
                    "sections_number": []
                })
            content = split_change[1:]
            content_string = " ".join(content)
            items = content_string.split(")")
            for item in items:
                if item:
                    pages, type_ = item.split("(")
                    if pages:
                        pages = self._unzip_page_numbers(pages)
                    else:
                        pages = [num + 1 for num in range(0, int(number_of_sheets))]
                    if "зам" in type_:
                        changes_list.append({
                            "change_number": change_number,
                            "change_type": "replace",
                            "pages": pages
                        })
                    elif "нов" in type_:
                        changes_list.append({
                            "change_number": change_number,
                            "change_type": "new",
                            "pages": pages
                        })
                    elif "анн" in type_:
                        changes_list.append({
                            "change_number": change_number,
                            "change_type": "cancel",
                            "pages": pages
                        })
                    elif "уч" in type_ or "изм" in type_:
                        changes_list.append({
                            "change_number": change_number,
                            "change_type": "patch",
                            "pages": pages,
                            "sections_number": []
                        })
        return changes_list

    @staticmethod
    def _find_list_of_documents_of_the_set(folder_path):
        list_of_paths = []
        for filename in os.listdir(folder_path):
            split_filename = filename.split("-")
            if len(split_filename) > 1 and "MAB" in split_filename[1] and filename.endswith("docx"):
                list_of_paths.append(filename)
        return list_of_paths

    @staticmethod
    def _unzip_page_numbers(pages):
        pages_list = []
        for item in pages.split(","):
            split_item = item.split("-")
            if len(split_item) == 1:
                pages_list.append(int(split_item[0].split(".")[1]))
            elif len(split_item) == 2:
                first_number = int(split_item[0].split(".")[1])
                second_number = int(split_item[1].split(".")[1])
                for i in range(first_number, second_number + 1):
                    pages_list.append(i)
        return pages_list

    @staticmethod
    def _fill_geometry(doc_code):
        doc_letters = doc_code.split("-")[1][:-4].upper().strip("\n")
        return SIZES_COORDINATES[DOC_SIZES_MAP[doc_letters]]


if __name__ == "__main__":
    extractor = ChangesExtractor()
    pprinter = pprint.PrettyPrinter(depth=6)
    pprinter.pprint(extractor.extract(r"materials\folder"))
