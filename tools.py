from copy import deepcopy


def get_latest_change_number(changes):
    highest_change_number = 0
    for set_info in changes.values():
        for doc_info in set_info.values():
            for change in doc_info["changes"]:
                if change["change_number"] > highest_change_number:
                    highest_change_number = change["change_number"]
    return highest_change_number


def get_latest_changes(changes):
    copied_changes = deepcopy(changes)
    highest_change_number = get_latest_change_number(changes)
    docs_for_deletion = []
    sets_for_deletion = []
    for set_code, set_info in changes.items():
        set_has_actual_changes = False
        for doc_code, doc_info in set_info.items():
            doc_has_actual_changes = False
            doc_changes = doc_info["changes"]
            actual_doc_changes = list(filter(lambda x: x["change_number"] == highest_change_number, doc_changes))
            if len(actual_doc_changes) > 0:
                if not doc_has_actual_changes:
                    doc_has_actual_changes = True
                if not set_has_actual_changes:
                    set_has_actual_changes = True
                copied_changes[set_code][doc_code]["changes"] = actual_doc_changes

            if not doc_has_actual_changes:
                docs_for_deletion.append((set_code, doc_code))
        if not set_has_actual_changes:
            sets_for_deletion.append(set_code)
    for set_code, doc_code in docs_for_deletion:
        if set_code in copied_changes.keys() and doc_code in copied_changes[set_code].keys():
            del copied_changes[set_code][doc_code]
    for set_code in sets_for_deletion:
        if set_code in copied_changes.keys():
            del copied_changes[set_code]
    return copied_changes


def update_extracted_changes_with_saved_changes(saved_changes, extracted_changes):
    extracted_changes = deepcopy(extracted_changes)
    for set_code, set_info in extracted_changes.items():
        for doc_code, doc_info in set_info.items():
            doc_changes = doc_info["changes"]
            doc_saved_info = None
            doc_saved_changes = None
            if set_code in saved_changes.keys() and doc_code in saved_changes[set_code].keys():
                doc_saved_info = saved_changes[set_code][doc_code]
                doc_saved_changes = saved_changes[set_code][doc_code]["changes"]
            if doc_saved_changes:
                for change in doc_changes:
                    change_pages = change["pages"]
                    related_change = None
                    for page in change_pages:
                        related_changes_list = list(filter(lambda x: page in x["pages"], doc_saved_changes))
                        if related_changes_list and len(related_changes_list) == 1:
                            related_change = related_changes_list[0]
                            break
                    if related_change:
                        change["change_description_ru"] = related_change["change_description_ru"]
                        change["change_description_en"] = related_change["change_description_en"]
                        if "sections_number" in change.keys() and "sections_number" in related_change.keys():
                            change["sections_number"] = related_change["sections_number"]
                doc_info["has_archive_number"] = doc_saved_info["has_archive_number"]
                doc_info["page_size"] = doc_saved_info["page_size"]
    return extracted_changes


def zip_pages(set_position, pages_list):
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
    if result == "":
        result = str(set_position) + "." + "1"
    return result.strip(",")


def unzip_page_numbers(pages):
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


def chunk_string(string, length):
    result = []
    this_index = 0
    while this_index < len(string):
        result.append(string[this_index:this_index+length])
        this_index += length
    return result


def pack_change_info(stamp_string, change_info):
    this_dict = change_info[int(stamp_string[0])]
    stamp_type = None
    sections_number = None
    if stamp_string[1] == "n":
        stamp_type = "new"
    elif stamp_string[1] == "c":
        stamp_type = "cancel"
    elif stamp_string[1] == "r":
        stamp_type = "replace"
    elif stamp_string[1] == "p":
        stamp_type = "patch"
        sections_number = int(stamp_string[2])
    return {
        "stamp_type": stamp_type,
        "ii_number": this_dict["change_notice_number"],
        "ii_date": this_dict["change_notice_date"],
        "ii_author": this_dict["author"],
        "number_of_sections": sections_number
    }
