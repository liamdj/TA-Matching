import copy
import csv
from typing import List, Any, Tuple, Dict, Set
import math
import write_to_google_sheets as write_gs
import g_sheet_consts as gs_consts


def __is_nan(num) -> bool:
    if type(num) == str:
        return num == 'nan'
    return math.isnan(num)


def compare_matrices_multiple_keys(old_matrix: List[Dict[str, Any]],
                                   new_matrix: List[Dict[str, Any]],
                                   keys: List[str],
                                   keys_to_ignore: Set[str] = None,
                                   keys_to_include: Set[str] = None) -> List[
    Tuple[Tuple[str, ...], Dict[str, Any], Dict[str, Any]]]:
    """
    Compare `old_matrix` and `new_matrix` by comparing lines with the same
    values in `keys`. `keys_to_ignore` may contain fields that are not in either
    matrix (they will simply be ignored). `keys_to_include` may contain fields
    that are not in either matrix (they will simply be ignored).
    """
    keys_to_include = set() if keys_to_include is None else keys_to_include
    keys_to_ignore = set() if keys_to_ignore is None else keys_to_ignore

    new_matrix_by_key = {}
    for entry in new_matrix:
        entry_key = tuple(entry[key] for key in keys)
        new_matrix_by_key[entry_key] = entry

    diffs = []
    for old_entry in old_matrix:
        old_key = tuple(old_entry[key] for key in keys)
        old_diffs, new_diffs = find_differences_between_entries(
            old_key, old_entry, new_matrix_by_key, keys_to_ignore,
            keys_to_include)
        if old_diffs or new_diffs:
            diffs.append((old_key, old_diffs, new_diffs))

    for new_key, new_entry in new_matrix_by_key.items():
        for field in keys_to_ignore:
            new_entry.pop(field, None)
        diffs.append((new_key, {}, new_entry))

    return diffs


def compare_matrices(old_matrix: List[Dict[str, Any]],
                     new_matrix: List[Dict[str, Any]], key: str,
                     keys_to_ignore: Set[str] = None,
                     keys_to_include: Set[str] = None) -> List[
    Tuple[str, Dict[str, Any], Dict[str, Any]]]:
    """
    Compare `old_matrix` and `new_matrix` by comparing lines with the same
    values in `key`. `keys_to_ignore` may contain fields that are not in either
    matrix (they will simply be ignored). `keys_to_include` may contain fields
    that are not in either matrix (they will simply be ignored).
    """
    diffs = compare_matrices_multiple_keys(
        old_matrix, new_matrix, [key], keys_to_ignore, keys_to_include)
    new_diffs = []
    for tup, old_diff, new_diff in diffs:
        new_diffs.append((tup[0], old_diff, new_diff))

    return sorted(new_diffs, key=lambda x: x[0])


def find_differences_between_entries(key: Tuple[str, ...],
                                     old_entry: Dict[str, Any],
                                     new_matrix_by_key: Dict[
                                         Tuple[str, ...], Any],
                                     fields_to_ignore: Set[str],
                                     fields_to_include: Set[str]) -> Tuple[
    Dict[str, Any], Dict[str, Any]]:
    """
    `fields_to_ignore` may contain fields that are not in either matrix
    (they will simply be ignored).
    """
    if key not in new_matrix_by_key:
        for field in fields_to_ignore:
            old_entry.pop(field, None)
        return old_entry, {}

    new_entry = new_matrix_by_key[key]
    old_diffs = {}
    for field in old_entry:
        if field in fields_to_ignore or (
                field in new_entry and field not in fields_to_include and (
                new_entry[field] == old_entry[field] or (
                __is_nan(new_entry[field]) and __is_nan(old_entry[field])))):
            del new_entry[field]
        else:
            old_diffs[field] = old_entry[field]
    del new_matrix_by_key[key]
    return old_diffs, new_entry


def combine_old_and_new(differences: List[
    Tuple[Tuple[str, ...], Dict[str, Any], Dict[str, Any]]], keys: List[str],
                        fields_to_include: Set[str]) -> List[Dict[str, Any]]:
    combined = []
    for key_val, old_entry, new_entry in differences:
        combined_entry = {key: key_val[i] for i, key in enumerate(keys)}
        for field, old_val in old_entry.items():
            new_val = new_entry.get(field, "(_removed_)")
            if field in fields_to_include and new_val == old_val:
                combined_entry[field] = old_val
            else:
                combined_entry[field] = f'{old_val} -> {new_val}'
            new_entry.pop(field, None)
        for field, new_val in new_entry.items():
            combined_entry[field] = f'(_added_) -> {new_val}'
        combined.append(combined_entry)
    return combined


def write_csv_from_combined(combined_diffs: List[Dict[str, Any]],
                            header_order: List[str], out_filename: str) -> bool:
    if len(combined_diffs) == 0:
        return False
    with open(out_filename, "w+") as f:
        w = csv.DictWriter(f, header_order, restval="")
        w.writeheader()
        w.writerows(combined_diffs)
        return True


def remove_unused_fields(combined_diff: List[Dict[str, Any]],
                         fields_order: List[str]) -> List[str]:
    unused_fields = set(copy.deepcopy(fields_order))
    for row in combined_diff:
        for row_field in row:
            unused_fields.discard(row_field)
    for row_field in unused_fields:
        fields_order.remove(row_field)
    return fields_order


def compare_two_worksheets_multiple_keys(sheet: write_gs.Spreadsheet,
                                         old_ws_title: str, new_ws_title: str,
                                         compare_keys: List[str],
                                         output_filename: str,
                                         fields_to_ignore: Set[str] = None,
                                         fields_to_include: Set[
                                             str] = None) -> bool:
    """
    `fields_to_ignore` may contain fields that are not in either matrix
    (they will simply be ignored).
    """
    fields_to_ignore = set() if fields_to_ignore is None else fields_to_ignore
    fields_to_include = set() if fields_to_include is None else fields_to_include
    old_ws = write_gs.get_worksheet_from_sheet(sheet, old_ws_title)
    new_ws = write_gs.get_worksheet_from_sheet(sheet, new_ws_title)
    old_records = old_ws.get_all_records()
    new_records = new_ws.get_all_records()
    old_header_row = \
        write_gs.get_rows_of_cells(old_ws, 1, 1, len(old_records[0]))[0]
    new_header_row = \
        write_gs.get_rows_of_cells(new_ws, 1, 1, len(new_records[0]))[0]
    header_row = combine_header_rows(old_header_row, new_header_row)
    diffs = compare_matrices_multiple_keys(
        old_records, new_records, compare_keys, fields_to_ignore,
        fields_to_include)
    combined_diff = combine_old_and_new(
        diffs, compare_keys, fields_to_include)
    header_row = remove_unused_fields(combined_diff, header_row)

    wrote = write_csv_from_combined(combined_diff, header_row, output_filename)
    if not wrote:
        print(f"No differences between {old_ws_title} and {new_ws_title}")
    return wrote


def compare_two_worksheets(sheet: write_gs.Spreadsheet, old_ws_title: str,
                           new_ws_title: str, compare_key: str,
                           output_filename: str,
                           fields_to_ignore: Set[str] = None,
                           fields_to_include: Set[str] = None) -> bool:
    """
    `fields_to_ignore` may contain fields that are not in either matrix
    (they will simply be ignored).
    """
    return compare_two_worksheets_multiple_keys(
        sheet, old_ws_title, new_ws_title, [compare_key], output_filename,
        fields_to_ignore, fields_to_include)


def combine_header_rows(old_row: List[str], new_row: List[str]):
    new_row, old_row, suffix = find_suffix(new_row, old_row)
    old_set = set(old_row)
    new_set = set(new_row)
    combined_row = []
    i = 0
    j = 0
    while i < len(old_row) and j < len(new_row):
        if old_row[i] == new_row[j]:
            combined_row.append(old_row[i])
            i += 1
            j += 1
        # out of order, so pick new order
        elif old_row[i] in new_set:
            while new_row[j] not in old_set:
                combined_row.append(new_row[j])
                j += 1
            old_set.remove(old_row[i])
            del old_row[i]
        else:  # fill in from old order
            while i < len(old_row) and old_row[i] not in new_set:
                combined_row.append(old_row[i])
                i += 1

    while i < len(old_row) and j >= len(new_row):
        combined_row.append(old_row[i])
        i += 1
    while i >= len(old_row) and j < len(new_row):
        combined_row.append(new_row[j])
        j += 1
    return combined_row + suffix


def find_suffix(new_row: List[str], old_row: List[str]) -> Tuple[
    List[str], List[str], List[str]]:
    old_set = set(old_row)
    new_set = set(new_row)
    suffix = []
    i = len(old_row) - 1
    j = len(new_row) - 1
    while i >= 0:
        if old_row[i] in new_set:
            break
        i -= 1
    matching_index_in_old = i
    while i < len(old_row) - 1:
        suffix.append(old_row[i + 1])
        i += 1
    while j >= 0:
        if new_row[j] in old_set:
            break
        j -= 1
    matching_index_in_new = j
    while j < len(new_row) - 1:
        suffix.append(new_row[j + 1])
        j += 1
    old_row = old_row[:matching_index_in_old + 1]
    new_row = new_row[:matching_index_in_new + 1]
    return new_row, old_row, suffix


def write_comparison_of_two_worksheets_multiple_keys(sheet_title: str,
                                                     sheet_abbreviation: str,
                                                     from_tab_title: str,
                                                     to_tab_title: str,
                                                     compare_keys: List[str],
                                                     csv_path: str,
                                                     fields_to_ignore: Set[
                                                         str] = None,
                                                     fields_to_include: Set[
                                                         str] = None):
    fields_to_include = set() if fields_to_include is None else fields_to_include
    fields_to_ignore = set() if fields_to_ignore is None else fields_to_ignore

    sheet = write_gs.get_sheet(sheet_title)
    diffs_sheet = write_gs.get_sheet(gs_consts.GENERIC_DIFFS_SHEET_TITLE)
    diffs_ws_titles = {ws.title: ws for ws in diffs_sheet.worksheets()}

    new_ws_title = f"{sheet_abbreviation}_{from_tab_title}->{to_tab_title}"
    if new_ws_title in diffs_ws_titles:
        url = write_gs.build_url_to_sheet(
            diffs_sheet.id, diffs_ws_titles[new_ws_title].id)
        print(f"Existed: {url}")
    else:
        wrote = compare_two_worksheets_multiple_keys(
            sheet, from_tab_title, to_tab_title, compare_keys, csv_path,
            fields_to_ignore, fields_to_include)
        if wrote:
            ws = write_gs.write_csv_to_new_tab_from_sheet(
                csv_path, diffs_sheet, new_ws_title)
            print(write_gs.build_url_to_sheet(diffs_sheet.id, ws.id))


def write_comparison_of_two_worksheets(sheet_title: str,
                                       sheet_abbreviation: str,
                                       from_tab_title: str, to_tab_title: str,
                                       compare_key: str, csv_path: str,
                                       fields_to_ignore: Set[str] = None,
                                       fields_to_include: Set[str] = None):
    write_comparison_of_two_worksheets_multiple_keys(
        sheet_title, sheet_abbreviation, from_tab_title, to_tab_title,
        [compare_key], csv_path, fields_to_ignore, fields_to_include)
