import copy
import csv
from typing import List, Any, Tuple, Dict, Optional, Set
import math
import write_to_google_sheets as write_gs


def __is_nan(num) -> bool:
    if type(num) == str:
        return num == 'nan'
    return math.isnan(num)


def compare_matrices(old_matrix: List[Dict[str, Any]],
                     new_matrix: List[Dict[str, Any]], key: str,
                     keys_to_ignore: Set[str] = set()) -> List[
    Tuple[str, Dict[str, Any], Dict[str, Any]]]:
    """
    Compare `old_matrix` and `new_matrix` by comparing lines with the same
    values in `key`.
    """
    new_matrix_by_key = {}
    for entry in new_matrix:
        new_matrix_by_key[entry[key]] = entry

    diffs = []
    for old_entry in old_matrix:
        old_key = old_entry[key]
        old_diffs, new_diffs = find_differences_between_entries(
            old_key, old_entry, new_matrix_by_key, keys_to_ignore)
        if old_diffs or new_diffs:
            diffs.append((old_key, old_diffs, new_diffs))

    for new_key, new_entry in new_matrix_by_key.items():
        for field in keys_to_ignore:
            new_entry.pop(field, None)
        diffs.append((new_key, {}, new_entry))

    return diffs


def find_differences_between_entries(key: str, old_entry: Dict[str, Any],
                                     new_matrix_by_key: Dict[str, Any],
                                     fields_to_ignore: Set[str]) -> Tuple[
    Dict[str, Any], Dict[str, Any]]:
    if key not in new_matrix_by_key:
        for field in fields_to_ignore:
            old_entry.pop(field, None)
        return old_entry, {}

    new_entry = new_matrix_by_key[key]
    old_diffs = {}
    for field in old_entry:
        if field in fields_to_ignore or field in new_entry and (
                new_entry[field] == old_entry[field] or (
                __is_nan(new_entry[field]) and __is_nan(old_entry[field]))):
            del new_entry[field]
        else:
            old_diffs[field] = old_entry[field]
    del new_matrix_by_key[key]
    return old_diffs, new_entry


def combine_old_and_new(
        differences: List[Tuple[str, Dict[str, Any], Dict[str, Any]]],
        key: str) -> List[Dict[str, Any]]:
    combined = []
    for key_val, old_entry, new_entry in differences:
        combined_entry = {key: key_val}
        for field, old_val in old_entry.items():
            combined_entry[
                field] = f'{old_val} -> {new_entry.get(field, "(_removed_)")}'
            new_entry.pop(field, None)
        for field, new_val in new_entry.items():
            combined_entry[field] = f'(_added_) -> {new_entry.get(field, "")}'
        combined.append(combined_entry)
    return sorted(combined, key=lambda x: x[key])


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


def compare_two_worksheets(sheet: write_gs.Spreadsheet, old_ws_title: str,
                           new_ws_title: str, compare_key: str,
                           output_filename: str,
                           fields_to_ignore: Set[str] = set()) -> bool:
    old_ws = write_gs.get_worksheet_from_sheet(sheet, old_ws_title)
    new_ws = write_gs.get_worksheet_from_sheet(sheet, new_ws_title)
    old_records = old_ws.get_all_records()
    diffs = compare_matrices(
        old_records, new_ws.get_all_records(), compare_key, fields_to_ignore)
    combined_diff = combine_old_and_new(diffs, compare_key)
    header_row = remove_unused_fields(
        combined_diff,
        write_gs.get_rows_of_cells(old_ws, 1, 1, len(old_records[0]))[0])
    wrote = write_csv_from_combined(combined_diff, header_row, output_filename)
    if not wrote:
        print(f"No differences between {old_ws_title} and {new_ws_title}")
    return wrote
