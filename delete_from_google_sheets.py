import argparse
import re
from typing import List, Dict

import g_sheet_consts as gs_consts
from write_to_google_sheets import get_sheet, get_rows_of_cells, \
    get_columns_of_cells_formatted, Spreadsheet, Worksheet


def remove_worksheets_for_execution(tab_num: str):
    remove_worksheets_for_executions([tab_num])


def remove_worksheets_for_executions(tab_nums: List[str]):
    matching_sheet = get_sheet(gs_consts.MATCHING_OUTPUT_SHEET_TITLE)
    matching_sheet_worksheets = remove_worksheets_from_sheets(
        matching_sheet, tab_nums)
    max_ws = get_most_recent_worksheet_title(matching_sheet_worksheets)
    toc_ws = matching_sheet_worksheets[gs_consts.OUTPUT_TOC_TAB_TITLE]
    remove_entries_from_toc(toc_ws, max_ws, tab_nums)


def get_most_recent_worksheet_title(
        worksheet_titles: Dict[str, Worksheet]) -> int:
    """ Returns the most recent worksheet by highest integer worksheet title """
    titles = list(worksheet_titles.keys())
    titles.remove(gs_consts.OUTPUT_TOC_TAB_TITLE)
    return int(sorted(titles, key=lambda x: int(x), reverse=True)[0])


def remove_worksheets_from_sheets(matching_sheet: Spreadsheet,
                                  tab_nums: List[str]) -> Dict[str, Worksheet]:
    matching_sheet_worksheets = remove_worksheets_from_sheet(
        matching_sheet, tab_nums)
    remove_worksheets_substring_titles(
        gs_consts.MATCHING_OUTPUT_DIFF_SHEET_TITLE, tab_nums)
    remove_worksheets_substring_titles(
        gs_consts.GENERIC_DIFFS_SHEET_TITLE, tab_nums)
    remove_worksheets_substring_titles(
        gs_consts.AUGMENTING_PATHS_OUTPUT_SHEET_TITLE, tab_nums)
    remove_worksheets(gs_consts.REMOVE_TA_OUTPUT_SHEET_TITLE, tab_nums)
    remove_worksheets(gs_consts.ADDITIONAL_TA_OUTPUT_SHEET_TITLE, tab_nums)
    remove_worksheets(gs_consts.REMOVE_SLOT_OUTPUT_SHEET_TITLE, tab_nums)
    remove_worksheets(gs_consts.ADD_SLOT_OUTPUT_SHEET_TITLE, tab_nums)
    remove_worksheets(gs_consts.COURSE_INTERVIEW_SHEET_TITLE, tab_nums)

    alternates_to_delete = []
    for tab_num in tab_nums:
        for j in range(26):
            alternates_to_delete.append(f"{tab_num}{chr(ord('A') + j)}")

    remove_worksheets(
        gs_consts.ALTERNATES_OUTPUT_SHEET_TITLE, alternates_to_delete)

    planning_input_tabs_to_del = []
    for tab_num in tab_nums:
        for i in ['C', 'S', 'F']:
            planning_input_tabs_to_del.append(f'{tab_num}({i})')
    remove_worksheets(
        gs_consts.PLANNING_INPUT_COPY_SHEET_TITLE, planning_input_tabs_to_del)

    remove_worksheets(gs_consts.TA_PREFERENCES_INPUT_COPY_SHEET_TITLE, tab_nums)
    remove_worksheets(
        gs_consts.INSTRUCTOR_PREFERENCES_INPUT_COPY_SHEET_TITLE, tab_nums)
    remove_worksheets(gs_consts.PARAMS_INPUT_COPY_SHEET_TITLE, tab_nums)
    return matching_sheet_worksheets


def remove_worksheets_from_sheet(sheet: Spreadsheet,
                                 titles_to_remove: List[str]) -> Dict[
    str, Worksheet]:
    worksheets_titles = {ws.title: ws for ws in sheet.worksheets()}
    for title in set(titles_to_remove):
        if title in worksheets_titles:
            sheet.del_worksheet(worksheets_titles[title])
    return worksheets_titles


def remove_worksheets(sheet_title: str, titles_to_remove: List[str]) -> Dict[
    str, Worksheet]:
    return remove_worksheets_from_sheet(
        get_sheet(sheet_title), titles_to_remove)


def remove_worksheets_substring_titles(sheet_title: str,
                                       substrings_to_remove: List[str]):
    sheet = get_sheet(sheet_title)
    for substring_to_remove in substrings_to_remove:
        for ws in sheet.worksheets():
            if substring_to_remove in ws.title:
                sheet.del_worksheet(ws)


def remove_entries_from_toc(toc_ws: Worksheet, max_ws: int,
                            tab_nums: List[str]):
    # Remove from TA Diffs, Course Diffs, Augmenting Paths from Previous
    cells = get_columns_of_cells_formatted(toc_ws, 6, 3)
    for tab_num in tab_nums:
        for row in cells:
            for i, s in enumerate(row):
                if tab_num in s:
                    row[i] = ""
    toc_ws.update(f'F:H', cells, value_input_option='USER_ENTERED')

    tab_nums_ints = {int(tab_num) for tab_num in tab_nums}
    tabs_deleted = 0
    cells = get_rows_of_cells(toc_ws, 2, max_ws + 2, 5)
    for i, row in enumerate(cells):
        if not row:
            continue
        if all(v == '' for v in row):
            toc_ws.delete_rows(i + 2 - tabs_deleted)
            print(f'Deleted empty row {2 + i}')
        match = re.match(r"#([0-9]{3})", row[4])
        if match is None:
            continue
        title = match.group(1)
        if int(title) in tab_nums_ints:
            toc_ws.delete_rows(i + 2 - tabs_deleted)
            print(f'Deleted {title} from ToC')
            tabs_deleted += 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Delete Executions.')
    parser.add_argument(
        'tab_titles', type=str, nargs='+',
        help='the exact titles of the tabs to delete')
    remove_worksheets_for_executions(
        [tab for tab in parser.parse_args().tab_titles])
