import re
from typing import List

import g_sheet_consts as gs_consts
from write_to_google_sheets import get_sheet, get_worksheet_from_sheet, \
    get_rows_of_cells, get_rows_of_cells_full


def remove_worksheets_for_execution(tab_num: str):
    remove_worksheets_for_executions([tab_num])


def remove_worksheets_for_executions(tab_nums: List[str]):
    matching_sheet = get_sheet(gs_consts.MATCHING_OUTPUT_SHEET_TITLE)
    remove_worksheets_from_sheets(matching_sheet, tab_nums)
    remove_entries_from_toc(matching_sheet, tab_nums)


def remove_worksheets_from_sheets(matching_sheet, tab_nums: List[str]):
    remove_worksheets_exact_title(matching_sheet, tab_nums)
    remove_worksheets_substring_titles(
        get_sheet(gs_consts.MATCHING_OUTPUT_DIFF_SHEET_TITLE), tab_nums)
    remove_worksheets_exact_title(
        get_sheet(gs_consts.REMOVE_TA_OUTPUT_SHEET_TITLE), tab_nums)
    remove_worksheets_exact_title(
        get_sheet(gs_consts.ADDITIONAL_TA_OUTPUT_SHEET_TITLE), tab_nums)
    alternates_to_delete = []
    for tab_num in tab_nums:
        for j in range(26):
            alternates_to_delete.append(f"{tab_num}{chr(ord('A') + j)}")
    remove_worksheets_exact_title(
        get_sheet(gs_consts.ALTERNATES_OUTPUT_SHEET_TITLE),
        alternates_to_delete)

    planning_input_copy_sheet = get_sheet(
        gs_consts.PLANNING_INPUT_COPY_SHEET_TITLE)
    planning_input_copy_worksheets = planning_input_copy_sheet.worksheets()
    remove_worksheets_from_worksheets(
        planning_input_copy_sheet, planning_input_copy_worksheets,
        [f"{tab_num}(S)" for tab_num in tab_nums])
    remove_worksheets_from_worksheets(
        planning_input_copy_sheet, planning_input_copy_worksheets,
        [f"{tab_num}(F)" for tab_num in tab_nums])
    remove_worksheets_from_worksheets(
        planning_input_copy_sheet, planning_input_copy_worksheets,
        [f"{tab_num}(C)" for tab_num in tab_nums])

    remove_worksheets_exact_title(
        get_sheet(
            gs_consts.TA_PREFERENCES_INPUT_COPY_SHEET_TITLE), tab_nums)
    remove_worksheets_exact_title(
        get_sheet(
            gs_consts.INSTRUCTOR_PREFERENCES_INPUT_COPY_SHEET_TITLE), tab_nums)
    remove_worksheets_exact_title(
        get_sheet(gs_consts.PARAMS_INPUT_COPY_SHEET_TITLE), tab_nums)


def remove_worksheets_exact_title(sheet, titles_to_remove: List[str]):
    worksheets = {ws.title: ws for ws in (sheet.worksheets())}
    for title in titles_to_remove:
        if title in worksheets:
            sheet.del_worksheet(worksheets[title])


def remove_worksheets_substring_titles(sheet, substrings_to_remove: List[str]):
    for substring_to_remove in substrings_to_remove:
        for ws in sheet.worksheets():
            if substring_to_remove in ws.title:
                sheet.del_worksheet(ws)


def remove_worksheets_from_worksheets(sheet, worksheets,
                                      worksheet_titles: List[str]):
    worksheet_titles = set(worksheet_titles)
    for ws in worksheets:
        if ws.title in worksheet_titles:
            sheet.del_worksheet(ws)


def remove_entries_from_toc(matching_sheet, tab_nums: List[str]):
    toc_ws = get_worksheet_from_sheet(
        matching_sheet, gs_consts.OUTPUT_TOC_TAB_TITLE)
    max_ws = matching_sheet.worksheets()[1].title
    cells = get_rows_of_cells_full(
        toc_ws, 2, int(max_ws) + 2, 5, 1, formatted=True)
    for tab_num in tab_nums:
        for row in cells:
            if row:
                if tab_num in row[0]:
                    row[0] = ""
                if tab_num in row[1]:
                    row[1] = ""
    toc_ws.update(
        f'F2:G{int(max_ws) + 2}', cells, value_input_option='USER_ENTERED')

    tab_nums_ints = {int(tab_num) for tab_num in tab_nums}
    tabs_deleted = 0
    cells = get_rows_of_cells(toc_ws, 2, int(max_ws) + 2, 5)
    for i, row in enumerate(cells):
        if not row:
            continue
        match = re.match(r"#([0-9]{3})", row[4])
        if match is None:
            continue
        title = match.group(1)
        if int(title) in tab_nums_ints:
            toc_ws.delete_rows(i + 2 - tabs_deleted)
            print(f'Deleted {title} from ToC')
            tabs_deleted += 1
