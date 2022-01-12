import re
from typing import List

import g_sheet_consts as gs_consts
from write_to_google_sheets import get_sheet, get_worksheet_from_sheet, \
    get_rows_of_cells


def remove_worksheets_for_execution(tab_num: str):
    matching_sheet = get_sheet(gs_consts.MATCHING_OUTPUT_SHEET_TITLE)
    remove_worksheet(matching_sheet, tab_num)
    remove_worksheets_substring_title(
        get_sheet(gs_consts.MATCHING_OUTPUT_DIFF_SHEET_TITLE), tab_num)
    remove_worksheet(get_sheet(gs_consts.REMOVE_TA_OUTPUT_SHEET_TITLE), tab_num)
    remove_worksheet(
        get_sheet(gs_consts.ADDITIONAL_TA_OUTPUT_SHEET_TITLE), tab_num)
    remove_worksheets_exact_title(
        get_sheet(gs_consts.ALTERNATES_OUTPUT_SHEET_TITLE),
        [f"{tab_num}{chr(ord('A') + j)}" for j in range(26)])
    planning_input_copy_sheet = get_sheet(
        gs_consts.PLANNING_INPUT_COPY_SHEET_TITLE)
    planning_input_copy_worksheets = planning_input_copy_sheet.worksheets()
    remove_worksheet_from_worksheets(
        planning_input_copy_sheet, planning_input_copy_worksheets,
        f"{tab_num}(S)")
    remove_worksheet_from_worksheets(
        planning_input_copy_sheet, planning_input_copy_worksheets,
        f"{tab_num}(F)")
    remove_worksheet_from_worksheets(
        planning_input_copy_sheet, planning_input_copy_worksheets,
        f"{tab_num}(C)")
    remove_worksheet(
        get_sheet(gs_consts.TA_PREFERENCES_INPUT_COPY_SHEET_TITLE), tab_num)
    remove_worksheet(
        get_sheet(gs_consts.INSTRUCTOR_PREFERENCES_INPUT_COPY_SHEET_TITLE),
        tab_num)
    remove_worksheet(
        get_sheet(gs_consts.PARAMS_INPUT_COPY_SHEET_TITLE), tab_num)

    remove_entry_from_toc(matching_sheet, tab_num)


def remove_worksheets_exact_title(sheet, titles_to_remove: List[str]):
    worksheets = {ws.title: ws for ws in (sheet.worksheets())}
    for title in titles_to_remove:
        if title in worksheets:
            sheet.del_worksheet(worksheets[title])


def remove_worksheets_substring_title(sheet, substring_to_remove: str):
    for ws in sheet.worksheets():
        if substring_to_remove in ws.title:
            sheet.del_worksheet(ws)


def remove_worksheet_from_worksheets(sheet, worksheets,
                                     worksheet_title: str) -> bool:
    for ws in worksheets:
        if worksheet_title == ws.title:
            sheet.del_worksheet(ws)
            return True
    return False


def remove_worksheet(sheet, worksheet_title: str) -> bool:
    for ws in sheet.worksheets():
        if worksheet_title == ws.title:
            sheet.del_worksheet(ws)
            return True
    return False


def remove_entry_from_toc(matching_sheet, tab_num: str):
    toc_ws = get_worksheet_from_sheet(
        matching_sheet, gs_consts.OUTPUT_TOC_TAB_TITLE)
    max_ws = matching_sheet.worksheets()[1].title
    cells = get_rows_of_cells(toc_ws, 2, int(max_ws) + 2, 5)
    for i, row in enumerate(cells):
        if not row:
            continue
        match = re.match(r"#([0-9]{3})", row[4].value)
        if match is None:
            continue
        title = match.group(1)
        if int(title) <= int(tab_num):
            if tab_num in title:
                toc_ws.delete_rows(i + 2)
            return
