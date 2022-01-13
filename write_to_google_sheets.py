import csv
import datetime
import gspread
import pytz
from typing import List, Tuple, Optional

import g_sheet_consts as gs_consts


def get_worksheet_from_sheet(sheet, worksheet_title):
    return get_worksheet_from_worksheets(
        sheet.worksheets(), worksheet_title, sheet.title)


def get_worksheet_from_worksheets(worksheets, worksheet_title, sheet_title):
    for ws in worksheets:
        if ws.title == worksheet_title:
            return ws
    raise ValueError(f"{worksheet_title} not a worksheet in {sheet_title}")


def get_num_execution_from_matchings_sheet(input_num_executed=None) -> str:
    matchings_sheet = get_sheet(gs_consts.MATCHING_OUTPUT_SHEET_TITLE)
    matchings_worksheets = [e.title for e in matchings_sheet.worksheets()]
    matchings_worksheets.remove('ToC')
    matchings_worksheets.sort(key=int, reverse=True)
    max_matchings_tab = int(matchings_worksheets[0])

    if input_num_executed:
        max_planning_tab = int(input_num_executed)
    else:
        max_planning_tab = get_input_num_execution()

    worksheet_title = "001"
    if len(matchings_worksheets) > 0:
        worksheet_title = str(
            f"{(max(max_matchings_tab, max_planning_tab) + 1):03d}")
    return worksheet_title


def get_input_num_execution() -> int:
    planning_inputs_sheet = get_sheet(gs_consts.PLANNING_INPUT_COPY_SHEET_TITLE)
    planning_inputs_worksheets = [e.title for e in
                                  planning_inputs_sheet.worksheets()]
    planning_inputs_worksheets.sort(key=str, reverse=True)
    max_planning_tab = int(planning_inputs_worksheets[0].split("(")[0])
    return max_planning_tab


def add_worksheet(sheet, worksheet_title, rows=100, cols=26, index=0):
    return sheet.add_worksheet(
        title=worksheet_title, rows=str(rows), cols=str(cols), index=index)


def get_worksheet(sheet_title, worksheet_title):
    sheet = get_sheet(sheet_title)
    return get_worksheet_from_sheet(sheet, worksheet_title)


def get_sheet(sheet_title):
    gc = gspread.service_account(filename='./credentials.json')
    return gc.open(sheet_title)


def get_sheet_by_id(sheet_id):
    gc = gspread.service_account(filename='./credentials.json')
    return gc.open_by_key(sheet_id)


def wrap_rows(worksheet, row_start, row_end, cells):
    """ row_start and row_end are inclusive """
    format(
        worksheet, row_start, row_end, 0, cells, center_align=False, bold=False,
        wrap=True)


def format(worksheet, start_row, end_row, start_col, end_col,
           center_align=False, bold=False, wrap=False):
    """start_col and end_col are zero indexed numbers"""
    formatting = {}
    if bold:
        formatting = {"textFormat": {"bold": bold}}
    if center_align:
        formatting["horizontalAlignment"] = "CENTER"
    if wrap:
        formatting["wrapStrategy"] = "WRAP"

    if formatting:
        worksheet.format(
            f"{chr(ord('A') + start_col)}{start_row}:{chr(ord('A') + end_col)}{end_row}",
            formatting)


def format_row(worksheet, row, cells, center_align=False, bold=False,
               wrap=False):
    format(worksheet, row, row, 0, cells, center_align, bold, wrap)


def resize_worksheet_columns(sheet, worksheet, cols):
    body = {"requests": [{"autoResizeDimensions": {
        "dimensions": {"sheetId": worksheet._properties['sheetId'],
                       "dimension": "COLUMNS", "startIndex": 0,
                       "endIndex": cols}, }}]}
    sheet.batch_update(body)


def update_cells(worksheet, cells, values):
    for i in range(len(values)):
        cells[i].value = values[i]
    worksheet.update_cells(cells, value_input_option='USER_ENTERED')


def get_row_of_cells(worksheet, row, cols_len):
    return worksheet.get_values(f"A{row}:{chr(ord('A') + cols_len)}{row}")


def get_rows_of_cells_full(worksheet, start_row, end_row, col_start, cols_len,
                           formatted=False):
    """
    `start_row` is 1 indexed; `col_start` is 0 indexed; `cols_len` is 0 indexed
    """
    if formatted:
        values = worksheet.get_values(
            f"{chr(ord('A') + col_start)}{start_row}:{chr(ord('A') + col_start + cols_len)}{end_row}",
            value_render_option='FORMULA')
        if len(values) == 0:
            return values
        width_diff = cols_len + 1 - len(values[0])
        if width_diff > 0:
            for i, row in enumerate(values):
                values[i] = row + ([""] * width_diff)
        while len(values) < end_row - start_row + 1:
            values.append([""] * (cols_len + 1))
        return values
    return worksheet.get_values(
        f"{chr(ord('A') + col_start)}{start_row}:{chr(ord('A') + col_start + cols_len)}{end_row}")


def get_rows_of_cells(worksheet, start_row, end_row, cols_len, formatted=False):
    return get_rows_of_cells_full(
        worksheet, start_row, end_row, 0, cols_len, formatted)


def append_to_last_row(worksheet, values):
    list_of_lists = worksheet.get_all_values()
    cells = get_row_of_cells(worksheet, len(list_of_lists) + 1, len(values))
    update_cells(worksheet, cells, values)


def build_hyperlink_to_sheet(sheet_id: str, link_text: str,
                             worksheet_id: str = None):
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
    if worksheet_id:
        url = f"{url}#gid={worksheet_id}"
    return f"=HYPERLINK(\"{url}\", \"{link_text}\")", url


def write_execution_to_ToC(toc_ws, executor: str, executed_num: str,
                           matching_weight: float, slots_unfilled: int,
                           alternate_matching_weights=[],
                           input_num_executed: str = None,
                           matching_diff_ws_title: str = None,
                           include_matching_diff=False, input_copy_ids: Tuple[
            Tuple[str, str, str, str], Tuple[str, str], Tuple[str, str]] = None,
                           params_copy_ids: Tuple[str, str] = None,
                           output_ids: Tuple[
                               Tuple[str, str], Optional[Tuple[str, str, str]],
                               Tuple[str, str], Tuple[str, str], Optional[
                                   Tuple[str, List[str]]]] = None):
    """
    If `include_matching_diff == False` then `matching_diff_ws_title` should be
    the message to write in the place of the first of both diff hyperlinks
    """
    now = datetime.datetime.now(pytz.timezone('America/New_York'))
    date = now.strftime('%m-%d-%Y')
    time = now.strftime('%H:%M:%S')

    if input_num_executed is None:
        input_num_executed = executed_num

    if output_ids is None:
        output_ids = initialize_output_ids(
            executed_num, matching_diff_ws_title,
            len(alternate_matching_weights))
    links_to_output, links_to_alternate_output = build_links_to_output(
        executed_num, matching_weight, slots_unfilled,
        alternate_matching_weights, *output_ids, matching_diff_ws_title,
        include_matching_diff)

    if input_copy_ids is None:
        input_copy_ids = initialize_input_copy_ids_tuples(input_num_executed)
    if params_copy_ids is None:
        params_copy_ids = initialize_params_copy_ids(executed_num)
    links_to_input = build_links_to_input(
        input_num_executed, executed_num, *input_copy_ids, params_copy_ids)

    toc_vals = [date, time, executor, "", *links_to_output, *links_to_input,
                *links_to_alternate_output]
    toc_ws.insert_row(toc_vals, 2, value_input_option='USER_ENTERED')


def initialize_input_copy_ids_tuples(input_executed_num: str) -> Tuple[
    Tuple[str, str, str, str], Tuple[str, str], Tuple[str, str]]:
    planning_sheet = get_sheet(gs_consts.PLANNING_INPUT_COPY_SHEET_TITLE)
    planning_worksheets = planning_sheet.worksheets()
    planning_input_copy_students_ws_id = get_worksheet_from_worksheets(
        planning_worksheets, f"{input_executed_num}(S)",
        planning_sheet.title).id
    planning_input_copy_faculty_ws_id = get_worksheet_from_worksheets(
        planning_worksheets, f"{input_executed_num}(F)",
        planning_sheet.title).id
    planning_input_copy_courses_ws_id = get_worksheet_from_worksheets(
        planning_worksheets, f"{input_executed_num}(C)",
        planning_sheet.title).id
    planning_ids = (planning_sheet.id, planning_input_copy_students_ws_id,
                    planning_input_copy_faculty_ws_id,
                    planning_input_copy_courses_ws_id)
    student_prefs_sheet = get_sheet(
        gs_consts.TA_PREFERENCES_INPUT_COPY_SHEET_TITLE)
    student_prefs_ws_id = get_worksheet_from_sheet(
        student_prefs_sheet, input_executed_num).id
    student_prefs_ids = (student_prefs_sheet.id, student_prefs_ws_id)
    instructor_prefs_sheet = get_sheet(
        gs_consts.INSTRUCTOR_PREFERENCES_INPUT_COPY_SHEET_TITLE)
    instructor_prefs_ws_id = get_worksheet_from_sheet(
        instructor_prefs_sheet, input_executed_num).id
    instructor_prefs_ids = (instructor_prefs_sheet.id, instructor_prefs_ws_id)
    return planning_ids, student_prefs_ids, instructor_prefs_ids


def initialize_params_copy_ids(params_executed_num: str) -> Tuple[str, str]:
    params_copy_sheet = get_sheet(gs_consts.PARAMS_INPUT_COPY_SHEET_TITLE)
    params_copy_ws_id = get_worksheet_from_sheet(
        params_copy_sheet, params_executed_num)
    return params_copy_sheet.id, params_copy_ws_id


def initialize_output_ids(num_executed: str,
                          matching_diffs_ws_title: str = None, alternates=0) -> \
        Tuple[Tuple[str, str], Optional[Tuple[str, str, str]], Tuple[str, str],
              Tuple[str, str], Optional[Tuple[str, List[str]]]]:
    matching_output_sheet = get_sheet(gs_consts.MATCHING_OUTPUT_SHEET_TITLE)
    matchings_worksheet = get_worksheet_from_sheet(
        matching_output_sheet, num_executed)
    matchings_ids = (matching_output_sheet.id, matchings_worksheet.id)

    matching_diffs_ids = None
    if matching_diffs_ws_title is not None:
        matching_diff_sheet = get_sheet(
            gs_consts.MATCHING_OUTPUT_DIFF_SHEET_TITLE)
        students_diff_id = get_worksheet_from_sheet(
            matching_diff_sheet, matching_diffs_ws_title + '(S)').id
        courses_diff_id = get_worksheet_from_sheet(
            matching_diff_sheet, matching_diffs_ws_title + '(C)').id
        matching_diffs_ids = (
            matching_diff_sheet.id, students_diff_id, courses_diff_id)

    additional_ta_sheet = get_sheet(gs_consts.ADDITIONAL_TA_OUTPUT_SHEET_TITLE)
    additional_ta_ws = get_worksheet_from_sheet(
        additional_ta_sheet, num_executed)
    additional_ta_ids = (additional_ta_sheet.id, additional_ta_ws.id)

    remove_ta_sheet = get_sheet(gs_consts.REMOVE_TA_OUTPUT_SHEET_TITLE)
    remove_ta_ws = get_worksheet_from_sheet(remove_ta_sheet, num_executed)
    remove_ta_ids = (remove_ta_sheet.id, remove_ta_ws.id)

    alternates_ids = None
    if alternates > 0:
        alternates_sheet = get_sheet(gs_consts.ALTERNATES_OUTPUT_SHEET_TITLE)
        alternates_worksheets_ids = []
        for i in range(alternates):
            alternates_worksheets_ids.append(
                get_worksheet_from_sheet(
                    alternates_sheet, num_executed + chr(ord('A') + i), i).id)
        alternates_ids = (alternates_sheet.id, alternates_worksheets_ids)
    return matchings_ids, matching_diffs_ids, additional_ta_ids, remove_ta_ids, alternates_ids


def build_links_to_input(input_executed_num: str, params_executed_num: str,
                         planning_ids: Tuple[str, str, str, str],
                         student_prefs_ids: Tuple[str, str],
                         instructor_prefs_ids: Tuple[str, str],
                         params_copy_ids: Tuple[str, str]):
    """
    for each of the last 4 inputs, the strings are the id's for sheets or ws's.
    `planning_sheet` is `(sheet_id, students, faculty, courses)`
    """
    link_text = f"#{input_executed_num}"
    links_to_input = [build_hyperlink_to_sheet(
        planning_ids[0], link_text, planning_ids[1]), build_hyperlink_to_sheet(
        planning_ids[0], link_text, planning_ids[2]), build_hyperlink_to_sheet(
        planning_ids[0], link_text, planning_ids[3]), build_hyperlink_to_sheet(
        student_prefs_ids[0], link_text, student_prefs_ids[1]),
        build_hyperlink_to_sheet(
            instructor_prefs_ids[0], link_text, instructor_prefs_ids[1]),
        build_hyperlink_to_sheet(
            params_copy_ids[0], f"#{params_executed_num}", params_copy_ids[1])]
    return [link for link, _ in links_to_input]


def build_links_to_output(executed_num: str, matching_weight: float,
                          slots_unfilled: int,
                          alternate_matching_weights: List[float],
                          matchings_ids: Tuple[str, str],
                          matching_diffs_ids: Optional[Tuple[str, str, str]],
                          additional_ta_ids: Tuple[str, str],
                          remove_ta_ids: Tuple[str, str],
                          alternates_ids: Tuple[str, List[str]] = None,
                          matching_diff_ws_title: str = None,
                          include_matching_diff=False) -> Tuple[
    List[str], List[str]]:
    def _build_hyperlink(sheet_id: str, worksheet_id: str, text_prefix="",
                         text_suffix="") -> str:
        link_text = f"{text_prefix}#{executed_num}{text_suffix}"
        hyperlink, _ = build_hyperlink_to_sheet(
            sheet_id, link_text, worksheet_id)
        return hyperlink

    if slots_unfilled == 0:
        matching_suffix = f' ({matching_weight:.2f})'
    else:
        matching_suffix = f' ({slots_unfilled} slots unfilled)'

    matching_diffs_hyperlinks = [matching_diff_ws_title, ""]
    if include_matching_diff:
        for i, suffix in enumerate(["(S)", "(C)"]):
            matching_diffs_hyperlinks[i], _ = build_hyperlink_to_sheet(
                matching_diffs_ids[0], matching_diff_ws_title + suffix,
                matching_diffs_ids[1 + i])

    links_to_output = [
        _build_hyperlink(*matchings_ids, text_suffix=matching_suffix),
        *matching_diffs_hyperlinks, _build_hyperlink(*additional_ta_ids),
        _build_hyperlink(*remove_ta_ids)]
    links_to_alternate_output = []
    for i in range(1, len(alternate_matching_weights) + 1):
        s = f'{chr(ord("A") + i - 1)} ({alternate_matching_weights[i - 1]:.2f})'
        links_to_alternate_output.append(
            _build_hyperlink(
                alternates_ids[0], alternates_ids[1][i-1], text_prefix=f'Alt{i} ',
                text_suffix=s))
    return links_to_output, links_to_alternate_output


def write_matrix_to_sheet(matrix, sheetname, worksheet_name=None, wrap=False):
    gc = gspread.service_account(filename='./credentials.json')
    sheet = gc.create(sheetname)
    if worksheet_name:
        initial_worksheet = sheet.get_worksheet(0)
        worksheet = add_worksheet_from_matrix(matrix, sheet, worksheet_name)
        sheet.del_worksheet(initial_worksheet)
    else:
        worksheet = sheet.get_worksheet(0)
    print(
        f"Created new spreadsheet at https://docs.google.com/spreadsheets/d/{sheet.id}")
    write_full_worksheet(matrix, worksheet, wrap)
    return sheet.id


def add_worksheet_from_matrix(matrix, sheet, worksheet_name, index=0):
    rows = max(200, len(matrix) + 50)
    cols = max(26, len(matrix[0]) + 3)
    worksheet = add_worksheet(sheet, worksheet_name, rows, cols, index)
    return worksheet


def write_full_worksheet(matrix, worksheet, wrap):
    worksheet.update("A1", matrix)
    worksheet.freeze(rows=1)
    format_row(
        worksheet, 1, len(matrix[0]) + 1, center_align=False, bold=True,
        wrap=wrap)
    if wrap:
        wrap_rows(worksheet, 1, len(matrix), len(matrix[0]))


def write_csv_to_new_tab_from_sheet(csv_path: str, sheet, tab_name: str,
                                    tab_index=0, center_align=False,
                                    wrap=False):
    with open(csv_path, newline='') as csv_file:
        reader = csv.reader(csv_file, delimiter=',', quotechar='"')
        print(f'Writing {csv_path} to {tab_name} in sheet {sheet.title}')
        matrix = list(reader)
        worksheet = write_matrix_to_new_tab_from_sheet(
            matrix, sheet, tab_name, wrap, tab_index)
        resize_worksheet_columns(sheet, worksheet, len(matrix))
        if center_align:
            format(
                worksheet, 1, len(matrix), 0, len(matrix[0]),
                center_align=center_align)
    return worksheet


def write_csv_to_new_tab(csv_path: str, sheet_name: str, tab_name: str,
                         tab_index=0, center_align=False, wrap=False):
    sheet = get_sheet(sheet_name)
    return sheet, write_csv_to_new_tab_from_sheet(
        csv_path, sheet, tab_name, tab_index, center_align, wrap)


def write_matrix_to_new_tab(matrix, sheetname, tab_name, wrap=False,
                            tab_index=0):
    sheet = get_sheet(sheetname)
    write_matrix_to_new_tab_from_sheet(matrix, sheet, tab_name, wrap, tab_index)


def write_matrix_to_new_tab_from_sheet(matrix, sheet, tab_name, wrap,
                                       tab_index):
    worksheet = add_worksheet_from_matrix(matrix, sheet, tab_name, tab_index)
    write_full_worksheet(matrix, worksheet, wrap)
    return worksheet


def write_output_csvs(matching_output_sheet, alternates: int, num_executed: str,
                      output_dir_title: str,
                      matching_diffs_ws_title: str = None) -> Tuple[
    Tuple[str, str], Tuple[str, str, str], Tuple[str, str], Tuple[str, str],
    Tuple[str, List[str]]]:
    outputs_dir = f'{output_dir_title}/outputs'
    matchings_worksheet = write_csv_to_new_tab_from_sheet(
        f'{outputs_dir}/matchings.csv', matching_output_sheet, num_executed, 1)
    format(matchings_worksheet, "", "", 4, 18, center_align=True)
    matchings_ids = (matching_output_sheet.id, matchings_worksheet.id)
    matching_diffs_ids = None
    if matching_diffs_ws_title is not None:
        matching_diff_sheet = get_sheet(
            gs_consts.MATCHING_OUTPUT_DIFF_SHEET_TITLE)
        students_diff_id = write_csv_to_new_tab_from_sheet(
            f'{outputs_dir}/matchings_students_diff.csv', matching_diff_sheet,
            matching_diffs_ws_title + '(S)').id
        courses_diff_id = write_csv_to_new_tab_from_sheet(
            f'{outputs_dir}/matchings_courses_diff.csv', matching_diff_sheet,
            matching_diffs_ws_title + '(C)').id
        matching_diffs_ids = (
            matching_diff_sheet.id, students_diff_id, courses_diff_id)
    additional_ta_sheet, additional_ta_ws = write_csv_to_new_tab(
        f'{outputs_dir}/additional_TA.csv',
        gs_consts.ADDITIONAL_TA_OUTPUT_SHEET_TITLE, num_executed, wrap=True)
    additional_ta_ids = (additional_ta_sheet.id, additional_ta_ws.id)
    remove_ta_sheet, remove_ta_ws = write_csv_to_new_tab(
        f'{outputs_dir}/remove_TA.csv', gs_consts.REMOVE_TA_OUTPUT_SHEET_TITLE,
        num_executed, wrap=True)
    remove_ta_ids = (remove_ta_sheet.id, remove_ta_ws.id)
    alternates_ids = None
    if alternates > 0:
        alternates_sheet = get_sheet(gs_consts.ALTERNATES_OUTPUT_SHEET_TITLE)
        alternates_worksheets_ids = []
        for i in range(alternates):
            alternates_worksheets_ids.append(
                write_csv_to_new_tab_from_sheet(
                    f'{outputs_dir}/alternate{i + 1}.csv', alternates_sheet,
                    num_executed + chr(ord('A') + i), i).id)
        alternates_ids = (alternates_sheet.id, alternates_worksheets_ids)
    return matchings_ids, matching_diffs_ids, additional_ta_ids, remove_ta_ids, alternates_ids


def write_params_csv(num_executed: str, output_dir_title: str) -> Tuple[
    str, str]:
    sheet = get_sheet(gs_consts.PARAMS_INPUT_COPY_SHEET_TITLE)
    ws = write_csv_to_new_tab_from_sheet(
        f'{output_dir_title}/outputs/params.csv', sheet, num_executed)
    format(ws, "", "", 1, 1, center_align=True)
    return sheet.id, ws.id


def copy_input_worksheets(num_executed: str, planning_sheet_id: str,
                          student_prefs_sheet_id: str,
                          instructor_prefs_sheet_id: str) -> Tuple[
    Tuple[str, str, str, str], Tuple[str, str], Tuple[str, str]]:
    planning_sheet = get_sheet_by_id(planning_sheet_id)
    planning_worksheets = planning_sheet.worksheets()
    planning_input_copy_sheet = get_sheet(
        gs_consts.PLANNING_INPUT_COPY_SHEET_TITLE)
    _, planning_courses_copy_ws_id = copy_to_from_worksheets(
        planning_sheet.title, planning_worksheets, planning_input_copy_sheet,
        gs_consts.PLANNING_INPUT_COURSES_TAB_TITLE, f"{num_executed}(C)")
    _, planning_faculty_copy_ws_id = copy_to_from_worksheets(
        planning_sheet.title, planning_worksheets, planning_input_copy_sheet,
        gs_consts.PLANNING_INPUT_FACULTY_TAB_TITLE, f"{num_executed}(F)")
    _, planning_students_copy_ws_id = copy_to_from_worksheets(
        planning_sheet.title, planning_worksheets, planning_input_copy_sheet,
        gs_consts.PLANNING_INPUT_STUDENTS_TAB_TITLE, f"{num_executed}(S)")
    planning_input_copy_ids = (
        planning_input_copy_sheet.id, planning_students_copy_ws_id,
        planning_faculty_copy_ws_id, planning_courses_copy_ws_id)

    student_prefs_copy_ids = copy_to(
        get_sheet_by_id(student_prefs_sheet_id), get_sheet(
            gs_consts.TA_PREFERENCES_INPUT_COPY_SHEET_TITLE),
        gs_consts.PREFERENCES_INPUT_TAB_TITLE, num_executed)

    instructor_prefs_copy_ids = copy_to(
        get_sheet_by_id(instructor_prefs_sheet_id), get_sheet(
            gs_consts.INSTRUCTOR_PREFERENCES_INPUT_COPY_SHEET_TITLE),
        gs_consts.PREFERENCES_INPUT_TAB_TITLE, num_executed)
    return planning_input_copy_ids, student_prefs_copy_ids, instructor_prefs_copy_ids


def copy_to_from_worksheets(old_worksheet_title: str, old_worksheets, new_sheet,
                            old_tab_title: str, new_tab_title: str) -> Tuple[
    str, str]:
    worksheet = get_worksheet_from_worksheets(
        old_worksheets, old_tab_title, old_worksheet_title)
    return copy_to_from_worksheet(new_sheet, new_tab_title, worksheet)


def copy_to_from_worksheet(new_sheet, new_tab_title: str, worksheet) -> Tuple[
    str, str]:
    copied_worksheet_title = worksheet.copy_to(new_sheet.id)['title']
    new_ws = get_worksheet_from_sheet(new_sheet, copied_worksheet_title)
    new_ws.update_title(new_tab_title)
    new_sheet.reorder_worksheets([new_ws])
    return new_sheet.id, new_ws.id


def copy_to(old_sheet, new_sheet, old_tab_title: str, new_tab_title: str) -> \
        Tuple[str, str]:
    worksheet = get_worksheet_from_sheet(old_sheet, old_tab_title)
    return copy_to_from_worksheet(new_sheet, new_tab_title, worksheet)
