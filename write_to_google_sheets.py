import csv
import datetime
from typing import Optional, List, Tuple, Dict, Any

import gspread
import pytz
from gspread import Spreadsheet, Worksheet

import g_sheet_consts as gs_consts

InputCopyIDs = Tuple[
    Tuple[str, str, str, str], Tuple[str, str], Tuple[str, str]]
OutputIDs = Tuple[
    Tuple[str, str], Optional[Tuple[str, str, str]], Optional[Tuple[str, str]],
    Optional[Tuple[str, str]], Optional[Tuple[str, str]], Optional[
        Tuple[str, str]], Optional[Tuple[str, str]], Optional[Tuple[str, str]],
    Optional[Tuple[str, List[str]]]]


def get_worksheet_from_sheet(sheet: Spreadsheet, worksheet_title: str):
    return sheet.worksheet(worksheet_title)


def get_worksheet_from_worksheets(worksheets: List[Worksheet],
                                  worksheet_title: str, sheet_title: str):
    for ws in worksheets:
        if ws.title == worksheet_title:
            return ws
    raise ValueError(f"{worksheet_title} not a worksheet in {sheet_title}")


def get_num_execution_from_matchings_sheet(
        input_num_executed: str = None) -> str:
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


def add_worksheet(sheet: Spreadsheet, worksheet_title: str, rows=100, cols=26,
                  index=0) -> Worksheet:
    return sheet.add_worksheet(
        title=worksheet_title, rows=rows, cols=cols, index=index)


def get_worksheet(sheet_title: str, worksheet_title: str) -> Worksheet:
    sheet = get_sheet(sheet_title)
    return get_worksheet_from_sheet(sheet, worksheet_title)


def get_sheet(sheet_title: str) -> Spreadsheet:
    gc = gspread.service_account(filename='./credentials.json')
    return gc.open(sheet_title)


def get_sheet_by_id(sheet_id: str) -> Spreadsheet:
    gc = gspread.service_account(filename='./credentials.json')
    return gc.open_by_key(sheet_id)


def get_columns_of_cells_formatted(worksheet: Worksheet, col_start: int,
                                   cols: int) -> list:
    """
    `col_start` is 1 indexed; `cols` is 1 indexed. Any empty columns and
    rows will be truncated
    """
    val_range = f"{chr(ord('A') + col_start - 1)}:{chr(ord('A') + col_start + cols - 2)}"
    return worksheet.get_values(val_range, value_render_option='FORMULA')


def get_rows_of_cells(worksheet: Worksheet, start_row: int, end_row: int,
                      cols_len: int) -> list:
    """ `start_row` is 1 indexed; `cols_len` is 0 indexed """
    return worksheet.get_values(
        f"A{start_row}:{chr(ord('A') + cols_len)}{end_row}")


def build_url_to_sheet(sheet_id: str, worksheet_id: str = None) -> str:
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
    if worksheet_id:
        url = f"{url}#gid={worksheet_id}"
    return url


def build_hyperlink_to_sheet(sheet_id: str, link_text: str,
                             worksheet_id: str = None) -> Tuple[str, str]:
    url = build_url_to_sheet(sheet_id, worksheet_id)
    return f"=HYPERLINK(\"{url}\", \"{link_text}\")", url


def write_execution_to_ToC(toc_ws: Worksheet, executor: str, executed_num: str,
                           matching_weight: float, slots_unfilled: int,
                           include_remove_and_add_features=True,
                           include_interviews=True,
                           alternate_matching_weights=[],
                           input_num_executed: str = None,
                           matching_diff_ws_title: str = None,
                           include_matching_diff=False,
                           input_copy_ids: InputCopyIDs = None,
                           params_copy_ids: Tuple[str, str] = None,
                           output_ids: OutputIDs = None):
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
        include_remove_and_add_features,
        include_interviews, alternate_matching_weights, *output_ids,
        matching_diff_ws_title=matching_diff_ws_title,
        include_matching_diff=include_matching_diff)

    if input_copy_ids is None:
        input_copy_ids = initialize_input_copy_ids_tuples(input_num_executed)
    if params_copy_ids is None:
        params_copy_ids = initialize_params_copy_ids(executed_num)
    links_to_input = build_links_to_input(
        input_num_executed, executed_num, *input_copy_ids,
        params_copy_ids=params_copy_ids)

    toc_vals = [date, time, executor, "", *links_to_output, *links_to_input,
                *links_to_alternate_output]
    toc_ws.insert_row(toc_vals, 2, value_input_option='USER_ENTERED')


def initialize_input_copy_ids_tuples(input_executed_num: str) -> InputCopyIDs:
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
    params_copy_ws = get_worksheet_from_sheet(
        params_copy_sheet, params_executed_num)
    return params_copy_sheet.id, params_copy_ws.id


def initialize_output_ids(num_executed: str,
                          matching_diffs_ws_title: str = None,
                          alternates=0) -> OutputIDs:
    matching_output_sheet = get_sheet(gs_consts.MATCHING_OUTPUT_SHEET_TITLE)
    matchings_worksheet = get_worksheet_from_sheet(
        matching_output_sheet, num_executed)
    matchings_ids = (matching_output_sheet.id, matchings_worksheet.id)

    matching_diffs_ids, weighted_changes_ids = None, None
    if matching_diffs_ws_title is not None:
        matching_diff_sheet = get_sheet(
            gs_consts.MATCHING_OUTPUT_DIFF_SHEET_TITLE)
        students_diff_id = get_worksheet_from_sheet(
            matching_diff_sheet, matching_diffs_ws_title + '(S)').id
        courses_diff_id = get_worksheet_from_sheet(
            matching_diff_sheet, matching_diffs_ws_title + '(C)').id
        matching_diffs_ids = (
            matching_diff_sheet.id, students_diff_id, courses_diff_id)

        weighted_changes_sheet = get_sheet(
            gs_consts.WEIGHTED_CHANGES_OUTPUT_SHEET_TITLE)
        weighted_changes_ws = get_worksheet_from_sheet(
            matching_diff_sheet, matching_diffs_ws_title).id
        weighted_changes_ids = (
            weighted_changes_sheet.id, weighted_changes_ws.id)

    additional_ta_sheet = get_sheet(gs_consts.ADDITIONAL_TA_OUTPUT_SHEET_TITLE)
    additional_ta_ws = get_worksheet_from_sheet(
        additional_ta_sheet, num_executed)
    additional_ta_ids = (additional_ta_sheet.id, additional_ta_ws.id)

    remove_ta_sheet = get_sheet(gs_consts.REMOVE_TA_OUTPUT_SHEET_TITLE)
    remove_ta_ws = get_worksheet_from_sheet(remove_ta_sheet, num_executed)
    remove_ta_ids = (remove_ta_sheet.id, remove_ta_ws.id)

    add_slot_sheet = get_sheet(gs_consts.ADD_SLOT_OUTPUT_SHEET_TITLE)
    add_slot_ws = get_worksheet_from_sheet(add_slot_sheet, num_executed)
    add_slot_ids = (add_slot_sheet.id, add_slot_ws.id)

    remove_slot_sheet = get_sheet(gs_consts.REMOVE_SLOT_OUTPUT_SHEET_TITLE)
    remove_slot_ws = get_worksheet_from_sheet(remove_slot_sheet, num_executed)
    remove_slot_ids = (remove_slot_sheet.id, remove_slot_ws.id)

    interviews_sheet = get_sheet(gs_consts.COURSE_INTERVIEW_SHEET_TITLE)
    interviews_ws = get_worksheet_from_sheet(interviews_sheet, num_executed)
    interviews_ids = (interviews_sheet.id, interviews_ws.id)

    alternates_ids = None
    if alternates > 0:
        alternates_sheet = get_sheet(gs_consts.ALTERNATES_OUTPUT_SHEET_TITLE)
        alternates_worksheets_ids = []
        for i in range(alternates):
            alternates_worksheets_ids.append(
                get_worksheet_from_sheet(
                    alternates_sheet, num_executed + chr(ord('A') + i)).id)
        alternates_ids = (alternates_sheet.id, alternates_worksheets_ids)
    return matchings_ids, matching_diffs_ids, weighted_changes_ids, additional_ta_ids, remove_ta_ids, add_slot_ids, remove_slot_ids, interviews_ids, alternates_ids


def build_links_to_input(input_executed_num: str, params_executed_num: str,
                         planning_ids: Tuple[str, str, str, str],
                         student_prefs_ids: Tuple[str, str],
                         instructor_prefs_ids: Tuple[str, str],
                         params_copy_ids: Tuple[str, str]) -> List[str]:
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
                          include_remove_and_add_features: bool,
                          include_interviews: bool,
                          alternate_matching_weights: List[float],
                          matchings_ids: Tuple[str, str],
                          matching_diffs_ids: Optional[Tuple[str, str, str]],
                          weighted_changes_ids: Optional[Tuple[str, str, str]],
                          add_ta_ids: Tuple[str, str],
                          remove_ta_ids: Tuple[str, str],
                          add_slot_ids: Tuple[str, str],
                          remove_slot_ids: Tuple[str, str],
                          interview_ids: Tuple[str, str],
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
    weighted_changes_link = ""
    if include_matching_diff:
        for i, suffix in enumerate(["(S)", "(C)"]):
            matching_diffs_hyperlinks[i], _ = build_hyperlink_to_sheet(
                matching_diffs_ids[0], matching_diff_ws_title + suffix,
                matching_diffs_ids[1 + i])
        weighted_changes_link, _ = build_hyperlink_to_sheet(
            weighted_changes_ids[0], matching_diff_ws_title,
            weighted_changes_ids[1])

    add_ta_links, remove_ta_links, add_slot_links, remove_slot_links = "", "", "", ""
    if include_remove_and_add_features:
        add_ta_links = _build_hyperlink(*add_ta_ids)
        remove_ta_links = _build_hyperlink(*remove_ta_ids)
        add_slot_links = _build_hyperlink(*add_slot_ids)
        remove_slot_links = _build_hyperlink(*remove_slot_ids)

    interview_links = "" if not include_interviews else _build_hyperlink(
        *interview_ids)

    links_to_output = [
        _build_hyperlink(*matchings_ids, text_suffix=matching_suffix),
        *matching_diffs_hyperlinks, weighted_changes_link, add_ta_links,
        remove_ta_links, add_slot_links, remove_slot_links, interview_links]
    links_to_alternate_output = []
    for i in range(len(alternate_matching_weights)):
        suffix = f'{chr(ord("A") + i)} ({alternate_matching_weights[i]:.2f})'
        links_to_alternate_output.append(
            _build_hyperlink(
                alternates_ids[0], alternates_ids[1][i], text_suffix=suffix))
    return links_to_output, links_to_alternate_output


def write_matrix_to_sheet(matrix: List[List[str]], sheet_name: str,
                          worksheet_name: str = None,
                          wrap=False) -> Spreadsheet:
    gc = gspread.service_account(filename='./credentials.json')
    sheet = gc.create(sheet_name)
    initial_worksheet = sheet.get_worksheet(0)
    if not worksheet_name:
        worksheet_name = "Sheet1"
    print(f"Created new spreadsheet at {build_url_to_sheet(sheet.id)}")
    write_matrix_to_new_tab(matrix, sheet, worksheet_name, wrap=wrap)
    sheet.del_worksheet(initial_worksheet)
    return sheet


def create_row_data_from_matrix(matrix: List[List], bold=False,
                                center_align=False, wrap=False,
                                center_align_details: Tuple[int, int] = None) -> \
        List[Dict[str, List[Dict[str, Any]]]]:
    def make_formatting_object(row: int, col: int) -> Dict[str, Any]:
        formatting = {}
        if row == 0 or bold:
            formatting = {"textFormat": {"bold": True}}
        if center_align or (
                center_align_details and center_align_details[0] <= col <=
                center_align_details[1]):
            formatting["horizontalAlignment"] = "CENTER"
        if wrap:
            formatting["wrapStrategy"] = "WRAP"
        return formatting

    row_data = []
    for i, line in enumerate(matrix):
        row = []
        for j, c in enumerate(line):
            key = 'stringValue'
            if type(c) == 'float' or type(c) == 'int':
                key = 'numberValue'
            cell_details = {'userEnteredValue': {key: c}}
            format_object = make_formatting_object(i, j)
            if format_object:
                cell_details['userEnteredFormat'] = format_object
            row.append(cell_details)
        row_data.append({'values': row})
    return row_data


def write_matrix_to_new_tab(matrix: List[List], sheet: Spreadsheet,
                            worksheet_title: str, center_align=False,
                            wrap=False, tab_index=0,
                            center_align_details: Tuple[
                                int, int] = None) -> Worksheet:
    worksheet_id = abs(hash(worksheet_title)) % (10 ** 8)
    body = {"requests": [{"addSheet": {
        "properties": {"sheetId": worksheet_id, "title": worksheet_title,
                       "index": tab_index,
                       "gridProperties": {"rowCount": max(2 * len(matrix), 200),
                                          "columnCount": max(
                                              2 * len(matrix[0]), 26),
                                          "frozenRowCount": 1,
                                          "frozenColumnCount": 0, }, }}},
        {"updateCells": {"rows": create_row_data_from_matrix(
            matrix, False, center_align, wrap, center_align_details),
            "fields": "*", "start": {"sheetId": worksheet_id, "rowIndex": 0,
                                     "columnIndex": 0}}}, {
            "autoResizeDimensions": {"dimensions": {"sheetId": worksheet_id,
                                                    "dimension": "COLUMNS"}}}]}
    sheet.batch_update(body)
    return sheet.worksheet(worksheet_title)


def write_csv_to_new_tab_from_sheet(csv_path: str, sheet: Spreadsheet,
                                    tab_name: str, tab_index=0, wrap=False,
                                    center_align=False,
                                    center_align_details: Tuple[
                                        int, int] = None) -> Worksheet:
    with open(csv_path, newline='') as csv_file:
        reader = csv.reader(csv_file, delimiter=',', quotechar='"')
        print(f'Writing {csv_path} to {tab_name} in sheet {sheet.title}')
        matrix = list(reader)
        worksheet = write_matrix_to_new_tab(
            matrix, sheet, tab_name, center_align, wrap, tab_index,
            center_align_details)
    return worksheet


def write_csv_to_new_tab(csv_path: str, sheet_name: str, tab_name: str,
                         tab_index=0, center_align=False, wrap=False) -> Tuple[
    Spreadsheet, Worksheet]:
    sheet = get_sheet(sheet_name)
    return sheet, write_csv_to_new_tab_from_sheet(
        csv_path, sheet, tab_name, tab_index, wrap, center_align)


def write_output_csvs(matching_output_sheet: Spreadsheet,
                      include_remove_and_add_features: bool,
                      include_interviews: bool, alternates: int,
                      num_executed: str, outputs_dir_path: str,
                      matching_diffs_ws_title: str = None) -> OutputIDs:
    matchings_worksheet = write_csv_to_new_tab_from_sheet(
        f'{outputs_dir_path}/matchings.csv', matching_output_sheet,
        num_executed, 1, center_align_details=(3, 18))
    matchings_ids = (matching_output_sheet.id, matchings_worksheet.id)
    matching_diffs_ids, alternates_ids, weighted_changes_ids = None, None, None
    if matching_diffs_ws_title is not None:
        matching_diff_sheet = get_sheet(
            gs_consts.MATCHING_OUTPUT_DIFF_SHEET_TITLE)
        students_diff_id = write_csv_to_new_tab_from_sheet(
            f'{outputs_dir_path}/matchings_students_diff.csv',
            matching_diff_sheet, matching_diffs_ws_title + '(S)').id
        courses_diff_id = write_csv_to_new_tab_from_sheet(
            f'{outputs_dir_path}/matchings_courses_diff.csv',
            matching_diff_sheet, matching_diffs_ws_title + '(C)').id
        matching_diffs_ids = (
            matching_diff_sheet.id, students_diff_id, courses_diff_id)

        weighted_changes_sheet = get_sheet(
            gs_consts.WEIGHTED_CHANGES_OUTPUT_SHEET_TITLE)
        weighted_changes_ws = write_csv_to_new_tab_from_sheet(
            f'{outputs_dir_path}/weighted_changes.csv', weighted_changes_sheet,
            matching_diffs_ws_title, wrap=True, center_align_details=(0, 1))
        weighted_changes_ids = (
            weighted_changes_sheet.id, weighted_changes_ws.id)

    add_ta_ids, remove_ta_ids = write_add_remove_csvs(
        include_remove_and_add_features, num_executed, outputs_dir_path,
        'additional_TA.csv', 'remove_TA.csv',
        gs_consts.ADDITIONAL_TA_OUTPUT_SHEET_TITLE,
        gs_consts.REMOVE_TA_OUTPUT_SHEET_TITLE)

    add_slot_ids, remove_slot_ids = write_add_remove_csvs(
        include_remove_and_add_features, num_executed, outputs_dir_path,
        'add_slot.csv', 'remove_slot.csv',
        gs_consts.ADD_SLOT_OUTPUT_SHEET_TITLE,
        gs_consts.REMOVE_SLOT_OUTPUT_SHEET_TITLE)

    interview_ids = None
    if include_interviews:
        interview_sheet = get_sheet(gs_consts.COURSE_INTERVIEW_SHEET_TITLE)
        interview_ws = write_csv_to_new_tab_from_sheet(
            f'{outputs_dir_path}/interview_simulations.csv', interview_sheet,
            num_executed, center_align_details=(3, 3))
        interview_ids = (interview_sheet.id, interview_ws.id)

    if alternates > 0:
        alternates_sheet = get_sheet(gs_consts.ALTERNATES_OUTPUT_SHEET_TITLE)
        alternates_worksheets_ids = []
        for i in range(alternates):
            alternates_worksheets_ids.append(
                write_csv_to_new_tab_from_sheet(
                    f'{outputs_dir_path}/alternate{i + 1}.csv',
                    alternates_sheet, num_executed + chr(ord('A') + i), i).id)
        alternates_ids = (alternates_sheet.id, alternates_worksheets_ids)
    return matchings_ids, matching_diffs_ids, weighted_changes_ids, add_ta_ids, remove_ta_ids, add_slot_ids, remove_slot_ids, interview_ids, alternates_ids


def write_add_remove_csvs(include_remove_and_add: bool, num_executed: str,
                          outputs_dir: str, add_csv_suffix: str,
                          remove_csv_suffix: str, add_sheet_title: str,
                          remove_sheet_title: str) -> Tuple[
    Optional[Tuple[str, str]], Optional[Tuple[str, str]]]:
    if not include_remove_and_add:
        return None, None
    add_sheet, add_ws = write_csv_to_new_tab(
        f'{outputs_dir}/{add_csv_suffix}', add_sheet_title, num_executed,
        wrap=True)
    remove_sheet, remove_ws = write_csv_to_new_tab(
        f'{outputs_dir}/{remove_csv_suffix}', remove_sheet_title, num_executed,
        wrap=True)
    return (add_sheet.id, add_ws.id), (remove_sheet.id, remove_ws.id)


def write_params_csv(num_executed: str, output_dir_title: str) -> Tuple[
    str, str]:
    sheet = get_sheet(gs_consts.PARAMS_INPUT_COPY_SHEET_TITLE)
    ws = write_csv_to_new_tab_from_sheet(
        f'{output_dir_title}/params.csv', sheet, num_executed,
        center_align_details=(1, 1))
    return sheet.id, ws.id


def copy_input_worksheets(num_executed: str, planning_sheet_id: str,
                          student_prefs_sheet_id: str,
                          instructor_prefs_sheet_id: str) -> InputCopyIDs:
    print(f"Copying input for execution #{num_executed}")
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


def copy_to_from_worksheets(old_worksheet_title: str,
                            old_worksheets: List[Worksheet],
                            new_sheet: Spreadsheet, old_tab_title: str,
                            new_tab_title: str) -> Tuple[str, str]:
    worksheet = get_worksheet_from_worksheets(
        old_worksheets, old_tab_title, old_worksheet_title)
    return copy_to_from_worksheet(new_sheet, new_tab_title, worksheet)


def copy_to_from_worksheet(new_sheet: Spreadsheet, new_tab_title: str,
                           worksheet: Worksheet) -> Tuple[str, str]:
    new_ws = write_matrix_to_new_tab(
        worksheet.get_all_values(), new_sheet, new_tab_title)
    return new_sheet.id, new_ws.id


def copy_to(old_sheet: Spreadsheet, new_sheet: Spreadsheet, old_tab_title: str,
            new_tab_title: str) -> Tuple[str, str]:
    worksheet = get_worksheet_from_sheet(old_sheet, old_tab_title)
    return copy_to_from_worksheet(new_sheet, new_tab_title, worksheet)
