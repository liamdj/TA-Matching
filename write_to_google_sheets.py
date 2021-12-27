import csv
import datetime
import gspread
import pytz
import re

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
    max_planning_tab = int(planning_inputs_worksheets[0].split(":")[0])
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
    return worksheet.range(f"A{row}:{chr(ord('A') + cols_len)}{row}")


def get_rows_of_cells(worksheet, start_row, end_row, cols_len):
    one_d = worksheet.range(f"A{start_row}:{chr(ord('A') + cols_len)}{end_row}")
    two_d = []
    for i in range(len(one_d) // cols_len):
        two_d.append(one_d[i * (cols_len + 1):(i + 1) * (cols_len + 1)])
    return two_d


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


def write_execution_to_ToC(executor: str, executed_num: str,
                           matching_weight: float, slots_unfilled: int,
                           alternate_matching_weights=[],
                           input_num_executed=None):
    now = datetime.datetime.now(pytz.timezone('America/New_York'))
    date = now.strftime('%m-%d-%Y')
    time = now.strftime('%H:%M:%S')

    if input_num_executed is None:
        input_num_executed = executed_num

    links_to_output, links_to_alternate_output = build_links_to_output(
        executed_num, matching_weight, slots_unfilled,
        alternate_matching_weights)
    links_to_input = build_links_to_input(input_num_executed, executed_num)

    toc_vals = [date, time, executor, "", *links_to_output, *links_to_input,
                *links_to_alternate_output]
    toc_wsh = get_worksheet(
        gs_consts.MATCHING_OUTPUT_SHEET_TITLE, gs_consts.OUTPUT_TOC_TAB_TITLE)
    toc_wsh.insert_row(toc_vals, 2, value_input_option='USER_ENTERED')


def build_links_to_input(input_executed_num: str, params_executed_num: str):
    links_to_input = ['', '', '', '', '', '']
    planning_sheet = get_sheet(gs_consts.PLANNING_INPUT_COPY_SHEET_TITLE)
    planning_input_copy_sheet_id = planning_sheet.id
    planning_worksheets = planning_sheet.worksheets()
    links_to_input[0], _ = build_hyperlink_to_sheet(
        planning_input_copy_sheet_id, f"#{input_executed_num}",
        get_worksheet_from_worksheets(
            planning_worksheets, f"{input_executed_num}:Students",
            planning_sheet.title).id)
    links_to_input[1], _ = build_hyperlink_to_sheet(
        planning_input_copy_sheet_id, f"#{input_executed_num}",
        get_worksheet_from_worksheets(
            planning_worksheets, f"{input_executed_num}:Faculty",
            planning_sheet.title).id)
    links_to_input[2], _ = build_hyperlink_to_sheet(
        planning_input_copy_sheet_id, f"#{input_executed_num}",
        get_worksheet_from_worksheets(
            planning_worksheets, f"{input_executed_num}:Courses",
            planning_sheet.title).id)
    student_prefs_sheet = get_sheet(
        gs_consts.TA_PREFERENCES_INPUT_COPY_SHEET_TITLE)
    links_to_input[3], _ = build_hyperlink_to_sheet(
        student_prefs_sheet.id, f"#{input_executed_num}",
        get_worksheet_from_sheet(
            student_prefs_sheet, input_executed_num).id)
    instructor_prefs = get_sheet(
        gs_consts.INSTRUCTOR_PREFERENCES_INPUT_COPY_SHEET_TITLE)
    links_to_input[4], _ = build_hyperlink_to_sheet(
        instructor_prefs.id, f"#{input_executed_num}", get_worksheet_from_sheet(
            instructor_prefs, input_executed_num).id)
    params_copy_sheet = get_sheet(gs_consts.PARAMS_INPUT_COPY_SHEET_TITLE)
    links_to_input[5], _ = build_hyperlink_to_sheet(
        params_copy_sheet.id, f"#{params_executed_num}",
        get_worksheet_from_sheet(params_copy_sheet, params_executed_num).id)
    return links_to_input


def build_links_to_output(executed_num: str, matching_weight: float,
                          slots_unfilled: int, alternate_matching_weights):
    def _build_hyperlink(sheet_title: str, text_prefix="", text_suffix="",
                         num_alternate=0):
        sheet = get_sheet(sheet_title)
        ws_title = executed_num
        if num_alternate:
            ws_title += chr(ord('A') + num_alternate - 1)
        worksheet_id = sheet.worksheet(ws_title).id
        link_text = f"{text_prefix}#{executed_num}{text_suffix}"
        link_to_output, _ = build_hyperlink_to_sheet(
            sheet.id, link_text, worksheet_id)
        return link_to_output

    if slots_unfilled != 0:
        matching_suffix = f' ({slots_unfilled} slots unfilled)'
    else:
        matching_suffix = f' ({matching_weight:.2f})'

    links_to_output = [_build_hyperlink(
        gs_consts.MATCHING_OUTPUT_SHEET_TITLE, text_suffix=matching_suffix),
        _build_hyperlink(gs_consts.ADDITIONAL_TA_OUTPUT_SHEET_TITLE),
        _build_hyperlink(gs_consts.REMOVE_TA_OUTPUT_SHEET_TITLE)]
    links_to_alternate_output = []
    for i in range(1, len(alternate_matching_weights) + 1):
        links_to_alternate_output.append(
            _build_hyperlink(
                gs_consts.ALTERNATES_OUTPUT_SHEET_TITLE, f'Alt{i} ',
                text_suffix=f' ({alternate_matching_weights[i - 1]:.2f})',
                num_alternate=i))
    return links_to_output, links_to_alternate_output


def write_matrix_to_sheet(matrix, sheetname, worksheet_name=None, wrap=False):
    gc = gspread.service_account(filename='./credentials.json')
    sheet = gc.create(sheetname)
    if worksheet_name:
        inital_worksheet = sheet.get_worksheet(0)
        worksheet = add_worksheet_from_matrix(matrix, sheet, worksheet_name)
        sheet.del_worksheet(inital_worksheet)
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
    return write_csv_to_new_tab_from_sheet(
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


def remove_worksheets_for_execution(tab_num: str):
    def remove_ws(sheet, worksheets, worksheet_title):
        for ws in worksheets:
            if worksheet_title == ws.title:
                sheet.del_worksheet(ws)
                return True
        return False

    def remove_alternates_ws(alternates_sheet):
        alternates_worksheets = alternates_sheet.worksheets()
        ws_titles = [ws.title for ws in alternates_worksheets]
        for j in range(100):
            alternate_title = f"{tab_num}{chr(ord('A') + j)}"
            if alternate_title in ws_titles:
                if not remove_ws(
                        alternates_sheet, alternates_worksheets,
                        alternate_title):
                    return

    matching_sheet = get_sheet(gs_consts.MATCHING_OUTPUT_SHEET_TITLE)
    remove_ws(matching_sheet, matching_sheet.worksheets(), tab_num)
    remove_TA_sheet = get_sheet(gs_consts.REMOVE_TA_OUTPUT_SHEET_TITLE)
    remove_ws(remove_TA_sheet, remove_TA_sheet.worksheets(), tab_num)
    additional_TA_sheet = get_sheet(gs_consts.ADDITIONAL_TA_OUTPUT_SHEET_TITLE)
    remove_ws(additional_TA_sheet, additional_TA_sheet.worksheets(), tab_num)
    remove_alternates_ws(get_sheet(gs_consts.ALTERNATES_OUTPUT_SHEET_TITLE))
    planning_input_copy_sheet = get_sheet(
        gs_consts.PLANNING_INPUT_COPY_SHEET_TITLE)
    planning_input_copy_worksheets = planning_input_copy_sheet.worksheets()
    remove_ws(
        planning_input_copy_sheet, planning_input_copy_worksheets,
        f"{tab_num}:Students")
    remove_ws(
        planning_input_copy_sheet, planning_input_copy_worksheets,
        f"{tab_num}:Faculty")
    remove_ws(
        planning_input_copy_sheet, planning_input_copy_worksheets,
        f"{tab_num}:Courses")
    students_preferences_input_copy_sheet = get_sheet(
        gs_consts.TA_PREFERENCES_INPUT_COPY_SHEET_TITLE)
    remove_ws(
        students_preferences_input_copy_sheet,
        students_preferences_input_copy_sheet.worksheets(), tab_num)
    instructor_preferences_input_copy_sheet = get_sheet(
        gs_consts.INSTRUCTOR_PREFERENCES_INPUT_COPY_SHEET_TITLE)
    remove_ws(
        instructor_preferences_input_copy_sheet,
        instructor_preferences_input_copy_sheet.worksheets(), tab_num)

    params_copy_input_sheet = get_sheet(gs_consts.PARAMS_INPUT_COPY_SHEET_TITLE)
    remove_ws(
        params_copy_input_sheet, params_copy_input_sheet.worksheets(), tab_num)

    remove_entry_from_toc(matching_sheet, tab_num)


def remove_entry_from_toc(matching_sheet, tab_num):
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


def write_output_csvs(alternates, num_executed, output_dir_title):
    outputs_dir = f'{output_dir_title}/outputs'
    matchings_worksheet = write_csv_to_new_tab(
        f'{outputs_dir}/matchings.csv', gs_consts.MATCHING_OUTPUT_SHEET_TITLE,
        num_executed, 1)
    format(matchings_worksheet, "", "", 4, 18, center_align=True)
    write_csv_to_new_tab(
        f'{outputs_dir}/additional_TA.csv',
        gs_consts.ADDITIONAL_TA_OUTPUT_SHEET_TITLE, num_executed, wrap=True)
    write_csv_to_new_tab(
        f'{outputs_dir}/remove_TA.csv', gs_consts.REMOVE_TA_OUTPUT_SHEET_TITLE,
        num_executed, wrap=True)
    for i in range(alternates):
        write_csv_to_new_tab(
            f'{outputs_dir}/alternate{i + 1}.csv',
            gs_consts.ALTERNATES_OUTPUT_SHEET_TITLE,
            num_executed + chr(ord('A') + i), i)


def write_params_csv(num_executed, output_dir_title):
    sheet = get_sheet(gs_consts.PARAMS_INPUT_COPY_SHEET_TITLE)
    ws = write_csv_to_new_tab_from_sheet(
        f'{output_dir_title}/outputs/params.csv', sheet, num_executed)
    format(ws, "", "", 1, 1, center_align=True)


def copy_input_worksheets(num_executed: str, planning_sheet_id: str,
                          student_prefs_sheet_id: str,
                          instructor_prefs_sheet_id: str):
    planning_sheet = get_sheet_by_id(planning_sheet_id)
    planning_worksheets = planning_sheet.worksheets()
    planning_input_copy_sheet = get_sheet(
        gs_consts.PLANNING_INPUT_COPY_SHEET_TITLE)
    copy_to_from_worksheets(
        planning_sheet.title, planning_worksheets, planning_input_copy_sheet,
        gs_consts.PLANNING_INPUT_COURSES_TAB_TITLE, f"{num_executed}:Courses")
    copy_to_from_worksheets(
        planning_sheet.title, planning_worksheets, planning_input_copy_sheet,
        gs_consts.PLANNING_INPUT_FACULTY_TAB_TITLE, f"{num_executed}:Faculty")
    copy_to_from_worksheets(
        planning_sheet.title, planning_worksheets, planning_input_copy_sheet,
        gs_consts.PLANNING_INPUT_STUDENTS_TAB_TITLE, f"{num_executed}:Students")

    student_prefs_input_copy_sheet = get_sheet(
        gs_consts.TA_PREFERENCES_INPUT_COPY_SHEET_TITLE)
    copy_to(
        get_sheet_by_id(student_prefs_sheet_id), student_prefs_input_copy_sheet,
        gs_consts.PREFERENCES_INPUT_TAB_TITLE, num_executed)

    instructor_prefs_input_copy_sheet = get_sheet(
        gs_consts.INSTRUCTOR_PREFERENCES_INPUT_COPY_SHEET_TITLE)
    copy_to(
        get_sheet_by_id(instructor_prefs_sheet_id),
        instructor_prefs_input_copy_sheet,
        gs_consts.PREFERENCES_INPUT_TAB_TITLE, num_executed)
    return planning_input_copy_sheet.id, student_prefs_input_copy_sheet.id, instructor_prefs_input_copy_sheet.id


def copy_to_from_worksheets(old_worksheet_title: str, old_worksheets, new_sheet,
                            old_tab_title: str, new_tab_title: str):
    worksheet = get_worksheet_from_worksheets(
        old_worksheets, old_tab_title, old_worksheet_title)
    copy_to_from_worksheet(new_sheet, new_tab_title, worksheet)


def copy_to_from_worksheet(new_sheet, new_tab_title, worksheet):
    copied_worksheet_title = worksheet.copy_to(new_sheet.id)['title']
    new_ws = get_worksheet_from_sheet(new_sheet, copied_worksheet_title)
    new_ws.update_title(new_tab_title)
    new_sheet.reorder_worksheets([new_ws])


def copy_to(old_sheet, new_sheet, old_tab_title, new_tab_title):
    worksheet = get_worksheet_from_sheet(old_sheet, old_tab_title)
    copy_to_from_worksheet(new_sheet, new_tab_title, worksheet)
