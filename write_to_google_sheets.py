import csv
import datetime
import gspread
import pytz

import g_sheet_consts as gs_consts


def get_worksheet_from_sheet(sheet, worksheet_title):
    worksheets = [e.title for e in sheet.worksheets()]
    if not worksheet_title in worksheets:
        raise ValueError(f"{worksheet_title} not a worksheet in {sheet.title}")
    return sheet.worksheet(worksheet_title)


def get_num_execution_from_matchings_sheet(sheet) -> str:
    worksheets = [e.title for e in sheet.worksheets()]
    worksheets.remove('ToC')
    worksheets.sort(key=int, reverse=True)
    worksheet_title = "001"
    if len(worksheets) > 0:
        worksheet_title = str(f"{(int(worksheets[0]) + 1):03d}")
    return worksheet_title


def add_worksheet(sheet, worksheet_title, rows=100, cols=26, index=0):
    return sheet.add_worksheet(title=worksheet_title, rows=str(rows),
                               cols=str(cols), index=index)


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
    format(worksheet, row_start, row_end, 0, cells, center_align=False,
           bold=False, wrap=True)


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


def build_hyperlink_to_sheet(sheet_id, link_text, worksheet_id=None):
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
    if worksheet_id:
        url = f"{url}#gid={worksheet_id}"
    return f"=HYPERLINK(\"{url}\", \"{link_text}\")", url


def write_execution_to_ToC(executor: str, executed_num: str,
                           num_alternate_matchings=0,
                           planning_sheet_id: str = None,
                           student_prefs_sheet_id: str = None,
                           instructor_prefs_sheet_id: str = None):
    now = datetime.datetime.now(pytz.timezone('America/New_York'))
    date = now.strftime('%m-%d-%Y')
    time = now.strftime('%H:%M:%S')

    links_to_output, links_to_alternate_output = build_links_to_output(
        executed_num, num_alternate_matchings)
    links_to_input = build_links_to_input(planning_sheet_id,
                                          student_prefs_sheet_id,
                                          instructor_prefs_sheet_id)

    toc_vals = [date, time, executor, "", *links_to_output, *links_to_input,
                *links_to_alternate_output]
    toc_wsh = get_worksheet(gs_consts.MATCHING_OUTPUT_SHEET_TITLE,
                            gs_consts.OUTPUT_TOC_TAB_TITLE)
    toc_wsh.insert_row(toc_vals, 2, value_input_option='USER_ENTERED')


def build_links_to_input(planning_sheet_id: str, student_prefs_sheet_id: str,
                         instructor_prefs_sheet_id: str):
    links_to_input = ['', '', '']
    if planning_sheet_id:
        links_to_input[0], _ = build_hyperlink_to_sheet(planning_sheet_id,
                                                        "TA Planning")
    if student_prefs_sheet_id:
        links_to_input[1], _ = build_hyperlink_to_sheet(student_prefs_sheet_id,
                                                        "Student Prefs")
    if instructor_prefs_sheet_id:
        links_to_input[2], _ = build_hyperlink_to_sheet(
            instructor_prefs_sheet_id, "Instructor Prefs")
    return links_to_input


def build_links_to_output(executed_num: str, num_alternate_matchings: int):
    def _build_hyperlink(sheet_title: str, text_prefix: str, num_alternate=0):
        sheet = get_sheet(sheet_title)
        ws_title = executed_num
        if num_alternate:
            ws_title += chr(ord('A') + num_alternate - 1)
        worksheet_id = sheet.worksheet(ws_title).id
        link_to_output, _ = build_hyperlink_to_sheet(sheet.id,
                                                     f"{text_prefix} #{executed_num}",
                                                     worksheet_id)
        return link_to_output

    links_to_output = [
        _build_hyperlink(gs_consts.MATCHING_OUTPUT_SHEET_TITLE, 'Matching'),
        _build_hyperlink(gs_consts.ADDITIONAL_TA_OUTPUT_SHEET_TITLE,
                         'Additional TA'),
        _build_hyperlink(gs_consts.REMOVE_TA_OUTPUT_SHEET_TITLE, 'Remove TA')]
    links_to_alternate_output = []
    for i in range(1, num_alternate_matchings + 1):
        links_to_alternate_output.append(
            _build_hyperlink(gs_consts.ALTERNATES_OUTPUT_SHEET_TITLE,
                             f'Alternate{i}', num_alternate=i))
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
    format_row(worksheet, 1, len(matrix[0]) + 1, center_align=False, bold=True,
               wrap=wrap)
    if wrap:
        wrap_rows(worksheet, 1, len(matrix), len(matrix[0]))


def write_csv_to_new_tab(csv_path: str, sheet_name: str, tab_name: str,
                         tab_index=0, center_align=False, wrap=False):
    with open(csv_path, newline='') as csv_file:
        reader = csv.reader(csv_file, delimiter=',', quotechar='"')
        print(f'Writing {csv_path} to {tab_name} in sheet {sheet_name}')
        matrix = list(reader)
        sheet = get_sheet(sheet_name)
        worksheet = write_matrix_to_new_tab_from_sheet(matrix, sheet, tab_name,
                                                       wrap, tab_index)
        resize_worksheet_columns(sheet, worksheet, len(matrix))
        if center_align:
            format(worksheet, 1, len(matrix), 0, len(matrix[0]),
                   center_align=center_align)


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
                if not remove_ws(alternates_sheet, alternates_worksheets,
                                 alternate_title):
                    return

    matching_sheet = get_sheet(gs_consts.MATCHING_OUTPUT_SHEET_TITLE)
    remove_ws(matching_sheet, matching_sheet.worksheets(), tab_num)
    remove_TA_sheet = get_sheet(gs_consts.REMOVE_TA_OUTPUT_SHEET_TITLE)
    remove_ws(remove_TA_sheet, remove_TA_sheet.worksheets(), tab_num)
    additional_TA_sheet = get_sheet(gs_consts.ADDITIONAL_TA_OUTPUT_SHEET_TITLE)
    remove_ws(additional_TA_sheet, additional_TA_sheet.worksheets(), tab_num)
    remove_alternates_ws(get_sheet(gs_consts.ALTERNATES_OUTPUT_SHEET_TITLE))

    toc_ws = get_worksheet_from_sheet(matching_sheet,
                                      gs_consts.OUTPUT_TOC_TAB_TITLE)
    max_ws = matching_sheet.worksheets()[1].title
    cells = get_rows_of_cells(toc_ws, 2, int(max_ws) + 2, 5)
    for i, row in enumerate(cells):
        if not row:
            continue
        title = row[4].value.split("#")
        if len(title) < 2:
            continue
        title = title[1]
        if int(title) <= int(tab_num):
            if tab_num in title:
                toc_ws.delete_rows(i + 2)
            return


def write_output_csvs(alternates, num_executed, output_dir_title):
    outputs_dir = f'{output_dir_title}/outputs'
    write_csv_to_new_tab(f'{outputs_dir}/matchings.csv',
                         gs_consts.MATCHING_OUTPUT_SHEET_TITLE, num_executed, 1)
    write_csv_to_new_tab(f'{outputs_dir}/additional_TA.csv',
                         gs_consts.ADDITIONAL_TA_OUTPUT_SHEET_TITLE,
                         num_executed, wrap=True)
    write_csv_to_new_tab(f'{outputs_dir}/remove_TA.csv',
                         gs_consts.REMOVE_TA_OUTPUT_SHEET_TITLE, num_executed,
                         wrap=True)
    for i in range(alternates):
        write_csv_to_new_tab(f'{outputs_dir}/alternate{i + 1}.csv',
                             gs_consts.ALTERNATES_OUTPUT_SHEET_TITLE,
                             num_executed + chr(ord('A') + i), i)
