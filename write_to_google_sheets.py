import csv
import datetime
import gspread

from oauth2client.client import GoogleCredentials


def get_worksheet_from_sheet(sheet, worksheet_title):
    worksheets = [e.title for e in sheet.worksheets()]
    if not worksheet_title in worksheets:
        raise ValueError(f"{worksheet_title} not a worksheet in {sheet.title}")
    return sheet.worksheet(worksheet_title)


def get_worksheet_from_matchings_sheet(sheet):
    worksheets = [e.title for e in sheet.worksheets()]
    worksheets.remove('ToC')
    worksheets.sort(key=int, reverse=True)
    worksheet_title = "001"
    if len(worksheets) > 0:
        worksheet_title = str(f"{(int(worksheets[0]) + 1):03d}")
    return sheet.add_worksheet(title=worksheet_title, rows="100", cols="26", index=1)


def get_worksheet(sheet_title, worksheet_title=None):
    sheet = get_sheet(sheet_title)
    if worksheet_title:
        return get_worksheet_from_sheet(sheet, worksheet_title)
    return get_worksheet_from_matchings_sheet(sheet)


def get_sheet(sheet_title):
    gc = gspread.authorize(GoogleCredentials.get_application_default())
    return gc.open(sheet_title)


def get_sheet_by_id(sheet_id):
    gc = gspread.authorize(GoogleCredentials.get_application_default())
    return gc.open_by_key(sheet_id)


def wrap_rows(worksheet, row_start, row_end, cells):
    """ row_start and row_end are inclusive """
    format(worksheet, row_start, row_end, 0, cells,
           center_align=False, bold=False, wrap=True)


def format(worksheet,  start_row, end_row, start_col, end_col, center_align=True, bold=False, wrap=False):
    """start_col and end_col are zero indexed numbers"""
    formatting = {"textFormat": {"bold": bold}}
    if center_align:
        formatting["horizontalAlignment"] = "CENTER"
    if wrap:
        formatting["wrapStrategy"] = "WRAP"
    worksheet.format(
        f"{chr(ord('A') + start_col)}{start_row}:{chr(ord('A') + end_col)}{end_row}", formatting)


def format_row(worksheet, row, cells, center_align=True, bold=False, wrap=False):
    format(worksheet, row, row, 0, cells, center_align, bold, wrap)


def write_csv_to_worksheet(csv_path, sheet, worksheet):
    with open(csv_path, newline='') as csv_file:
        reader = csv.reader(csv_file, delimiter=',', quotechar='"')
        print(f'Writing {csv_path} to {worksheet.title}')
        row = 1
        length = 0
        for entry in reader:
            length = max(length, len(entry) + 1)
            cells = get_row_of_cells(worksheet, row, len(entry) + 1)
            update_cells(worksheet, cells, entry)
            format_row(worksheet, row, len(entry)+1,
                       center_align=True, bold=(row == 1))
            row += 1
        resize_worksheet_columns(sheet, worksheet, length)
        worksheet.freeze(rows=1)


def resize_worksheet_columns(sheet, worksheet, cols):
    body = {
        "requests": [
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": worksheet._properties['sheetId'],
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": cols
                    },
                }
            }
        ]
    }
    sheet.batch_update(body)


def update_cells(worksheet, cells, values):
    for i in range(len(values)):
        cells[i].value = values[i]
    worksheet.update_cells(cells, value_input_option='USER_ENTERED')


def get_row_of_cells(worksheet, row, cols_len):
    return worksheet.range(f"A{row}:{chr(ord('A') + cols_len)}{row}")


def append_to_last_row(worksheet, values):
    list_of_lists = worksheet.get_all_values()
    cells = get_row_of_cells(worksheet, len(list_of_lists) + 1, len(values))
    update_cells(worksheet, cells, values)


def write_execution_to_ToC(executor, executed_ws_title, executed_id):
    now = datetime.datetime.now()
    date = now.strftime('%m-%d-%Y')
    time = now.strftime('%H:%M:%S')
    sheet = get_sheet('TA-Matching Output')
    toc_wsh = get_worksheet_from_sheet(sheet, 'ToC')
    output_url = f"https://docs.google.com/spreadsheets/d/{sheet.id}/edit#gid={executed_id}"
    toc_vals = [date, time, executor, '',
                f"=HYPERLINK(\"{output_url}\", \"Run #{executed_ws_title}\")"]
    toc_wsh.insert_row(toc_vals, 2, value_input_option='USER_ENTERED')
    print(
        f"Wrote table of contents entry for sheet {executed_ws_title} ({output_url})")


def write_matrix_to_sheet(matrix, sheetname, worksheet_name=None, wrap=False):
    gc = gspread.authorize(GoogleCredentials.get_application_default())
    sheet = gc.create(sheetname)
    if worksheet_name:
        inital_worksheet = sheet.get_worksheet(0)
        rows = max(200, len(matrix) + 50)
        cols = max(10, len(matrix[0])+3)
        worksheet = sheet.add_worksheet(
            title=worksheet_name, rows=str(rows), cols=str(cols))
        sheet.del_worksheet(inital_worksheet)
    else:
        worksheet = sheet.get_worksheet(0)
    print(
        f"Created new spreadsheet at https://docs.google.com/spreadsheets/d/{sheet.id}")
    write_full_worksheet(matrix, worksheet, wrap)
    return sheet.id


def write_full_worksheet(matrix, worksheet, wrap):
    worksheet.update("A1", matrix)
    worksheet.freeze(rows=1)
    format_row(worksheet, 1, len(matrix[0])+1,
               center_align=False, bold=True, wrap=wrap)
    if wrap:
        wrap_rows(worksheet, 1, len(matrix), len(matrix[0]))


def write_matrix_to_new_tab(matrix, sheetname, tab_name, wrap=False):
    sheet = get_sheet(sheetname)
    rows = max(200, len(matrix) + 50)
    cols = max(10, len(matrix[0])+3)
    worksheet = sheet.add_worksheet(
        title=tab_name, rows=str(rows), cols=str(cols))
    write_full_worksheet(matrix, worksheet, wrap)
