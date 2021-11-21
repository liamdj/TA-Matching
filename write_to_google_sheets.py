import csv
import datetime
import gspread

from oauth2client.client import GoogleCredentials

gc = gspread.authorize(GoogleCredentials.get_application_default())


def get_worksheet_from_sheet(sheet, worksheet_title=None):
    worksheets = [e.title for e in sheet.worksheets()]
    if worksheet_title:
        if not worksheet_title in worksheets:
            raise ValueError(
                f"{worksheet_title} not a worksheet in {sheet.title}")
        return sheet.worksheet(worksheet_title)

    worksheets.remove('ToC')
    worksheets.sort(key=int, reverse=True)
    worksheet_title = "001"
    if len(worksheets) > 0:
        worksheet_title = str(f"{(int(worksheets[0]) + 1):03d}")
    return sheet.add_worksheet(title=worksheet_title, rows="100", cols="26", index=1)


def get_worksheet(sheet_title, worksheet_title=None):
    return get_worksheet_from_sheet(get_sheet(sheet_title), worksheet_title)


def get_sheet(sheet_title):
    return gc.open(sheet_title)


def format_row(worksheet, row, cells, center_align=True, make_bold=False):
    formatting = {"textFormat": {"bold": make_bold}}
    if center_align:
        formatting["horizontalAlignment"] = "CENTER"
    worksheet.format(
        f"A{row}:{chr(ord('A') + cells)}{row}", formatting)


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
                       center_align=True, make_bold=(row == 1))
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


def write_matrix_to_sheet(matrix, sheetname, worksheet_name=None):
    sheet = gc.create(sheetname)
    if worksheet_name:
        worksheet = sheet.add_worksheet(
            title=worksheet_name, rows=str(len(matrix)), cols=(len(matrix[0])))
    else:
        worksheet = sheet.get_worksheet(0)
    print(
        f"Created new spreadsheet at https://docs.google.com/spreadsheets/d/{sheet.id}")
    worksheet.update("A1", matrix)
    worksheet.freeze(rows=1)
    format_row(worksheet, 1, len(matrix[0])+1,
               center_align=False, make_bold=True)
