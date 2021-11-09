import csv


def get_worksheet(gc, sheet_title, worksheet_title=None):
    sheet = gc.open(sheet_title)
    if worksheet_title:
        if not worksheet_title in sheet.worksheets():
            raise ValueError(
                "{worksheet_title} not a worksheet in {sheet_title}")
        return sheet.worksheet(worksheet_title)

    worksheets = sheet.worksheets()
    worksheet_title = "1"
    if len(worksheets) > 1:
        worksheet_title = str(int(worksheets[1].title) + 1)
    return sheet.add_worksheet(title=worksheet_title, rows="100", cols="26", index=1)


def write_csv_to_worksheet(csv_path, worksheet):
    with open(csv_path, newline='') as csv_file:
        reader = csv.reader(csv_file, delimiter=',', quotechar='"')
        print(f'Writing {csv_path} to {worksheet}')
        row = 1
        for entry in reader:
            cells = worksheet.range(
                f"A{row}:{chr(ord('A') + len(entry)+1)}{row}")
            for i in range(len(entry)):
                cells[i].value = entry[i]
            worksheet.update_cells(cells)
            row += 1
