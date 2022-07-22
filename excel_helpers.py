from openpyxl import load_workbook


def get_worksheet(wb, worksheet_name):
    worksheet_names = wb.sheetnames
    sheet_index = worksheet_names.index(worksheet_name)
    wb.active = sheet_index  # this will activate the worksheet by index
    ws = wb.active
    print('Setting active worksheet to: ' + str(ws))
    return ws


def clear_worksheet(worksheet, cell_range):
    for row in worksheet[cell_range]:
        for cell in row:
            cell.value = None
