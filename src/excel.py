from openpyxl import load_workbook

def create_ws(path: str, sheet_name: str, save_as: str):
    wb = load_workbook(filename = path)
    ws = wb.create_sheet(sheet_name, 0)
    wb.save(save_as)