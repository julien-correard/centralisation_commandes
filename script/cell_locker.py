import openpyxl
from openpyxl.styles import Protection

from pathlib import Path
from excel_writer import save_workbook

def check_if_sheet_locked(sheet, file: Path, protect: bool):
    if protect == True:
        sheet.protection.sheet = True

        for row in sheet.iter_rows():
            for cell in row:
                cell.protection = Protection(locked=True)

        wb = sheet.parent
        save_workbook(wb, file)



def check_if_cell_unlocked(sheet, row: int, col: int, protect: bool):
    if protect == True:
        cell = sheet.cell(row=row, column=col)   
        cell.protection = Protection(locked=False)
