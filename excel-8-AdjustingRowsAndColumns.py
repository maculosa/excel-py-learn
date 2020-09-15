from openpyxl import Workbook
from openpyxl.styles import Font

wb = Workbook()
mysheet = wb.active
mysheet['F6'] = 'Tall row'
mysheet['F7'] = 'Wide column'
mysheet.row_dimensions[3].height = 65
mysheet.column_dimensions['F'].width = 25

wb.save('excel-8-AdjustingRowsAndColumns.xlsx')
