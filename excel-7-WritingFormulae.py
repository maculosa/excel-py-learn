from openpyxl import Workbook
from openpyxl.styles import Font

wb = Workbook()
mysheet = wb.active
mysheet['F6'] = 500
mysheet['F7'] = 800

mysheet['D3'] = '=SUM(F6:F7)'
wb.save('excel-7-WritingFormulae.xlsx')
