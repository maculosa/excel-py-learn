from openpyxl import Workbook
from openpyxl.styles import Font

wb = Workbook()
mysheet = wb.get_sheet_by_name('Sheet')

firstFontObj = Font(name='Arial', bold=True)
mysheet['F6'].font = firstFontObj
mysheet['F6'] = 'Bold Arial'

secondFontObj = Font(size=32, italic=True)
mysheet['D7'].font = secondFontObj
mysheet['D7'] = '32 pt Italic'

wb.save('excel-6-settingFonts.xlsx')

