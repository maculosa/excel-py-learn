import openpyxl

mywb = openpyxl.load_workbook('ExcelDemo1.xlsx')
sheet = mywb.active
sheet.title = 'Working on Save as'
mywb.save('example_filetest.xlsx')
