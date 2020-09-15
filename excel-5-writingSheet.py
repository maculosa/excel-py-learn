import openpyxl

mywb = openpyxl.Workbook()
mysheet = mywb.get_sheet_by_name('Sheet')
mysheet['F6'] = 'Writing new Value!'

print(mysheet['F6'].value)
mywb.save('excel-5-writingSheet.xlsx')

