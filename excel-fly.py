import xlwings as xw
app = xw.App(visible = True, add_book = False)
for i in range(1, 11):
    workbook = app.books.add()
    workbook.save(f'E:\\learning\\python\\example\\员工信息表\\分公司{i}.xlsx')
    workbook.close()
app.quit()
