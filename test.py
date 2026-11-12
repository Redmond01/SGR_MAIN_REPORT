import xlwings as xw

app = xw.App(
    visible=False
)

wb = app.books.add()

sheet = wb.sheets[0]

sheet.range("A1").value = "HELLO"

wb.save(r"C:\Users\HP\Desktop\test.xlsx")

print(
    wb.api.ReadOnly
)

wb.close()

app.quit()