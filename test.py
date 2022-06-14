import openpyxl as xl

wbook = xl.load_workbook('a.xlsx')
sheet = wbook.active

form = sheet.conditional_formatting

for row in form:
    print("Lahtrid:", row.cells, "Reeglid: ", row.cfRule)
wbook.close()