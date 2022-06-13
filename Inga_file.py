import openpyxl as xl
wbook = xl.load_workbook("homework1-1 (answers).xlsx")
sheet = wbook.active

print("Data validation'i näitemine:")
for data_val in sheet.data_validations.dataValidation:
    print("Veerg:", data_val)
    for data in data_val:
        print(data)

wbook.close()