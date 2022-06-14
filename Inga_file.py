import openpyxl as xl
wbook = xl.load_workbook("homework1-1 (answers).xlsx")
sheet = wbook.active

for data_val in sheet.data_validations.dataValidation:
    temp_data = str(data_val)
    info = [temp_data.split(", ")[0],temp_data.split(", ")[12], temp_data.split(", ")[13], temp_data.split(", ")[14]]
    print(temp_data)
wbook.close()