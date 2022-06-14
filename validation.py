from functions import *

student_file = r"homework1-1 (answers).xlsx"
wb = openpyxl.load_workbook(student_file)
sheet = wb.active
lists = "Validation"

good = []
bad = []

not_a_validation = []
bad_type= []
false_operator = []
wrong_formula = []
answers = [['<MultiCellRange [E3:E20]>', 'date', None, '$H$1', '$H$2'], ['<MultiCellRange [D3:D20]>', 'custom', None, 'ISTEXT(D3)', None], ['<MultiCellRange [C3:C20]>', 'decimal', 'greaterThanOrEqual', '0.0', None]]
validation_data = []
for data_val in sheet.data_validations.dataValidation:
    cell_data = []
    type = data_val.type
    adress = data_val.sqref
    operator = data_val.operator
    formula1 = data_val.formula1
    formula2 = data_val.formula2
    cell_data.extend([adress, type, operator, formula1, formula2])
    validation_data.append(cell_data)

for i in range(len(validation_data)):
    if (str(validation_data[i][0]).find("C3")) != -1:
        if str(validation_data[i][1]) == "decimal":
            if str(validation_data[i][2]) == "greaterThanOrEqual":
                if str(validation_data[i][3]) == "0.0":
                    for j in range(5):
                        good.append("C" + str(j + 3))
                else:
                    for j in range(5):
                        bad.append("C" + str(j + 3))
                    wrong_formula.append("C3")
            else:
                for j in range(5):
                    bad.append("C" + str(j + 3))
                wrong_formula.append("C3")
        else:
            for j in range(5):
                bad.append("C" + str(j + 3))
            bad_type.append("C3")
    if (str(validation_data[i][0]).find("E3")) != -1:
        if str(validation_data[i][1]) == "date":
            if str(validation_data[i][3]) == "$H$1" and str(validation_data[i][4]) == "$H$2":
                for j in range(5):
                    good.append("E" + str(j + 3))
            else:
                for j in range(5):
                    bad.append("E" + str(j + 3))
                wrong_formula.append("E3")
        else:
            for j in range(5):
                bad.append("E" + str(j + 3))
            bad_type.append("E3")

    if (str(validation_data[i][0]).find("D3")) != -1:
        if str(validation_data[i][1]) == "custom":
            if str(validation_data[i][3]) == 'ISTEXT(D3)':
                for j in range(5):
                    good.append("D" + str(j + 3))
            else:
                for j in range(5):
                    bad.append("D" + str(j + 3))
                wrong_formula.append("D3")
        else:
            for j in range(5):
                bad.append("D" + str(j + 3))
            bad_type.append("D3")

    for i in good:
        cell_change_colour(wb, lists, i, "33FF33")

    for i in bad:
        cell_change_colour(wb, lists, i, "FF6666")

    for i in not_a_validation:
        vals = str(i[0]) + str(int(i[1:])+7)
        cell_write(wb, lists, vals, "Not a validation")
        cell_change_colour(wb, lists, vals, "FDDA0D")

    for i in bad_type:
        vals = str(i[0]) + str(int(i[1:])+7)
        cell_write(wb, lists, vals, "Wrong Type")
        cell_change_colour(wb, lists, vals, "FDDA0D")

    for i in false_operator:
        vals = str(i[0]) + str(int(i[1:]) + 7)
        cell_write(wb, lists, vals, "False operator")
        cell_change_colour(wb, lists, vals, "FDDA0D")

    for i in wrong_formula:
        vals = str(i[0]) + str(int(i[1:]) + 7)
        cell_write(wb, lists, vals, "Wrong formula")
        cell_change_colour(wb, lists, vals, "FDDA0D")

wb.save(student_file)
wb.close()