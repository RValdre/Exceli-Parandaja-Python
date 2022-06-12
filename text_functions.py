from functions import *
from openpyxl import load_workbook
def text_functions(student_file, wb):
    lists = "text functions"

    good = []
    bad = []

    not_formula = []
    wrong_answer = []
    formula_error = []

    fields = dict(B2= "MA", B3 = "CA", B4 = "CA", B5 = "AZ", B6="TX", B10="1BB2", B11 ="1PT", B12="1Z",B13="D",B14="1V24C", B15="1AA",B16="1ZFD3", C10="AC12",C11="AB34",C12="CD8",C13="PO65S3",C14="BV45",C15="DS96S",C16="CD90")
    if_fields = dict(B26=1300,D30="Pass",D31="Pass",D32="Fail",D33="Pass",D34="Pass",D35="Pass",D36="Fail")



    for key, value in fields.items():

        formula = cell_string(wb, lists, key)

        if is_formula(formula):
            delete_excel_cell_formating(wb, lists, key)
            if str(cell_answer(student_file, lists, key)) == str(value):
                good.append(key)
            else:
                bad.append(key)
                wrong_answer.append(key)
        else:
            bad.append(key)
            not_formula.append(key)

    for key, value in if_fields.items():
        formula = cell_string(wb, lists, key)

        if is_formula(formula):
            if formula.find('IF') > 0 or formula.find('IFS') > 0:
                delete_excel_cell_formating(wb, lists, key)
                if cell_answer(student_file, lists, key) == value:
                    good.append(key)
                else:
                    bad.append(key)
                    wrong_answer.append(key)
            else:
                bad.append(key)
                formula_error.append(key)
        else:
            bad.append(key)
            not_formula.append(key)


    for i in good:
        cell_change_colour(wb, lists, i.replace("G", "F"), "33FF33")
        cell_change_colour(wb, lists, i, "33FF33")


    for i in bad:
        cell_change_colour(wb, lists, i.replace("G", "F"), "FF6666")
        cell_change_colour(wb, lists, i, "FF6666")
