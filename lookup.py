from functions import *
def lookup_functions(student_file, wb):
    lists = "lookup functions"

    good = []
    bad = []

    not_formula = []
    formula_error = []
    wrong_answer = []

    look_up = dict(F6=200, F7=25, F8=150, F9=75, F10=750, F11=75, F12=300, F13=150, F14=25, F15=150, F16=150)

    for key, value in look_up.items():
        formula = cell_string(wb, lists, key)
        if is_formula(formula):
            if formula.find('VLOOKUP') != -1:
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
        cell_change_colour(wb, lists, i, "33FF33")

    for i in bad:
        cell_change_colour(wb, lists, i, "FF6666")

    for i in not_formula:
        if i.find("F") != -1:
            cell_write(wb, lists, i.replace("F", "E"), "Not a formula")
            cell_change_colour(wb, lists, i.replace("F", "E"), "FDDA0D")


    for i in formula_error:
        if i.find("F") != -1:
            cell_write(wb, lists, i.replace("F", "E"), "Formula error")
            cell_change_colour(wb, lists, i.replace("F", "E"), "FDDA0D")

    for i in wrong_answer:
        if i.find("F") != -1:
            cell_write(wb, lists, i.replace("F", "E"), "Wrong answer")
            cell_change_colour(wb, lists, i.replace("F", "E"), "FDDA0D")
