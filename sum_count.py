from functions import *
from openpyxl import load_workbook


student_file = r"homework1-1 (answers).xlsx"
wb = openpyxl.load_workbook(student_file)
lists = "sum, count"

good = []
bad = []

not_formula = []
formula_error = []
wrong_answer = []

count_if = dict(G29=4, G30=5, G31=8, G32=6, G33=9, G42=2, G43=2, G44=2, G45=9)

sum_if = dict (G36=105, G37=164, G38=156, G39=511, G47=25, G48=75, G49=309, G52=386)



for key, value in count_if.items():

    formula = cell_string(wb, lists, key)

    if is_formula(formula):
       if formula.find('COUNTIF') > 0 or formula.find('COUNTIFS') > 0:
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


for key, value in sum_if.items():

    formula = cell_string(wb, lists, key)

    if is_formula(formula):
       if formula.find('SUMIF') > 0 or formula.find('SUMIFS') > 0:
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

for i in not_formula:
    cell_write(wb, lists, i.replace("G", "H"), "Not a formula")

for i in formula_error:
    cell_write(wb, lists, i.replace("G", "H"), "Formula error")

for i in wrong_answer:
    cell_write(wb, lists, i.replace("G", "H"), "Wrong answer")

wb.save(student_file)
wb.close()