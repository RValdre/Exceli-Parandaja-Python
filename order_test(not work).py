from functions import *
import random


student_file = r"homework1-1 (answers).xlsx"
wb = openpyxl.load_workbook(student_file)
lists = "Validation"

order = ["A3","A4","A5","A6","A7"]
answers = []

for i in order:
    answers.append(cell_string(wb, lists, i))

for i in order:
    cell_write(wb, lists, i, "AA-1001")

if len(unique(answers)) == len(answers):
    for i in order:
        try:
            cell_write(wb, lists, i, "sdasdsadasdas")
            cell_change_colour(wb, lists, i, "FF6666")
        except:
            cell_change_colour(wb, lists, i, "33FF33")

else:
    for i in order:
        cell_change_colour(wb, lists, i, "FF6666")

wb.save(student_file)
wb.close()