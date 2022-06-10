from functions import *
from openpyxl import load_workbook

wb = load_workbook(r"homework1-1 (answers).xlsx")
student_file = r'TestFile.xlsx'
lists = "Лист1"

formula_dictionary = dict(G29=4, G30=5, G31=8, G32=6, G33=9, G36=105, G37=164, G38=156, G39=511, G42=2, G43=2, G44=2,
                          G45=9, G47=25, G48=75, G49=309, G52=389)

for key, value in formula_dictionary.items():

    formula = cell_string(student_file, lists, key)

    if is_formula(formula):
        print(formula)
        print(value)
        print(cell_answer(student_file, lists, key))
        if cell_answer(student_file, lists, key) == value:
            print("Good")

    else:
        print("not good")




