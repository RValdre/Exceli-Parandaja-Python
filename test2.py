from functions import *
import openpyxl
import pandas as pd
from openpyxl import load_workbook
wb = load_workbook(r"homework1-1 (answers).xlsx")
student_file = r'homework1-1.xlsx'

ws = wb['sum, count']

formula_dictionary = dict(H29=4, H30=5, H31=8, H32=6, H33=9, H36=105, H37=164, H38=156, H39=511, H42=2, H43=2, H44=2, H45=9, H47=25, H48=75, H49=309, H52=389)


for key, value in formula_dictionary.items():
    print(value)
    formula = cell_string(student_file, "sum, count", key)
    print(formula)
    print(cell_adress_to_row_col(key))
    temp_row_col = cell_adress_to_row_col(key)
    print(cell_value(student_file, "sum, count", temp_row_col[0], temp_row_col[1]))
    if is_formula(formula):

        if value == cell_value(student_file, "sum, count",temp_row_col[0]+2, temp_row_col[1]):
            cell_change_colour(student_file, "sum, count", key, "33FF33")

    else:
        cell_change_colour(student_file, "sum, count", key, "FF6666")
