from functions import *

student_file = r"homework1-2 (answers).xlsx"
wb = openpyxl.load_workbook(student_file)

logical_functions(student_file, wb)
date_functions(student_file, wb)
lookup_functions(student_file, wb)

wb.save(student_file)
wb.close()