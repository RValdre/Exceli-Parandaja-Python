from functions import *

student_file = r"homework1-1 (answers).xlsx"
wb = openpyxl.load_workbook(student_file)

sum_count(student_file, wb)
text_functions(student_file, wb)

wb.save(student_file)
wb.close()