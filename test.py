
from lookup import *
student_file = r"homework1-2 (answers).xlsx"
wb = openpyxl.load_workbook(student_file)
lists = "lookup functions"
lookup_functions(student_file,wb)

wb.save(student_file)
wb.close()
