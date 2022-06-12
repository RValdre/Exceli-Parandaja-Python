from functions import *

student_file = r'homework1-1 (answers).xlsx'
wb = openpyxl.load_workbook(student_file)
delete_excel_cell_formating(wb, "sum, count", "G29")

wb.save(student_file)
wb.close()