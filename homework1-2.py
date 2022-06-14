from functions import *

student_file = list_from_txt("uploaded-files-info.txt")
for i in student_file:
    if i.endswith("\n"):
        i = i[:-1]
    wb_start = r""
    file = wb_start + str(i)
    exea = file.replace("\\", "/")

    wb_start = r""

    wb = openpyxl.load_workbook(exea)

    logical_functions(student_file, wb)
    date_functions(student_file, wb)
    lookup_functions(student_file, wb)

    wb.save(exea)
    wb.close()
