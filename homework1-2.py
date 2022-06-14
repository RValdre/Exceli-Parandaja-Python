from functions import *

name = "homework1_2"
student_file = list_from_txt("uploaded-files-info.txt")
create_zip(name)

for i in student_file:
    if i.endswith("\n"):
        i = i[:-1]
    if i.endswith("false"):
        i = i[:-5]
    wb_start = r""
    file = wb_start + str(i)
    exea = file.replace("\\", "/")
    copy_file(exea)
    exea = exea.replace(".xlsx", "_copy.xlsx")

    wb = openpyxl.load_workbook(exea)
    logical_functions(exea, wb)
    date_functions(exea, wb)
    lookup_functions(exea, wb)

    wb.save(exea)
    wb.close()

    add_file_to_zip_without_directory(exea,name)
    delete_file(exea)