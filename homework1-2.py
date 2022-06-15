from functions import *
import time

name = "homework1_2"
student_file = list_from_txt("uploaded-files-info.txt")
create_zip(name)
count = 1
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
    conditional_function(exea, wb)


    wb.save(exea)
    wb.close()
    print(str(count) + "/" + str(len(student_file)) + " have been controlled")
    add_file_to_zip_without_directory(exea, name)
    delete_file(exea)
    count = count + 1
print("-------------------")
print("Files are zipped")

time.sleep(10)