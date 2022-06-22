from functions import create_zip, list_from_txt, validation_functions, sum_count, text_functions, \
    file_check, finish, end_of_testing, script_start, sheet_list_maker, delete_file
from openpyxl import load_workbook
from warnings import filterwarnings

filterwarnings("ignore")
name = "homework1_1"
create_zip(name)
count = 1
file_list = list_from_txt("uploaded-files-info.txt")
script_start(file_list)
try:
    for i in file_list:
        this_file = file_check(i)

        wb = load_workbook(this_file)
        sheet_names = sheet_list_maker(wb)

        validation_functions(wb, sheet_names)
        sum_count(this_file, wb, sheet_names)
        text_functions(this_file, wb, sheet_names)

        end_of_testing(wb, this_file, count, file_list, name)
        count = count + 1
    finish()
except:
    this_file = file_check(i)
    delete_file(this_file)
    print(i + " file is broken")
    input("Press enter to close:")
