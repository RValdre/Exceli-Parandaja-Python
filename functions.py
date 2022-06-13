import os

import openpyxl
from openpyxl.styles import PatternFill
import numpy as np



def cell_string(wb, sheet_name, cell_name):
    sheet = wb[sheet_name]
    cell = sheet[cell_name]
    return cell.value

def is_formula(cell_name):
    if str(cell_name) is None:
        return False
    if str(cell_name)[0] == '=':
        return True


def cell_adress_to_row_col(cell_adress):

    return int(cell_adress[1:])-2, ord(cell_adress[0]) - ord('A')

def cell_change_colour(wb, sheet_name, cell_name, colour):
    sheet = wb[sheet_name]
    cell = sheet[cell_name]
    cell.fill = PatternFill(start_color=colour, end_color=colour, fill_type="solid")




def cell_write(wb, sheet_name, cell_name, value):

    sheet = wb[sheet_name]
    cell = sheet[cell_name]
    cell.value = value


def cell_answer(file_path, sheet_name, cell_name):
    wb2 = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb2[sheet_name]
    cell = sheet[cell_name]
    return cell.value


def delete_excel_table_formating(wb, sheet_name):
    sheet = wb[sheet_name]
    for row in sheet.iter_rows():
        for cell in row:
            cell.style = 'Normal'

def delete_excel_cell_formating(wb, sheet_name, cell_name):
    sheet = wb[sheet_name]
    cell = sheet[cell_name]
    cell.style = 'Normal'

def unique(list1):
    x = np.array(list1)
    return np.unique(x)

def read_files_from_folder(folder_path):
    files = []
    for file in os.listdir(folder_path):
        if file.endswith(".xlsx"):
            files.append(file)
    return files

def sum_count(student_file,wb):
    lists = "sum, count"

    good = []
    bad = []

    not_formula = []
    formula_error = []
    wrong_answer = []

    count_if = dict(G29=4, G30=5, G31=8, G32=6, G33=9, G42=2, G43=2, G44=2, G45=9)

    sum_if = dict (G36=105, G37=164, G38=156, G39=511, G47=25, G48=75, G49=309, G52=386)



    for key, value in count_if.items():

        formula = cell_string(wb, lists, key)

        if is_formula(formula):
           if formula.find('COUNTIF') > 0 or formula.find('COUNTIFS') > 0:
               delete_excel_cell_formating(wb, lists, key)
               if cell_answer(student_file, lists, key) == value:
                   good.append(key)
               else:
                   bad.append(key)
                   wrong_answer.append(key)
           else:
               bad.append(key)
               formula_error.append(key)
        else:
            bad.append(key)
            not_formula.append(key)


    for key, value in sum_if.items():

        formula = cell_string(wb, lists, key)

        if is_formula(formula):
           if formula.find('SUMIF') > 0 or formula.find('SUMIFS') > 0:
               delete_excel_cell_formating(wb, lists, key)
               if cell_answer(student_file, lists, key) == value:
                   good.append(key)
               else:
                   bad.append(key)
                   wrong_answer.append(key)
           else:
               bad.append(key)
               formula_error.append(key)
        else:
            bad.append(key)
            not_formula.append(key)


    for i in good:
        cell_change_colour(wb, lists, i.replace("G", "F"), "33FF33")
        cell_change_colour(wb, lists, i, "33FF33")


    for i in bad:
        cell_change_colour(wb, lists, i.replace("G", "F"), "FF6666")
        cell_change_colour(wb, lists, i, "FF6666")

    for i in not_formula:
        cell_write(wb, lists, i.replace("G", "H"), "Not a formula")

    for i in formula_error:
        cell_write(wb, lists, i.replace("G", "H"), "Formula error")

    for i in wrong_answer:
        cell_write(wb, lists, i.replace("G", "H"), "Wrong answer")

def text_functions(student_file, wb):
    lists = "text functions"

    good = []
    bad = []

    not_formula = []
    wrong_answer = []
    formula_error = []

    fields = dict(B2= "MA", B3 = "CA", B4 = "CA", B5 = "AZ", B6="TX", B10="1BB2", B11 ="1PT", B12="1Z",B13="D",B14="1V24C", B15="1AA",B16="1ZFD3", C10="AC12",C11="AB34",C12="CD8",C13="PO65S3",C14="BV45",C15="DS96S",C16="CD90")
    if_fields = dict(B26=1300,D30="Pass",D31="Pass",D32="Fail",D33="Pass",D34="Pass",D35="Pass",D36="Fail")



    for key, value in fields.items():

        formula = cell_string(wb, lists, key)

        if is_formula(formula):
            delete_excel_cell_formating(wb, lists, key)
            if str(cell_answer(student_file, lists, key)) == str(value):
                good.append(key)
            else:
                bad.append(key)
                wrong_answer.append(key)
        else:
            bad.append(key)
            not_formula.append(key)

    for key, value in if_fields.items():
        formula = cell_string(wb, lists, key)

        if is_formula(formula):
            if formula.find('IF') > 0 or formula.find('IFS') > 0:
                delete_excel_cell_formating(wb, lists, key)
                if cell_answer(student_file, lists, key) == value:
                    good.append(key)
                else:
                    bad.append(key)
                    wrong_answer.append(key)
            else:
                bad.append(key)
                formula_error.append(key)
        else:
            bad.append(key)
            not_formula.append(key)


    for i in good:
        cell_change_colour(wb, lists, i, "33FF33")


    for i in bad:
        cell_change_colour(wb, lists, i, "FF6666")


    for i in not_formula:
        if str(i).find("B") != -1:
            cell_write(wb, lists, i.replace("B", "E"), "Not a formula")
            cell_change_colour(wb, lists, i.replace("B", "E"), "FDDA0D")
            cell_change_colour(wb, lists, i.replace("B", "F"), "FDDA0D")
        if str(i).find("C") != -1:
            cell_write(wb, lists, i.replace("C", "F"), "Not a formula")
            cell_change_colour(wb, lists, i.replace("C", "F"), "FDDA0D")
        if str(i).find("D") != -1:
            cell_write(wb, lists, i.replace("D", "G"), "Not a formula")
            cell_change_colour(wb, lists, i.replace("D", "G"), "FDDA0D")
            cell_change_colour(wb, lists, i.replace("D", "H"), "FDDA0D")

    for i in formula_error:
        if str(i).find("B") != -1:
            cell_write(wb, lists, i.replace("B", "E"), "Formula error")
            cell_change_colour(wb, lists, i.replace("B", "E"), "FDDA0D")
            cell_change_colour(wb, lists, i.replace("B", "F"), "FDDA0D")
        if str(i).find("C") != -1:
            cell_write(wb, lists, i.replace("C", "F"), "Formula error")
            cell_change_colour(wb, lists, i.replace("C", "F"), "FDDA0D")
        if str(i).find("D") != -1:
            cell_write(wb, lists, i.replace("D", "G"), "Formula error")
            cell_change_colour(wb, lists, i.replace("D", "G"), "FDDA0D")
            cell_change_colour(wb, lists, i.replace("D", "H"), "FDDA0D")

    for i in wrong_answer:
        if str(i).find("B") != -1:
            cell_write(wb, lists, i.replace("B", "E"), "Wrong answer")
            cell_change_colour(wb, lists, i.replace("B", "E"), "FDDA0D")
            cell_change_colour(wb, lists, i.replace("B", "F"), "FDDA0D")
        if str(i).find("C") != -1:
            cell_write(wb, lists, i.replace("C", "F"), "Wrong answer")
            cell_change_colour(wb, lists, i.replace("C", "F"), "FDDA0D")
        if str(i).find("D") != -1:
            cell_write(wb, lists, i.replace("D", "G"), "Wrong answer")
            cell_change_colour(wb, lists, i.replace("D", "G"), "FDDA0D")
            cell_change_colour(wb, lists, i.replace("D", "H"), "FDDA0D")


def date_functions(student_file, wb):
    lists = "Date functions"

    good = []
    bad = []

    not_formula = []
    formula_error = []
    wrong_answer = []

    dates = dict(E3="2020-03-01 00:00:00", E4="2019-06-03 00:00:00", E6=8, I3="2020-04-30 00:00:00",
                 I4="2019-03-03 00:00:00")
    week = dict(F3=1, F4=2)
    end = dict(G3="2020-04-30 00:00:00", G4="2019-07-31 00:00:00")
    func = dict(B8="=NOW()", D8="=HOUR(B8)", G8="=MINUTE(B8)")

    for key, value in dates.items():

        formula = cell_string(wb, lists, key)

        if is_formula(formula):
            if formula.find('DATE') > 0 or formula.find('DATES') > 0:
                delete_excel_cell_formating(wb, lists, key)
                if str(cell_answer(student_file, lists, key)) == str(value):
                    good.append(key)
                else:
                    bad.append(key)
                    wrong_answer.append(key)
            else:
                bad.append(key)
                formula_error.append(key)
        else:
            bad.append(key)
            not_formula.append(key)

    for key, value in week.items():

        formula = cell_string(wb, lists, key)

        if is_formula(formula):
            if formula.find('WEEKDAY') > 0 or formula.find('WEEKDAYS') > 0:
                delete_excel_cell_formating(wb, lists, key)
                if cell_answer(student_file, lists, key) == value:
                    good.append(key)
                else:
                    bad.append(key)
                    wrong_answer.append(key)
            else:
                bad.append(key)
                formula_error.append(key)
        else:
            bad.append(key)
            not_formula.append(key)

    for key, value in end.items():

        formula = cell_string(wb, lists, key)

        if is_formula(formula):
            if formula.find('EOMONTH') > 0 or formula.find('EOMONTHS') > 0:
                delete_excel_cell_formating(wb, lists, key)
                if str(cell_answer(student_file, lists, key)) == str(value):
                    good.append(key)
                else:
                    bad.append(key)
                    wrong_answer.append(key)
            else:
                bad.append(key)
                formula_error.append(key)
        else:
            bad.append(key)
            not_formula.append(key)

    for key, value in func.items():

        formula = cell_string(wb, lists, key)

        if is_formula(formula):
            delete_excel_cell_formating(wb, lists, key)
            if formula.find(value) != -1:
                good.append(key)
            else:
                bad.append(key)
                wrong_answer.append(key)
        else:
            bad.append(key)
            not_formula.append(key)

    formula_test = cell_string(wb, lists, "E5")
    if is_formula(formula_test):
        delete_excel_cell_formating(wb, lists, "E5")
        if cell_answer(student_file, lists, "E5") == 272:
            good.append("E5")
        else:
            bad.append("E5")
            wrong_answer.append("E5")
    else:
        bad.append("E5")
        not_formula.append("E5")

    for i in good:
        cell_change_colour(wb, lists, i, "33FF33")

    for i in bad:
        cell_change_colour(wb, lists, i, "FF6666")


def logical_functions(student_file, wb):
    lists = "Logical functions"

    good_if = []
    good_else = []
    bad_if = []
    bad_else = []

    not_formula_if = []
    formula_error_if = []
    wrong_answer_if = []
    not_formula_else = []
    formula_error_else = []
    wrong_answer_else = []

    if_condition = dict(D2=0.0325, D3='On target', D4='On target', D5=0.365, D6=0.54, D7=0.13, D8='On target', D9=1.27)
    if_and_condition = dict(D19=133, D20=250, D21=700, D22=267, D23=300, D24=300)
    if_average_condition = dict(E19="decommissioned", E20="decommissioned", E21="ok", E22="decommissioned", E23="ok",
                                E24="ok")

    for key, value in if_condition.items():
        formula = cell_string(wb, lists, key)


        if is_formula(formula):
            if formula.find('IF') != -1:
                delete_excel_cell_formating(wb, lists, key)
                if cell_answer(student_file, lists, key) == value:
                    good_if.append(key)
                else:
                    bad_if.append(key)
                    wrong_answer_if.append(key)
            else:
                bad_if.append(key)
                formula_error_if.append(key)
        else:
            bad_if.append(key)
            not_formula_if.append(key)

    for key, value in if_and_condition.items():
        formula = cell_string(wb, lists, key)


        if is_formula(formula):
            if formula.find('IF') != -1 and formula.find('AND') != -1:
                delete_excel_cell_formating(wb, lists, key)
                if round(cell_answer(student_file, lists, key)) == value:
                    good_else.append(key)
                else:
                    bad_else.append(key)
                    wrong_answer_else.append(key)
            else:
                bad_else.append(key)
                formula_error_else.append(key)
        else:
            bad_else.append(key)
            not_formula_else.append(key)

    for key, value in if_average_condition.items():
        formula = cell_string(wb, lists, key)


        if is_formula(formula):
            if formula.find('IF') != -1 and formula.find('AVERAGE') != -1:
                delete_excel_cell_formating(wb, lists, key)
                if cell_answer(student_file, lists, key) == value:
                    good_else.append(key)
                else:
                    bad_else.append(key)
                    wrong_answer_else.append(key)
            else:
                bad_else.append(key)
                formula_error_else.append(key)
        else:
            bad_else.append(key)
            not_formula_else.append(key)

    for i in good_if:
        cell_change_colour(wb, lists, i, "33FF33")

    for i in bad_if:
        cell_change_colour(wb, lists, i, "FF6666")

    for i in not_formula_if:
        cell_write(wb, lists, i.replace("D", "E"), "Not a formula")
        cell_change_colour(wb, lists, i.replace("D", "E"), "FDDA0D")

    for i in formula_error_if:
        cell_write(wb, lists, i.replace("D", "E"), "Formula error")
        cell_change_colour(wb, lists, i.replace("D", "E"), "FDDA0D")

    for i in wrong_answer_if:
        cell_write(wb, lists, i.replace("D", "E"), "Wrong answer")
        cell_change_colour(wb, lists, i.replace("D", "E"), "FDDA0D")

    for i in good_else:
        cell_change_colour(wb, lists, i, "33FF33")

    for i in bad_else:
        cell_change_colour(wb, lists, i, "FF6666")

    for i in not_formula_else:
        if i.find("D") != -1:
            cell_write(wb, lists, i.replace("D", "F"), "Not a formula")
            cell_change_colour(wb, lists, i.replace("D", "F"), "FDDA0D")
        if i.find("E") != -1:
            cell_write(wb, lists, i.replace("E", "G"), "Not a formula")
            cell_change_colour(wb, lists, i.replace("E", "G"), "FDDA0D")

    for i in formula_error_else:
        if i.find("D") != -1:
            cell_write(wb, lists, i.replace("D", "F"), "Formula error")
            cell_change_colour(wb, lists, i.replace("D", "F"), "FDDA0D")
        if i.find("E") != -1:
            cell_write(wb, lists, i.replace("E", "G"), "Formula error")
            cell_change_colour(wb, lists, i.replace("E", "G"), "FDDA0D")

    for i in wrong_answer_else:
        if i.find("D") != -1:
            cell_write(wb, lists, i.replace("D", "F"), "Wrong answer")
            cell_change_colour(wb, lists, i.replace("D", "F"), "FDDA0D")
        if i.find("E") != -1:
            cell_write(wb, lists, i.replace("E", "G"), "Wrong answer")
            cell_change_colour(wb, lists, i.replace("E", "G"), "FDDA0D")
