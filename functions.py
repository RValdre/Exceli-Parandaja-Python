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