import openpyxl
import pandas
import pandas as pd
import xlrd
import xlutils
from openpyxl.styles import PatternFill



def cell_string(file_path, sheet_name, cell_name):
    wb = openpyxl.load_workbook(file_path)
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

def cell_change_colour(file_path, sheet_name, cell_name, colour):
    """
    Changes the colour of a cell to a different colour
    """
    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]
    cell = sheet[cell_name]
    cell.fill = PatternFill(start_color=colour, end_color=colour, fill_type="solid")
    wb.save(file_path)
    wb.close()




def cell_write(file_path, sheet_name, cell_name, value):
    """
    Writes a value to a cell
    """
    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]
    cell = sheet[cell_name]
    cell.value = value
    wb.save(file_path)


def cell_answer(file_path, sheet_name, cell_name):
    wb2 = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb2[sheet_name]
    cell = sheet[cell_name]
    return cell.value


def delete_excel_table_formating(file_name, sheet_name):
    """
    Delete all Excel table formating
    """
    wb = openpyxl.load_workbook(file_name)
    sheet = wb[sheet_name]
    for row in sheet.iter_rows():
        for cell in row:
            cell.style = 'Normal'
    wb.save(file_name)
    wb.close()


