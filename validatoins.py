import openpyxl
import pandas as pd
from openpyxl import load_workbook
from functions import *
try:
    cell_write('homework1-1.xlsx (answers)', "Validation", "A4", "fgdfg")
    cell_change_colour('homework1-1 (answers).xlsx', "Validation", "A4", "FF6666")
except:
    cell_change_colour('homework1-1 (answers).xlsx', "Validation", "A4", "33FF33")