from functions import *

file = r'homework1-1.xlsx'

temp_row_col = cell_adress_to_row_col("H30")
print(cell_string(file, "sum, count", "H30"))
print(cell_answer(file, "sum, count", "H30"))
#print(cell_answer(file, "sum, count", temp_row_col[0], temp_row_col[1]))
print(pd.read_excel(file, sheet_name="sum, count"))