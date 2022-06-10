from functions import *

file = r'TestFile.xlsx'

temp_row_col = cell_adress_to_row_col("G30")
print(cell_string(file, "Лист1", "G30"))
print(cell_answer(file, "Лист1", "G30"))