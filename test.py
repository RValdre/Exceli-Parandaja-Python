from functions import *

file = r'homework1-1 (answers).xlsx'

print(cell_string(r'homework1-1.xlsx', "sum, count", "G29"))
temp_row_col = cell_adress_to_row_col("G29")
print(temp_row_col)
data = pd.read_excel(r'homework1-1.xlsx', sheet_name='sum, count')

print(cell_value(r'homework1-1.xlsx', "sum, count", temp_row_col[0]-2, temp_row_col[1]))
print(data.iloc[23,6])


if is_formula(cell_string(r'homework1-1.xlsx', "sum, count", "G29")):

    if cell_value(r'homework1-1.xlsx', "sum, count", temp_row_col[0]-2, temp_row_col[1]) == 4:
        print("True")
        cell_change_colour(r'homework1-1.xlsx', "sum, count", "G29", "33FF33")
else:
    print("False")
    cell_change_colour(r'homework1-1.xlsx', "sum, count", "G29", "FF6666")
