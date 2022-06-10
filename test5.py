import xlcalculator

file = r'homework1-1.xlsx'
sheet = 'Sheet1'
cell='A1'

def evalue(file,sheet,cell):
    xlcalculator.evalue(file,sheet,cell)

evalue(file,sheet,cell)