import formulas
func = formulas.Parser().ast('=(1 + 1) + 2 / 3')[1]

print(func.value)