months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]

jaanuar = dict(Week1=[1, 1, 272, 421, 272])
veebruar = dict(Week1=[1, 1, 251, 421, 251])
march = dict(Week1=[1, 1, 339, 421, 339])
april = dict(Week1=[1, 1, 412, 421, 412])
mai = dict(Week1=[1, 1, 16, 421, 16])
june = dict(Week1=[1, 1, 137, 421, 137])

cringe = [jaanuar, veebruar, march, april, mai, june]

excel_cell_month_value = "Apr"
excel_cell_week_value = "Week1"

index = months.index(excel_cell_month_value)
print(index)

dictionary_cringe = cringe[index]
print(dictionary_cringe)

for key, value in dictionary_cringe.items():
    print(key)
    print(value[2])


