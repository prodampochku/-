import openpyxl

cy = float(input()) #вводим с клавиатуры коэффициенты
cx = float(input())
cm = float(input())
number = int #переменная, в которую будет перезаписываться наиближайшая строчка
srav = 9999.9 #разность-ориентир, с которой сравниваем. очень большая, чтобы потом при сравнении с любой меньшей разностью последнюю сделать ориентиром


wb = openpyxl.load_workbook('data.xlsx') #доступ к эксель странице
ws = wb.active

for i in range(2,20): #пробегаем строчки. разности текуищх и введённых коэффициентов возводим в квадрат и изсуммы извлекаем корень
    a = (cy - float(ws.cell(row=i, column=3).value)) ** 2
    b = (cx - float(ws.cell(row=i, column=4).value)) ** 2
    c = (cm - float(ws.cell(row=i, column=5).value)) ** 2
    dist = (a + b + c)**0.5
    if dist < srav: #если геометрическое расстояние меньше ориентира, ориентир перезаписываем равным расстоянию
        srav = dist
        number = i # запоминаем номер строки

print(ws.cell(row=number, column=1).value) #профиль
print(ws.cell(row=number, column=2).value) #угол атаки
print(ws.cell(row=number, column=3).value)# Cy
print(ws.cell(row=number, column=4).value)# Cx
print(ws.cell(row=number, column=5).value)# Cm
