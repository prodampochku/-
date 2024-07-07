import matplotlib
import csv
import openpyxl
import pandas as pd

class Profiles:
    def __init__(self): #создаём класс, вводим коэффициенты
        self.cy = float(input())
        self.cx = float(input())
        self.cm = float(input())
        '''self.l = float(input()) #размах крыла
        self.b0 = float(input()) #корневая хорда
        self.bk = float(input()) #концевая хорда
        self.s = 0.5 * self.l * (self.b0 + self.bk) #площадь крыла
        self.udlinenie = ((self.l)**2)/self.s #удлинение
        self.suzhenie = self.b0/self.bk #сужение'''
        self.sp_cy = [] # списки для строк, которые нам нужны
        self.sp_cx = []
        self.sp_cm = []
        self.dy = 9999.9 #разности-ориентиры, с которыми сравниваем. очень большие, чтобы потом при сравнении с любой меньшей разностью последнюю сделать ориентиром
        self.dx = 9999.9
        self.dm = 9999.9


    '''def poisk(self): #рабочая функция, которая ищет полное совпадение, пробегает строчки, а в каждой строчке все столбцы
        wb = openpyxl.load_workbook('data.xlsx')
        ws = wb.active
        for i in range(2,20):
            for j in range(1,6):
                if self.cy == float(ws.cell(row=i, column=3).value) and self.cx == float(ws.cell(row=i, column=4).value) and self.cm == float(ws.cell(row=i, column=5).value):
                    print(ws.cell(row=1,column=j).value, ws.cell(row=i, column=j).value) #нашёл - выводит название профиля и характеристики
                    ''''''return ws.cell(row=1,column=j).value
                    return ws.cell(row=i, column=j).value''''''
                else:
                    break''' #не нашёл - переходит на следующую строчку

    def srav(self): #пробная функция для приближённого поиска. В процессе...
        wb = openpyxl.load_workbook('data.xlsx')
        ws = wb.active
        for i in range(2,20):
            for j in range(1,6):
                if (self.cy - float(ws.cell(row=i, column=3).value))**2 < (self.dy)**2: # разность целевого Cy и проверяемого возводим в квадрат и сравниваем с квадратом разности-ориентира
                    self.dy = abs(self.cy - float(ws.cell(row=i, column=3).value)) # если меньше разности-ориентира, то модуль разности целевого и проверяемого Cy становится новой разностью-ориентиром
                    self.sp_cy.append(i) # в список sp_cy добавляем номер строки, которая проверялась
                elif (self.cy - float(ws.cell(row=i, column=3).value)) ** 2 > (self.dy) ** 2: # если больше разности-ориентира, то ничего не делает программа
                    pass
                if (self.cx - float(ws.cell(row=i, column=4).value))**2 < (self.dx)**2: # разность целевого Cx и проверяемого возводим в квадрат и сравниваем с квадратом разности-ориентира
                    self.dx = abs(self.cx - float(ws.cell(row=i, column=4).value)) # если меньше разности-ориентира, то модуль разности целевого и проверяемого Cx становится новой разностью-ориентиром
                    self.sp_cx.append(i) # в список sp_cx добавляем номер строки, которая проверялась
                elif (self.cx - float(ws.cell(row=i, column=4).value)) ** 2 > (self.dx) ** 2: # если больше разности-ориентира, то ничего не делает программа
                    pass
                if (self.cm - float(ws.cell(row=i, column=5).value)) ** 2 < (self.dm) ** 2: # разность целевого Cm и проверяемого возводим в квадрат и сравниваем с квадратом разности-ориентира
                    self.dm = abs(self.cm - float(ws.cell(row=i, column=5).value)) # если меньше разности-ориентира, то модуль разности целевого и проверяемого Cm становится новой разностью-ориентиром
                    self.sp_cm.append(i) # в список sp_cm добавляем номер строки, которая проверялась
                elif (self.cm - float(ws.cell(row=i, column=5).value)) ** 2 > (self.dm) ** 2: # если больше разности-ориентира, то ничего не делает программа
                    pass



                '''print(ws.cell(row=1, column=j).value, ws.cell(row=i, column=j).value)'''
        print(self.sp_cy, self.sp_cx, self.sp_cm) #выводит номера строк, в которых были подходящие профили в порядке улучшения (3 списка для Сy, Cx, Cm)
a = Profiles()
'''a.poisk()'''
a.srav()

