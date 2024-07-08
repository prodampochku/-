#import matplotlib
#import csv
import openpyxl
#import pandas as pd

class Profiles:
    def __init__(self, cy, cx, cm): #''', l, b0, bk''' ): #создаём класс, вводим коэффициенты
        self.cy = cy
        self.cx = cx
        self.cm = cm
        '''
        self.l = l
        l = float(input()) #размах крыла
        self.b0 = bo
        b0 = float(input()) #корневая хорда
        self.bk = bk
        bk = float(input()) #концевая хорда
        self.s = 0.5 * l * (b0 + bk) #площадь крыла
        self.udlinenie = ((l)**2)/self.s #удлинение
        self.suzhenie = b0/bk #сужение'''
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
                if self.cy == float(ws.cell(row=i, column=3).value) and self.cx == float(ws.cell(row=i, column=4).value) 
                and self.cm == float(ws.cell(row=i, column=5).value):
                    print(ws.cell(row=1,column=j).value, ws.cell(row=i, column=j).value) #нашёл - выводит название профиля и характеристики
                    ''''''return ws.cell(row=1,column=j).value
                    return ws.cell(row=i, column=j).value''''''
                else:
                    break''' #не нашёл - переходит на следующую строчку

    def srav(self, cy, cx, cm, val): #пробная функция для приближённого поиска. В процессе...
        wb = openpyxl.load_workbook('data.xlsx')
        ws = wb.active
        #ищем строк, которые теоритчески могут подойти, сохраняем номер строк из бд, которые всё ближе и ближе к вводимым данным
        for i in range(2,20):
            for j in range(1,6):
                if (cy - float(ws.cell(row=i, column=3).value))**2 < (self.dy)**2: # разность целевого Cy и
                    # проверяемого возводим в квадрат и сравниваем с квадратом разности-ориентира
                    self.dy = abs(cy - float(ws.cell(row=i, column=3).value)) # если меньше разности-ориентира,
                    # то модуль разности целевого и проверяемого Cy становится новой разностью-ориентиром
                    self.sp_cy.append(i) # в список sp_cy добавляем номер строки, которая проверялась
                elif (cy - float(ws.cell(row=i, column=3).value)) ** 2 >= (self.dy) ** 2: # если больше
                    # разности-ориентира, то ничего не делает программа
                    pass

                if (cx - float(ws.cell(row=i, column=4).value))**2 < (self.dx)**2: # разность целевого Cx и
                    # проверяемого возводим в квадрат и сравниваем с квадратом разности-ориентира
                    self.dx = abs(cx - float(ws.cell(row=i, column=4).value)) # если меньше разности-ориентира,
                    # то модуль разности целевого и проверяемого Cx становится новой разностью-ориентиром
                    self.sp_cx.append(i) # в список sp_cx добавляем номер строки, которая проверялась
                elif (cx - float(ws.cell(row=i, column=4).value)) ** 2 >= (self.dx) ** 2: # если больше разности-ориентира, то ничего не делает программа
                    pass

                if (cm - float(ws.cell(row=i, column=5).value)) ** 2 < (self.dm) ** 2: # разность целевого Cm и
                    # проверяемого возводим в квадрат и сравниваем с квадратом разности-ориентира
                    self.dm = abs(cm - float(ws.cell(row=i, column=5).value)) # если меньше разности-ориентира,
                    # то модуль разности целевого и проверяемого Cm становится новой разностью-ориентиром
                    self.sp_cm.append(i) # в список sp_cm добавляем номер строки, которая проверялась
                elif (cm - float(ws.cell(row=i, column=5).value)) ** 2 >= (self.dm) ** 2: # если больше разности-ориентира, то ничего не делает программа
                    pass
        print(self.sp_cy, self.sp_cx, self.sp_cm) #выводит номера строк, в которых были подходящие профили в порядке улучшения (3 списка для Сy, Cx, Cm)

        #ищем наиболее близкое геометрически (корень из суммы квадратов)
        self.val = val #переменная, которая хранит локальный номер строки
        self.m = 9999.9 #переменная для сравнения
        self.m1 = float() #переменная для записи
        self.k1 = 1 #переменная, которая хранит номер строки с наиближайшими коэффициентами
        for i in range(0, len(self.sp_cy)): #прогон по списку с Cy
            val = self.sp_cy[i] #переменной val присваиваем i-е значение из списка sp_cy
            self.m1 = (((float(ws.cell(row=val, column=3).value) - float(cy))**2) + (
                    (float(ws.cell(row=val, column=4).value) - float(cx))** 2) + ((float(ws.cell(row=val, column=5).value) - float(cm))** 2))** 0.5
            #переменной m1 присваиваем расстояние между точками. Координаты 1-й точки - коэффициенты проверяемой строки, 2-й - введённые коэффициенты
            if self.m1 < self.m: # если расстояние между точками меньше хрени для сравнения, то сравнивать будет
                # теперь с вычисленным расстоянием между точками (m - минимум, с которым сравниваем и ищем меньшее),
                # иначе ничего не делаем
                self.m = self.m1
                self.k1 = val #сохраняем лучший пока что номер строки

        for j in range(0, len(self.sp_cx)): #аналогично проверяем списки с номерами строк близких cx и cm к введённым
            val = self.sp_cx[j]
            self.m1 = (((float(ws.cell(row=val, column=3).value) - float(cy)) ** 2) + (
                    (float(ws.cell(row=val, column=4).value) - float(cx)) ** 2) + ((float(ws.cell(row=val, column=5).value) - float(cm)) ** 2)) ** 0.5
            if self.m1 < self.m:
                self.m = self.m1
                self.k1 = val

        for k in range(0, len(self.sp_cm)):
            val = self.sp_cm[k]
            self.m1 = (((float(ws.cell(row=val, column=3).value) - float(cy)) ** 2) + (
                    (float(ws.cell(row=val, column=4).value) - float(cx)) ** 2) + ((float(ws.cell(row=val, column=5).value) - float(cm)) ** 2)) ** 0.5
            if self.m1 < self.m:
                self.m = self.m1
                self.k1 = val

        #вывод
        for n in range(1,6): #найденный номер наиближайшей строки хранится в k1, выводим значения из строки k1 и столбцов с коэффициентами
            print(ws.cell(row=self.k1, column=n).value)
        self.m = 9999.9 #возвращаем крупное значение для новых прогонов
        self.k1 = 1 #аналогично, чтобы при следующих прогонах не использовались старые данные
        self.m1 = 0

#вызов функций
a = Profiles(cy=float, cx=float, cm=float) #класс профилей, его параметры cy, cx, cm, они вещественные
'''a.poisk()''' #вызываем функцию поиска полного совпадения
a.srav(cy=float(input()), cx=float(input()), cm=float(input()), val=int) #вызываем функцию поиска приближённого, говоря, что параметры будем вводить с клавиатуры

