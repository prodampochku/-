import matplotlib
import csv
import openpyxl
import pandas as pd

class Profiles:
    def __init__(self): #создаём класс, вводим коэффициенты
        self.cy = float(input())
        self.cx = float(input())
        self.cm = float(input())
        self.l = float(input()) #размах крыла
        self.b0 = float(input()) #корневая хорда
        self.bk = float(input()) #концевая хорда
        self.s = 0.5 * self.l * (self.b0 + self.bk) #площадь крыла
        self.udlinenie = ((self.l)**2)/self.s #удлинение
        self.suzhenie = self.b0/self.bk #сужение
        self.sp = []
        self.r = []


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
                if  ws.cell(row=i, column=3).value ismin:
                    print(ws.cell(row=1,column=j).value, ws.cell(row=i, column=j).value)
                else:
                    break
a = Profiles()
'''a.poisk()'''
a.srav()
