import matplotlib
import csv
import openpyxl
import pandas as pd

class Profiles:
    def __init__(self):
        self.cy = float(input())
        self.cx = float(input())
        self.cm = float(input())
        self.l = float(input())
        self.b0 = float(input())
        self.bk = float(input())
        self.s = 0.5 * self.l * (self.b0 + self.bk)
        self.udlinenie = ((self.l)**2)/self.s
        self.suzhenie = self.b0/self.bk
        self.sp = []
        self.r = []


    '''def poisk(self):
        wb = openpyxl.load_workbook('data.xlsx')
        ws = wb.active
        for i in range(2,20):
            for j in range(1,6):
                if self.cy == float(ws.cell(row=i, column=3).value) and self.cx == float(ws.cell(row=i, column=4).value) and self.cm == float(ws.cell(row=i, column=5).value):
                    print(ws.cell(row=1,column=j).value, ws.cell(row=i, column=j).value)
                    ''''''return ws.cell(row=1,column=j).value
                    return ws.cell(row=i, column=j).value''''''
                else:
                    break'''

    def srav(self):
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
