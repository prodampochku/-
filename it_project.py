import matplotlib
import csv
import openpyxl
import pandas as pd

cy = float(input())
cx = float(input())
cm = float(input())

wb = openpyxl.load_workbook('data.xlsx')
ws = wb.active


for i in range(2,20):
    for j in range(1,6):
        if cy == float(ws.cell(row=i, column=3).value) and cx == float(ws.cell(row=i, column=4).value) and cm == float(ws.cell(row=i, column=5).value):
            print(ws.cell(row=1,column=j).value, ws.cell(row=i, column=j).value)
        else:
            break