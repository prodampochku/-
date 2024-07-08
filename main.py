from PyQt5 import QtWidgets, uic

import matplotlib
import csv
import openpyxl
import pandas as pd


app = QtWidgets.QApplication([])
ui = uic.loadUi('design.ui')

list_angle = []              #Заполняем углы атаки
for i in range(0, 100, 10):
    list_angle.append(str(i))
    
ui.alpha_Box.addItems(list_angle)


def algoritm():
    
    ui.Cy_current.setText('0.900')
    ui.Cx_current.setText('0.065')
    ui.Mz_current.setText('0.250')

    cy = float(ui.Cy_current.text())
    cx = float(ui.Cx_current.text())
    cm = float(ui.Mz_current.text())

    wb = openpyxl.load_workbook('data.xlsx')
    ws = wb.active

        
    for i in range(2,20):
        for j in range(1,6):
            if cy == float(ws.cell(row=i, column=3).value) and cx == float(ws.cell(row=i, column=4).value) and cm == float(ws.cell(row=i, column=5).value):
                ui.name_profile_print.setText(str((ws.cell(row=i,column=1).value)))#, ws.cell(row=i, column=j).value)
                ui.Cy_finall.setText(str((ws.cell(row=i,column=3).value)))
                ui.Cx_finall.setText(str((ws.cell(row=i,column=4).value)))
                ui.Mz_finall.setText(str((ws.cell(row=i,column=5).value)))
            else:
                break
ui.btn1.clicked.connect(algoritm)
#ui.name_profile_print.setText('ЦАГИ')


ui.show()
app.exec()
