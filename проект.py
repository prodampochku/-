from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow

import sys

class Start_window(QMainWindow):
    def __init__(self):
        super(Start_window, self).__init__()

        self.setWindowTitle('WingPro')
        self.setGeometry(0, 0, 1039, 662)


        self.hello_text = QtWidgets.QTextEdit(self)
        self.hello_text.setGeometry(190, 60, 650, 411)
        self.hello_text.setText('Добро пожаловать в WingPro!')
        

        self.btn = QtWidgets.QPushButton(self)
        self.btn.setGeometry(430, 520, 141, 51)
        self.btn.setText('Рассчитать')
        self.btn.clicked.connect(self.show_new_window)
        
    def show_new_window(self, checked):
        self.w = Main_Window()
        self.w.show()

class Main_window(QMainWindow):
    
    def __init__(self):
        super(Main_window, self).__init__()

        self.setWindowTitle('WingPro')
        self.setGeometry(0, 0, 1041, 662)

        

        self.Cy_current = QtWidgets.QLineEdit(self)
        self.Cy_current.setGeometry(50, 130, 141, 31)
        
        self.Cx_current = QtWidgets.QLineEdit(self)
        self.Cx_current.setGeometry(50, 270, 141, 31)
        
        self.Mz_current = QtWidgets.QLineEdit(self)
        self.Mz_current.setGeometry(50, 390, 141, 31)

        self.text_Cy = QtWidgets.QTextEdit(self)
        self.text_Cy.setGeometry(50, 60, 301, 41)
        self.text_Cy.setText('Коэффициент подъемной силы')

        self.text_Cx = QtWidgets.QTextEdit(self)
        self.text_Cx.setGeometry(50, 200, 361, 41)
        self.text_Cx.setText('Коэффициент лобового споротивления')
        

        self.text_Mz = QtWidgets.QTextEdit(self)
        self.text_Mz.setGeometry(50, 330, 301, 41)
        self.text_Mz.setText('Коэффициент момента тангажа')

        self.btn1 = QtWidgets.QPushButton(self)
        self.btn1.setGeometry(40, 560, 151, 41)
        self.btn1.setText('Подобрать профиль')
        self.btn1.clicked.connect(algoritm)

        self.alpha_Box = QtWidgets.QComboBox(self)
        self.alpha_Box.setGeometry(50, 510, 111, 31)
        list_angle = []              #Заполняем углы атаки
        for i in range(0, 100, 10):
            list_angle.append(str(i))
    
        self.alpha_Box.addItems(list_angle) 

        self.text_alpha = QtWidgets.QTextEdit(self)
        self.text_alpha.setGeometry(50, 440, 101, 41)
        self.text_alpha.setText('Угол атаки')

        self.name_profile_print = QtWidgets.QLineEdit(self)
        self.name_profile_print.setGeometry(820, 130, 161, 41)

        self.name_profile_text = QtWidgets.QTextEdit(self)
        self.name_profile_text.setGeometry(540, 130, 251, 41)
        self.name_profile_text.setText('Профиль крыла')


        self.Cy_finall = QtWidgets.QLineEdit(self)
        self.Cy_finall.setGeometry(820, 190, 161, 41)

        self.Cx_finall = QtWidgets.QLineEdit(self)
        self.Cx_finall.setGeometry(820, 250, 161, 41)

        self.Mz_finall = QtWidgets.QLineEdit(self)
        self.Mz_finall.setGeometry(820, 320, 161, 41)

        self.Cy_finall_text = QtWidgets.QTextEdit(self)
        self.Cy_finall_text.setGeometry(740, 190, 51, 41)
        self.Cy_finall_text.setText('Cy')

        self.Cx_finall_text = QtWidgets.QTextEdit(self)
        self.Cx_finall_text.setGeometry(740, 250, 51, 41)
        self.Cx_finall_text.setText('Cx')

        self.Mz_finall_text = QtWidgets.QTextEdit(self)
        self.Mz_finall_text.setGeometry(750, 320, 41, 41)
        self.Mz_finall_text.setText('Mz')
    
        
def algoritm():

    class Profiles(Main_window):
        def __init__(self, cy, cx, cm): #''', l, b0, bk''' ): #создаём класс, вводим коэффициенты
            self.Cy_current.setText('0.656')
            self.Cx_current.setText('0.0430')
            self.Mz_current.setText('0.155')
            cy = float(self.Cy_current.text())
            cx = float(self.Cx_current.text())
            cm = float(self.Mz_current.text())
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
            #print(self.sp_cy, self.sp_cx, self.sp_cm) #выводит номера строк, в которых были подходящие профили в порядке улучшения (3 списка для Сy, Cx, Cm)

            #ищем наиболее близкое геометрически (корень из суммы квадратов)
            self.val = val #переменная, которая хранит локальный номер строки
            self.m = 9999.9 #переменная для сравнения
            self.m1 = float() #переменная для записи
            self.k1 = 1 #переменная, которая хранит номер строки с наиближайшими коэффициентами
            for i in range(0, len(self.sp_cy)): #прогон по списку с Cy
                val = self.sp_cy[i] #переменной val присваиваем i-е значение из списка sp_cy
                self.m1 = (((float(ws.cell(row=val, column=3).value) - float(cy))**2) + ((float(ws.cell(row=val, column=4).value) - float(cx))** 2) + ((float(ws.cell(row=val, column=5).value) - float(cm))** 2))** 0.5
            #переменной m1 присваиваем расстояние между точками. Координаты 1-й точки - коэффициенты проверяемой строки, 2-й - введённые коэффициенты
                if self.m1 < self.m: # если расстояние между точками меньше хрени для сравнения, то сравнивать будет
                    # теперь с вычисленным расстоянием между точками (m - минимум, с которым сравниваем и ищем меньшее),
                    # иначе ничего не делаем
                    self.m = self.m1
                    self.k1 = val #сохраняем лучший пока что номер строки

            for j in range(0, len(self.sp_cx)): #аналогично проверяем списки с номерами строк близких cx и cm к введённым
                val = self.sp_cx[j]
                self.m1 = (((float(ws.cell(row=val, column=3).value) - float(cy)) ** 2) + ((float(ws.cell(row=val, column=4).value) - float(cx)) ** 2) + ((float(ws.cell(row=val, column=5).value) - float(cm)) ** 2)) ** 0.5
                if self.m1 < self.m:
                    self.m = self.m1
                    self.k1 = val

            for k in range(0, len(self.sp_cm)):
                val = self.sp_cm[k]
                self.m1 = (((float(ws.cell(row=val, column=3).value) - float(cy)) ** 2) + ((float(ws.cell(row=val, column=4).value) - float(cx)) ** 2) + ((float(ws.cell(row=val, column=5).value) - float(cm)) ** 2)) ** 0.5
                if self.m1 < self.m:
                    self.m = self.m1
                    self.k1 = val

            #вывод
            for n in range(1,6): #найденный номер наиближайшей строки хранится в k1, выводим значения из строки k1 и столбцов с коэффициентами
                self.name_profile_print.setText(str((ws.cell(row=self.k1, column=1).value)))
                self.Cy_finall.setText(str((ws.cell(row=self.k1, column=3).value)))
                self.Cx_finall.setText(str((ws.cell(row=self.k1, column=4).value)))
                self.Mz_finall.setText(str((ws.cell(row=self.k1, column=5).value)))                    
            self.m = 9999.9 #возвращаем крупное значение для новых прогонов
            self.k1 = 1 #аналогично, чтобы при следующих прогонах не использовались старые данные
            self.m1 = 0

    #вызов функций
    a = Profiles(cy=float, cx=float, cm=float) #класс профилей, его параметры cy, cx, cm, они вещественные
    '''a.poisk()''' #вызываем функцию поиска полного совпадения
    a.srav(cy=float(input()), cx=float(input()), cm=float(input()), val=int) #вызываем функцию поиска приближённого, говоря, что параметры будем вводить с клавиатуры
        
app = QApplication(sys.argv)
w1 = Start_window()
w1.show()
app.exec()
        

        
