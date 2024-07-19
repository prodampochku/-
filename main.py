from PyQt5 import QtWidgets, uic
import matplotlib
import openpyxl


app = QtWidgets.QApplication([])

ui = uic.loadUi('design.ui')


list_angle = []              #Заполняем углы атаки
for i in range(0, 100, 10):
    list_angle.append(str(i))
    
ui.alpha_Box.addItems(list_angle)

ui.label_9.setStyleSheet('color: rgb(0, 139, 210)')
ui.textEdit.setStyleSheet('color: rgb(0, 139, 210)')
ui.label_10.setStyleSheet('color: rgb(0, 139, 210)')
ui.label_2.setStyleSheet('color: rgb(0, 139, 210)')
ui.label_3.setStyleSheet('color: rgb(0, 139, 210)')
ui.label.setStyleSheet('color: rgb(0, 139, 210)')
ui.label_4.setStyleSheet('color: rgb(0, 139, 210)')
ui.label_5.setStyleSheet('color: rgb(0, 139, 210)')
ui.label_6.setStyleSheet('color: rgb(0, 139, 210)')
ui.label_7.setStyleSheet('color: rgb(0, 139, 210)')
ui.label_8.setStyleSheet('color: rgb(0, 139, 210)')
ui.label_11.setStyleSheet('color: rgb(0, 139, 210)')
ui.label_12.setStyleSheet('color: rgb(0, 139, 210)')
ui.label_15.setStyleSheet('color: rgb(0, 139, 210)')
ui.label_16.setStyleSheet('color: rgb(0, 139, 210)')

#v = 330 #скорость набегающего потока м/с
#p = 1.225 #плотномть воздуха кг/м^3
#def Count_Cy():
 #   Cy = 2*Fy/(S*p*v**2)
    
#def Count_Cx():
 #   Cx = 2*Fx/(S*p*v**2)
    
#def Count_Mz():
 #   Cm = (Mz*S*p*v**2)/2
    
def algoritm():
    # вводим с клавиатуры коэффициенты
    b_0 = ui.finall_horda.text()  #корневая хорда
    b_k = ui.current_horda.text()  #концевая хорда
    l = ui.razmah.text()  #размах крыла
    v = ui.V_wing.text()  #скорость набегающего потока
    Y = ui.Cy_current.text()  #подъёмная сила
    X = ui.Cx_current.text()  #сила лобового сопротивления
    Mz = ui.Mz_current.text()  #продольный момент (тангажа)
    rho = 0.125 #плотность воздуха при н.у.; константа
    def str_to_num(line):  #функция переводит данное ей значение в строку (вместо line даём какую-либо переменную)
        line = line.strip()
        if line.isdigit():  #если значение состоит только из цифр, возвращает целочисленное значение
            return int(line)
        elif '.' in line:  #если встречается точка, возвращает вещественное значение
            return float(line)
        elif ',' in line:  #если встречается запятая, она заменяется на точку, с которой возвращается вещественное значение
            return float(line.replace(',', '.'))

    b_0 = str_to_num(line=b_0)
    b_k = str_to_num(line=b_k)
    l = str_to_num(line=l)
    v = str_to_num(line=v)
    Y = str_to_num(line=Y)
    X = str_to_num(line=X)
    Mz = str_to_num(line=Mz)

    b_mid = (b_0 + b_k)/2  #средняя хорда
    s = 0.5 * (b_0 + b_k) * l  #площадь крыла
    cy = 2*Y/(rho*s*(v**2))  #коэффициент подъёмной силы
    cx = 2*X/(rho*s*(v**2))  #коэффициент лобового сопротивления
    cm = 2*Mz/(rho*s*b_mid*(v**2))  #коэффициент момента тангажа

    delta_cy = 9999.9  #ориентиры-сравнения
    delta_cx = 9999.9
    delta_cm = 9999.9
    sp_cy = []  #списки для строк, подходящих по одному параметру
    sp_cx = []
    sp_cm = []

    val = int  #переменная, которая хранит локальный номер строки
    m = 9999.9 #переменная для сравнения в списках
    m1 = float() #переменная для записи в списках
    k1 = 1  #переменная, которая хранит номер строки с наиближайшими коэффициентами

    wb = openpyxl.load_workbook('data.xlsx')
    ws = wb.active

    def calc(row_ind, column_ind, rate):  #разность введённого коэффициента и проверяемого. в функции используются параметры номера строки, столбца и коэффициента
        return (rate - float(ws.cell(row=row_ind, column=column_ind).value))


    #ищем строк, которые теоритчески могут подойти, сохраняем номер строк из бд, которые всё ближе и ближе к вводимым данным
    for i in range(2, ws.max_row):
            if not ws.cell(row=i, column=5).value:  #если ячейка cm в данной строке пустая, строка пропускается
                continue
            if calc(row_ind=i, column_ind=3, rate=cy) **2 - (delta_cy)**2 <= 0:  # разность целевого Cy и
                    # проверяемого возводим в квадрат и сравниваем с квадратом разности-ориентира
                delta_cy = abs(calc(i, 3, cy))  # если меньше разности-ориентира,
                    # то модуль разности целевого и проверяемого Cy становится новой разностью-ориентиром
                sp_cy.append(i)  # в список sp_cy добавляем номер строки, которая проверялась
            elif calc(row_ind=i, column_ind=3, rate=cy) **2 - (delta_cy)**2 > 0:  # если больше разности-ориентира, то ничего не делает программа
                pass
            if calc(row_ind=i, column_ind=4, rate=cx) **2 - (delta_cx)**2 <= 0:
                delta_cx = abs(calc(i, 4, cx))
                sp_cx.append(i)
            elif calc(row_ind=i, column_ind=4, rate=cx) **2 - (delta_cx)**2 > 0:
                pass
            if calc(row_ind=i, column_ind=5, rate=cm) **2 - (delta_cm)**2 <= 0:
                delta_cm = abs(calc(i, 5, cm))
                sp_cm.append(i)
            elif calc(row_ind=i, column_ind=5, rate=cm) **2 - (delta_cm)**2 > 0:
                pass

    #ищем наиболее близкое геометрически (корень из суммы квадратов)
    for i in range(0, len(sp_cy)): #прогон по списку с Cy
        val = sp_cy[i] #переменной val присваиваем i-е значение из списка sp_cy
        m1 = (((float(ws.cell(row=val, column=3).value) - float(cy))**2) + ((float(ws.cell(row=val, column=4).value) - float(cx))** 2) + ((float(ws.cell(row=val, column=5).value) - float(cm))** 2))** 0.5
    #переменной m1 присваиваем расстояние между точками. Координаты 1-й точки - коэффициенты проверяемой строки, 2-й - введённые коэффициенты
        if m1 < m: # если расстояние между точками меньше хрени для сравнения, то сравнивать будет
                # теперь с вычисленным расстоянием между точками (m - минимум, с которым сравниваем и ищем меньшее),
                # иначе ничего не делаем
            m = m1
            k1 = val #сохраняем лучший пока что номер строки

    for j in range(0, len(sp_cx)): #аналогично проверяем списки с номерами строк близких cx и cm к введённым
        val = sp_cx[j]
        m1 = (((float(ws.cell(row=val, column=3).value) - float(cy)) ** 2) + ((float(ws.cell(row=val, column=4).value) - float(cx)) ** 2) + ((float(ws.cell(row=val, column=5).value) - float(cm)) ** 2)) ** 0.5
        if m1 < m:
            m = m1
            k1 = val

    for k in range(0, len(sp_cm)):
        val = sp_cm[k]
        m1 = (((float(ws.cell(row=val, column=3).value) - float(cy)) ** 2) + ((float(ws.cell(row=val, column=4).value) - float(cx)) ** 2) + ((float(ws.cell(row=val, column=5).value) - float(cm)) ** 2)) ** 0.5
        if m1 < m:
            m = m1
            k1 = val

    

    #вывод
    ui.name_profile_print.setText(str((ws.cell(row=k1, column=1).value)))#print(ws.cell(row=k1, column=1).value)  # профиль
    #print(ws.cell(row=k1, column=2).value)  # угол атаки
    ui.Cy_finall.setText(str((ws.cell(row=k1, column=3).value)))# Cy
    ui.Cx_finall.setText(str((ws.cell(row=k1, column=4).value)))# Cx
    ui.Mz_finall.setText(str((ws.cell(row=k1, column=5).value)))# Cm
    pixmap = QPixmap('Профили/' + (ws.cell(row=k1, column=6).value))
    ui.label_14.setPixmap(pixmap)
    delta_cy = 9999.9  #возвращаем исходные значения для ориенторив
    delta_cx = 9999.9
    delta_cm = 9999.9
    m = 9999.9
    k1 = 1
    sp_cy = sp_cy.clear()  #списки очищаем
    sp_cx = sp_cx.clear()
    sp_cm = sp_cm.clear()
    #Fy = float(ui.Cy_current.text())
    #Fx = float(ui.Cx_current.text())
    #Mz = float(ui.Mz_current.text())
    #S = float(ui.S_wing.text())
    #Ba = float(ui.horda.text())
    #cy = Count_Cy
    #cx = Count_Cx
    #cm = Count_Mz
    #ui.Cy_current.setText('0.676')
    #ui.Cx_current.setText('0.0470')
    #ui.Mz_current.setText('0.155')
    #cy = float(ui.Cy_current.text())
    #cx = float(ui.Cx_current.text())
    #cm = float(ui.Mz_current.text())


    #a = int  # переменная, в которую будет перезаписываться наиближайшая строчка
    #srav = 9999.9  # разность-ориентир, с которой сравниваем. очень большая, чтобы потом при сравнении с любой меньшей разностью последнюю сделать ориентиром
    #k1 = []  # список для строк, в которых 3 параметра лучше введённых
    #k2 = []  # список для строк, в которых любые 2 параметра лучше введённых
    #k3 = []  # список для строк, в которых только 1 параметр лучше введённого
    #k4 = []  # список для строк, полностью совпадающей с введёнными данными

    #wb = openpyxl.load_workbook('data.xlsx')  # доступ к эксель странице
    #ws = wb.active

    #for i in range(2,20):  # пробегаем строчки
        #dy = ((float(ws.cell(row=i, column=3).value) - cy)/cy)*100  # относительное изменение Cy из таблицы относительно введённого в процентах. Положительное значение - коэффициент лучше введённого (больше)
        #dx = ((cx - float(ws.cell(row=i, column=4).value))/cx)*100  # относительное изменение Cx из таблицы относительно введённого в процентах. Положительное значение - коэффициент лучше введённого (меньше)
        #dm = ((cm - float(ws.cell(row=i, column=5).value))/cm)*100  # относительное изменение Cm из таблицы относительно введённого в процентах. Положительное значение - коэффициент лучше введённого (меньше)
        #d = (dy**2 + dx**2 + dm**2)**0.5  # геометрическое расстояние между проверяемыми и введёнными данными
        #d1 = 0  # ориентир для сравнения в списке k1. Сравниваем с малым в поисках бóльшего значения
        #d3 = 9999.9  # ориентир для сравнения в списке k3. Сравниваем с большим в поисках меньшего
        #dy = round(dy, 3)  # для наглядности при отработке (потом можно будет вырезать вместе с процентами, когда закончится отладки кода)
        #dx = round(dx, 3)
        #dm = round(dm, 3)
        #print(i, "dy = ", dy, "dx = ", dx, "dm = ", dm, "d = ", d)  # тоже для проверки

        #if dy > 0 and dx > 0 and dm > 0:  # список k1, если строка подходит по значениям, номер строки записывается в список
         #   k1.append(i)
          #  if d > d1:  # для 1-го списка можно находить наибольшее расстояние между точками, для 3-го - наименьшее
           #     d1 = d  # чтобы среди списка найти наилучшие показатели
               # kd1 = i  # номер строки с более хорошими коэффициентами в 1-м списке
        #if (dy > 0 and dx > 0 and dm < 0) or (dy > 0 and dx < 0 and dm > 0) or (dy < 0 and dx > 0 and dm > 0):  # до 39 строки кода аналогично
         #   k2.append(i)
       # if (dy > 0 and dx < 0 and dm < 0) or (dy < 0 and dx < 0 and dm > 0) or (dy < 0 and dx > 0 and dm < 0):
           # k3.append(i)
           # if d < d3:
             #   d3 = d
             #   kd3 = i
       # if (dy == 0 and dx == 0 and dm == 0):  # если находится полное совпадение, номер строки в отдельный список
           # k4.append(i)

        #print(k4,k1, k2, k3)  # для отладки
        #print("------NEXT------")

    # если 4-й список не пустой, присваиваем переменной а значение нужной строки
    #if k4:
        #a = k4[len(k4)-1]
    # если полного совпадения нет, просматриваем аналогично список с всеми коэффициентами выше введённых
   # else:
       # if k1:
            #a = kd1  # строка с наивысшими коэффициентами по отношению к введённым
    # если 1-й список пуст, проверяем с двумя коэффициентами выше введённых
       # else:
            #if k2:
                #a = k2[len(k2) - 1]  # пока используется последняя записанная в список строка. Нужно продумать алгоритм!
    # аналогия с первым списком
            #else:
                #if k3:
                   # a =kd3

    #ui.name_profile_print.setText(str((ws.cell(row=a, column=1).value))) #профиль
    #print(ws.cell(row=number, column=2).value) #угол атаки
    #ui.Cy_finall.setText(str((ws.cell(row=a, column=3).value)))# Cy
    #ui.Cx_finall.setText(str((ws.cell(row=a, column=4).value)))# Cx
    #ui.Mz_finall.setText(str((ws.cell(row=a, column=5).value)))# Cm
    #pixmap = QPixmap('Профили/' + str((ws.cell(row=a, column=6).value)))
    #ui.label_14.setPixmap(pixmap) #картинка
   
ui.btn1.clicked.connect(algoritm)
ui.Tab_widget.setCurrentIndex(0)
def tab_switch1():
    ui.Tab_widget.setCurrentIndex(1)
ui.start_btn.clicked.connect(tab_switch1)
def tab_switch2():
    ui.Tab_widget.setCurrentIndex(2)
ui.btn1.clicked.connect(tab_switch2)
ui.show()
app.exec()
