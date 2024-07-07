from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow

import sys

class Window(QMainWindow):
    def __init__(self):
        super(Window, self).__init__()

        self.setWindowTitle('Подбор профиля крыла')
        self.setGeometry(300, 250, 500, 500)

        self.new_text = QtWidgets.QLabel(self)

        self.main_text = QtWidgets.QLabel(self)
        self.main_text.setText('Подбор профиля крыла')
        self.main_text.move(180, 100)
        self.main_text.adjustSize()

        self.btn = QtWidgets.QPushButton(self)
        self.btn.move(150, 150)
        self.btn.setText('Найти профиль')
        self.btn.setFixedWidth(200)
        self.btn.clicked.connect(self.add_label)

        

    def add_label(self):
        self.new_text.setText('Геометрические параметры профиля:')
        self.new_text.move(150, 200)
        self.new_text.adjustSize()

    
def application():
    app = QApplication(sys.argv)
    window = Window()

    window.show()
    sys.exit(app.exec_())
print(application())
