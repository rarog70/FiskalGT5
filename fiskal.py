import sys
import os
import re
import openpyxl
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtGui import * 
from PyQt5.QtWidgets import *
import fiskal_gui

list = []
class Fiskal(QMainWindow, fiskal_gui.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.lineEdit.returnPressed.connect(self.AddLine)  # При нажатии "Энтер" добавляем строку
        self.pushButton_2.clicked.connect(QCoreApplication.instance().quit) # Выход из программы
        self.pushButton.clicked.connect(self.Export)  # Выполняем экспорт

    def center(self):  # центрируем форму на экране
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def AddLine(self):  # Добавляем новую строку в лист-виджет
        ser = self.Serial(self.lineEdit.text())
        if ser not in list:  # Проверяем уникальность введенных номеров
            list.append(self.Serial(self.lineEdit.text()))
            self.listWidget.addItem(ser)
        else:
            QMessageBox.about(self, "ВНИМАНИЕ", "Такой номер уже занесен в таблицу")
        self.lineEdit.clear()
        self.statusBar.showMessage(f"Всего записей: {str(self.listWidget.count())}")

    def Serial(self, sn):  # Формируем строку для добавления в список
        sn = re.sub(r"[жЖ]", ";", sn)  # Исправляем раскладку
        sn = sn.split(";")  # Разделяем поля по разделителю
        return sn[0]  # Возвращаем серийный номер
    
    def Export(self):  # формируем АПП_ФН.xls
        resp = QMessageBox.question(self, "Экспорт", "Экспортировать записи?")
        if resp == QMessageBox.Yes:
            try:
                wb = openpyxl.load_workbook('template.xlsx') # Открываем шаблон
            except FileNotFoundError:
                print(f"Шаблон template.xlsx не существует.\nСоздайте шаблон и повторите снова.")
            sheet = wb.get_sheet_by_name('Лист1')  # Активируем нужный лист
            s = 1
            for i in range(self.listWidget.count()):  # Переносим табличные данные
                sheet['A' + str(i+10)].value = s
                sheet['B' + str(i+10)].value = "Фискальный накопитель"
                sheet['C' + str(i+10)].value = self.listWidget.item(i).text()
                sheet['D' + str(i+10)].value = "1"
                s += 1
            wb.save("АПП_ФН.xlsx")  # Сохраняем изменения в новом файле
            self.statusBar.showMessage("Экспорт успешно завершен")
            self.listWidget.clear()  # Очищаем лист виджет
        elif resp == QMessageBox.No:
            self.statusBar.showMessage("Экспорт отменен")        
        resp = QMessageBox.question(self, "АПП", "Открыть файл АПП??")
        if resp == QMessageBox.Yes:  # Открываем полученный файл в зависимости от системы
            if os.name == "posix":
                os.system("libreoffice --calc АПП_ФН.xlsx")
            elif os.name == "nt":
                os.system("АПП_ФН.xlsx")
        elif resp == QMessageBox.No:
            pass

    def openMenu(self, position): # Формируем контекстное меню
        menu = QtWidgets.QMenu()
        addDes = QtWidgets.QAction('Удалить', menu)
        addDes.triggered.connect(self.del_current)
        menu.addAction(addDes)
        menu.exec_(self.listWidget.viewport().mapToGlobal(position))

    def del_current(self):
        print("Типа что то удалил")





if __name__ == '__main__':
    app = QApplication(sys.argv)
    form = Fiskal()
    form.center()
    form.show()
    app.exec()
