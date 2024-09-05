import os
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QFileDialog, QMessageBox
from os import getcwd
import pandas as pd
import numpy as np
import re
import main
import threading
import time




class Ui_Dialog(object):
    def __init__(self):
        self.path_1 = None
        self.path_2 = None
        self.res = None
        self.collumn_file_2_lst = None
        self.collumn_file_1_lst = None
        self.data_file_2 = None
        self.data_file_1 = None
        self.df_total = None

    def setupUi(self, Dialog):
        Dialog.setObjectName("Иваныч")
        Dialog.resize(900, 800)

        self.tabWidget = QtWidgets.QTabWidget(parent=Dialog)
        self.tabWidget.setGeometry(QtCore.QRect(10, 10, 891, 791))
        self.tabWidget.setAutoFillBackground(True)
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.label = QtWidgets.QLabel(parent=self.tab)
        self.label.setGeometry(QtCore.QRect(350, 20, 181, 25))
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_3 = QtWidgets.QLabel(parent=self.tab)
        self.label_3.setGeometry(QtCore.QRect(650, 105, 300, 41))
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(parent=self.tab)
        self.label_4.setGeometry(QtCore.QRect(22, 300, 251, 41))
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.lineEdit = QtWidgets.QLineEdit(parent=self.tab)
        self.lineEdit.setGeometry(QtCore.QRect(20, 367, 50, 20))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit.setText('0')
        self.lineEdit.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.label_5 = QtWidgets.QLabel(parent=self.tab)
        self.label_5.setGeometry(QtCore.QRect(80, 367, 300, 16))
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.lineEdit_2 = QtWidgets.QLineEdit(parent=self.tab)
        self.lineEdit_2.setGeometry(QtCore.QRect(20, 400, 50, 20))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.lineEdit_2.setText('0')
        self.lineEdit_2.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.label_6 = QtWidgets.QLabel(parent=self.tab)
        self.label_6.setGeometry(QtCore.QRect(80, 400, 241, 20))
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.progressBar = QtWidgets.QProgressBar(parent=self.tab)
        self.progressBar.setGeometry(QtCore.QRect(300, 600, 391, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")

        self.pushButton_4 = QtWidgets.QPushButton(parent=self.tab)
        self.pushButton_4.setGeometry(QtCore.QRect(200, 595, 91, 31))
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_4.clicked.connect(self.start)

        self.label_7 = QtWidgets.QLabel(parent=self.tab)
        self.label_7.setGeometry(QtCore.QRect(550, 310, 171, 16))
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")

        self.label_8 = QtWidgets.QLabel(parent=self.tab)
        self.label_8.setGeometry(QtCore.QRect(350, 650, 281, 41))

        self.label_1 = QtWidgets.QLabel(parent=self.tab)
        self.label_1.setGeometry(QtCore.QRect(300, 560, 391, 41))
        self.label_1.setAlignment(Qt.AlignmentFlag.AlignCenter)


        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_8.setFont(font)
        self.label_8.setObjectName("label_8")

        self.label_1.setFont(font)
        self.label_1.setObjectName("label_1")

        self.pushButton_5 = QtWidgets.QPushButton(parent=self.tab)
        self.pushButton_5.setGeometry(QtCore.QRect(380, 700, 150, 31))
        self.pushButton_5.setObjectName("pushButton_5")


        self.pushButton_5.setEnabled(False)
        self.pushButton_5.clicked.connect(self.save_file)

        self.gridLayoutWidget = QtWidgets.QWidget(parent=self.tab)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(550, 360, 201, 132))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setSpacing(0)
        self.gridLayout.setObjectName("gridLayout")
        self.label_10 = QtWidgets.QLabel(parent=self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_10.setFont(font)
        self.label_10.setObjectName("label_10")
        self.gridLayout.addWidget(self.label_10, 4, 0, 1, 1)
        self.label_9 = QtWidgets.QLabel(parent=self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_9.setFont(font)
        self.label_9.setObjectName("label_9")
        self.gridLayout.addWidget(self.label_9, 0, 0, 1, 1)
        self.label_14 = QtWidgets.QLabel(parent=self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_14.setFont(font)
        self.label_14.setObjectName("label_14")
        self.gridLayout.addWidget(self.label_14, 3, 0, 1, 1)
        # self.label_12 = QtWidgets.QLabel(parent=self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(11)
        # self.label_12.setFont(font)
        # self.label_12.setObjectName("label_12")
        # self.gridLayout.addWidget(self.label_12, 5, 0, 1, 1)
        self.label_11 = QtWidgets.QLabel(parent=self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_11.setFont(font)
        self.label_11.setObjectName("label_11")
        self.gridLayout.addWidget(self.label_11, 2, 0, 1, 1)
        self.comboBox_3 = QtWidgets.QComboBox(parent=self.tab)
        self.comboBox_3.setGeometry(QtCore.QRect(540, 160, 101, 22))
        self.comboBox_3.setObjectName("comboBox_3")
        self.comboBox_3.setEnabled(False)
        self.label_20 = QtWidgets.QLabel(parent=self.tab)
        self.label_20.setGeometry(QtCore.QRect(650, 210, 161, 16))
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_20.setFont(font)
        self.label_20.setObjectName("label_20")
        self.comboBox_4 = QtWidgets.QComboBox(parent=self.tab)
        self.comboBox_4.setGeometry(QtCore.QRect(540, 210, 101, 22))
        self.comboBox_4.setObjectName("comboBox_4")
        self.comboBox_4.setEnabled(False)
        self.label_21 = QtWidgets.QLabel(parent=self.tab)
        self.label_21.setGeometry(QtCore.QRect(650, 160, 121, 16))
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_21.setFont(font)
        self.label_21.setObjectName("label_21")
        self.label_22 = QtWidgets.QLabel(parent=self.tab)
        self.label_22.setGeometry(QtCore.QRect(20, 550, 130, 110))
        self.label_22.setText("")
        self.label_22.setPixmap(QtGui.QPixmap("free-icon-miner-307894.png"))
        self.label_22.setScaledContents(True)
        self.label_22.setObjectName("label_22")
        self.label_45 = QtWidgets.QLabel(parent=self.tab)
        self.label_45.setGeometry(QtCore.QRect(130, 210, 161, 16))
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_45.setFont(font)
        self.label_45.setObjectName("label_45")
        self.pushButton_8 = QtWidgets.QPushButton(parent=self.tab)
        self.pushButton_8.setGeometry(QtCore.QRect(20, 110, 101, 31))
        self.pushButton_8.setObjectName("pushButton_8")
        self.res = self.pushButton_8.clicked.connect(self.open_file_1)




        self.pushButton_1 = QtWidgets.QPushButton(parent=self.tab)
        self.pushButton_1.setGeometry(QtCore.QRect(540, 110, 101, 31))
        self.pushButton_1.setObjectName("pushButton_1")
        self.pushButton_1.clicked.connect(self.open_file_2)

        self.label_18 = QtWidgets.QLabel(parent=self.tab)
        self.label_18.setGeometry(QtCore.QRect(130, 105, 300, 41))
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_18.setFont(font)
        self.label_18.setObjectName("label_18")
        self.comboBox_9 = QtWidgets.QComboBox(parent=self.tab)
        self.comboBox_9.setGeometry(QtCore.QRect(20, 210, 101, 22))
        self.comboBox_9.setObjectName("comboBox_9")
        self.comboBox_9.setEnabled(False)
        self.label_46 = QtWidgets.QLabel(parent=self.tab)
        self.label_46.setGeometry(QtCore.QRect(130, 160, 121, 16))
        font = QtGui.QFont()
        font.setPointSize(11)
        self.label_46.setFont(font)
        self.label_46.setObjectName("label_46")
        self.comboBox_10 = QtWidgets.QComboBox(parent=self.tab)
        self.comboBox_10.setGeometry(QtCore.QRect(20, 160, 101, 22))
        self.comboBox_10.setObjectName("comboBox_10")
        self.comboBox_10.setEnabled(False)
        self.tabWidget.addTab(self.tab, "")
        # self.tab_2 = QtWidgets.QWidget()
        # self.tab_2.setObjectName("tab_2")
        # self.tab_2.setEnabled(False)
        # self.pushButton_6 = QtWidgets.QPushButton(parent=self.tab_2)
        # self.pushButton_6.setGeometry(QtCore.QRect(40, 50, 91, 31))
        # self.pushButton_6.setObjectName("pushButton_6")
        # self.label_13 = QtWidgets.QLabel(parent=self.tab_2)
        # self.label_13.setGeometry(QtCore.QRect(150, 54, 121, 20))
        font = QtGui.QFont()
        font.setPointSize(11)
        # self.label_13.setFont(font)
        # self.label_13.setObjectName("label_13")
        # self.progressBar_2 = QtWidgets.QProgressBar(parent=self.tab_2)
        # self.progressBar_2.setGeometry(QtCore.QRect(660, 50, 171, 23))
        # self.progressBar_2.setProperty("value", 24)
        # self.progressBar_2.setObjectName("progressBar_2")
        # self.label_15 = QtWidgets.QLabel(parent=self.tab_2)
        # self.label_15.setGeometry(QtCore.QRect(470, 53, 161, 20))
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        # self.label_15.setFont(font)
        # self.label_15.setObjectName("label_15")
        # self.label_16 = QtWidgets.QLabel(parent=self.tab_2)
        # self.label_16.setGeometry(QtCore.QRect(20, 120, 211, 16))
        font = QtGui.QFont()
        font.setPointSize(11)
        # self.label_16.setFont(font)
        # self.label_16.setObjectName("label_16")
        # self.lineEdit_3 = QtWidgets.QLineEdit(parent=self.tab_2)
        # self.lineEdit_3.setGeometry(QtCore.QRect(38, 170, 111, 16))
        # self.lineEdit_3.setObjectName("lineEdit_3")
        # self.tableWidget = QtWidgets.QTableWidget(parent=self.tab_2)
        # self.tableWidget.setGeometry(QtCore.QRect(80, 320, 711, 311))
        # self.tableWidget.setObjectName("tableWidget")
        # self.tableWidget.setColumnCount(5)
        # self.tableWidget.setRowCount(13)
        item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setVerticalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setVerticalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setVerticalHeaderItem(4, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setVerticalHeaderItem(5, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setVerticalHeaderItem(6, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setVerticalHeaderItem(7, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setVerticalHeaderItem(8, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setVerticalHeaderItem(9, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setVerticalHeaderItem(10, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setVerticalHeaderItem(11, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setVerticalHeaderItem(12, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setHorizontalHeaderItem(0, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setHorizontalHeaderItem(1, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setHorizontalHeaderItem(2, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setHorizontalHeaderItem(3, item)
        # item = QtWidgets.QTableWidgetItem()
        # self.tableWidget.setHorizontalHeaderItem(4, item)
        # self.tableWidget.horizontalHeader().setDefaultSectionSize(140)
        # self.tableWidget.horizontalHeader().setMinimumSectionSize(24)
        # self.textBrowser = QtWidgets.QTextBrowser(parent=self.tab_2)
        # self.textBrowser.setGeometry(QtCore.QRect(220, 140, 491, 101))
        # self.textBrowser.setObjectName("textBrowser")
        # self.pushButton_3 = QtWidgets.QPushButton(parent=self.tab_2)
        # self.pushButton_3.setGeometry(QtCore.QRect(30, 220, 111, 31))
        # self.pushButton_3.setObjectName("pushButton_3")
        # self.label_17 = QtWidgets.QLabel(parent=self.tab_2)
        # self.label_17.setGeometry(QtCore.QRect(340, 640, 251, 41))
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        # self.label_17.setFont(font)
        # self.label_17.setObjectName("label_17")
        # self.pushButton_15 = QtWidgets.QPushButton(parent=self.tab_2)
        # self.pushButton_15.setGeometry(QtCore.QRect(390, 700, 101, 31))
        # self.pushButton_15.setObjectName("pushButton_15")
        # self.tabWidget.addTab(self.tab_2, "")
        self.retranslateUi(Dialog)
        self.tabWidget.setCurrentIndex(1)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Ivanich&Co"))
        self.label.setText(_translate("Dialog", "Ввод данных"))
        self.label_3.setText(_translate("Dialog", "Файл 2 (не задан)"))
        self.label_4.setText(_translate("Dialog", "Установка пороговых значений"))
        self.label_5.setText(_translate("Dialog", "Коэф. векторного расстояния"))
        self.label_6.setText(_translate("Dialog", "Коэф. сравнения Жаккарда"))
        self.pushButton_4.setText(_translate("Dialog", "Старт"))
        self.label_7.setText(_translate("Dialog", "Результаты рассчетов"))
        self.label_8.setText(_translate("Dialog", "Сохранить отчет после рассчета \n       в итоговый файл xlsx"))

        self.label_1.setText(_translate("Dialog", ""))

        self.pushButton_5.setText(_translate("Dialog", "Сохранить данные"))
        self.label_10.setText(_translate("Dialog", "Не вошло:                   0"))
        self.label_9.setText(_translate("Dialog", "Всего данных файл 1: 0"))
        self.label_14.setText(_translate("Dialog", "Построенно связей:   0"))
        # self.label_12.setText(_translate("Dialog", "Точность (%): 0"))
        self.label_11.setText(_translate("Dialog", "Всего данных файл 2: 0"))
        self.label_20.setText(_translate("Dialog", "Столбец с данными"))
        self.label_21.setText(_translate("Dialog", "Столбец id"))
        self.label_45.setText(_translate("Dialog", "Столбец с данными"))
        self.pushButton_8.setText(_translate("Dialog", "Открыть"))
        self.pushButton_1.setText(_translate("Dialog", "Открыть"))
        self.label_18.setText(_translate("Dialog", "Файл 1 (не задан)"))
        self.label_46.setText(_translate("Dialog", "Столбец id"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("Dialog", "Рассчет связей"))
        # self.pushButton_6.setText(_translate("Dialog", "Расчет весов"))
        # self.label_13.setText(_translate("Dialog", "для 0 пар"))
        # self.label_15.setText(_translate("Dialog", "<html><head/><body><p><span style=\" font-weight:400;\">Прогресс рассчета</span></p></body></html>"))
        # self.label_16.setText(_translate("Dialog", "Введите № id мероприятия"))
        # item = self.tableWidget.horizontalHeaderItem(0)
        # item.setText(_translate("Dialog", "id"))
        # item = self.tableWidget.horizontalHeaderItem(1)
        # item.setText(_translate("Dialog", "Текст"))
        # item = self.tableWidget.horizontalHeaderItem(2)
        # item.setText(_translate("Dialog", "Вектор"))
        # item = self.tableWidget.horizontalHeaderItem(3)
        # item.setText(_translate("Dialog", "Жаккард"))
        # item = self.tableWidget.horizontalHeaderItem(4)
        # item.setText(_translate("Dialog", "Вес"))
        # self.pushButton_3.setText(_translate("Dialog", "Показать связи"))
        # self.label_17.setText(_translate("Dialog", "Сохранить итоговый файл в xlsx"))
        # self.pushButton_15.setText(_translate("Dialog", "ОК"))
        # self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("Dialog", "Рассчет весов"))

    def open_file_1(self):
        res = QFileDialog.getOpenFileName(Dialog, 'Open file', f'{getcwd()}', 'XLSX Files (*.xlsx)')
        if res[0] != '':
            self.comboBox_9.clear(), self.comboBox_10.clear()
            self.label_18.setText(f"Файл 1 ({res[0].split('/')[len(res[0].split('/')) - 1]})")
            self.data_file_1 = pd.read_excel(res[0])
            if self.data_file_1.shape[0] == 0:
               self.Error('Файл 1 не имеет данных')
            else:
                self.label_9.setText(f"Всего данных файл 1: {self.data_file_1.shape[0]}")
                self.collumn_file_1_lst = list(self.data_file_1.columns)
                self.comboBox_9.setEnabled(True)
                self.comboBox_9.addItems(list(self.data_file_1.columns))
                self.comboBox_10.setEnabled(True)
                self.comboBox_10.addItems(list(self.data_file_1.columns))
                self.path_1 = res[0]
        else:
            pass
        return self.path_1

    def open_file_2(self):
        res = QFileDialog.getOpenFileName(Dialog, 'Open file', f'{getcwd()}', 'XLSX Files (*.xlsx)')
        if res[0] != '':
            self.comboBox_3.clear(), self.comboBox_4.clear()
            self.label_3.setText(f"Файл 2 ({res[0].split('/')[len(res[0].split('/')) - 1]})")
            self.data_file_2 = pd.read_excel(res[0])
            if self.data_file_2.shape[0] == 0:
                self.Error('Файл 2 не имеет данных')
            else:
                self.label_11.setText(f"Всего данных файл 1: {self.data_file_2.shape[0]}")
                self.comboBox_3.setEnabled(True)
                self.comboBox_3.addItems(list(self.data_file_2.columns))
                self.comboBox_4.setEnabled(True)
                self.comboBox_4.addItems(list(self.data_file_2.columns))
                self.path_2 = res[0]
        else:
            pass
        return self.path_2

    def progr_bar(self):
        i = 0
        while self.stop_treading:
            time.sleep(1)
            self.progressBar.setValue(int(self.coin))
            i += 1
        print(i)
        time.sleep(1000)



    def start(self):
        self.stop_treading = True
        self.coin = 0
        self.th = threading.Thread(target=self.progr_bar, daemon=True)
        self.th.start()
        if isinstance(self.path_1, str) and isinstance(self.path_2, str):
            if str(self.comboBox_10.currentText()) == str(self.comboBox_9.currentText()):
                self.Error('Совпадают указанные id и данные файла 1')
            elif str(self.comboBox_4.currentText()) == str(self.comboBox_3.currentText()):
                    self.Error('Совпадают указанные id и данные файла 2')
            else:
                self.label_1.setText('Создаем токены файл 1')
                self.df_file_1 = pd.read_excel(self.path_1).loc[:, [self.comboBox_10.currentText(), self.comboBox_9.currentText()]]
                if self.df_file_1[self.comboBox_10.currentText()].value_counts().sum() != self.df_file_1.shape[0]:
                    self.Error('В файле 1 не хватает id')
                elif self.df_file_1[self.comboBox_9.currentText()].value_counts().sum() != self.df_file_1.shape[0]:
                    self.Error('В файле 1 не хватает данных')
                else:
                    self.df_file_1["Токены"] = pd.Series(main.morf(main.open_without_sub_lower(self.path_1)))
                    self.df_file_1.to_excel('слова приведенные к нормальному виду из файла 1.xlsx', index=False)

                    self.label_1.setText('Создаем токены файл 2')
                    self.df_file_2 = pd.read_excel(self.path_2).loc[:, [self.comboBox_3.currentText(), self.comboBox_4.currentText()]]
                    if self.df_file_2[self.comboBox_3.currentText()].value_counts().sum() != self.df_file_2.shape[0]:
                        self.Error('В файле 2 не хватает id')
                    elif self.df_file_2[self.comboBox_4.currentText()].value_counts().sum() != self.df_file_2.shape[0]:
                        self.Error('В файле 2 не хватает данных')
                    else:
                        self.df_file_2["Токены"] = pd.Series(main.morf(main.open_without_sub_lower(self.path_2)))
                        self.df_file_2.to_excel('слова приведенные к нормальному виду из файла 2.xlsx', index=False)
                        self.over_link = self.data_file_1.shape[0] * self.data_file_2.shape[0]
                        self.to_vector()
        else:
            self.Error('Не выбраны файлы')


    def to_vector(self):

        df = np.array(pd.read_excel('слова приведенные к нормальному виду из файла 1.xlsx')['Токены'])
        lst_file_1 = [re.sub('\W+', ' ', text.lower()) for text in (df)]

        """Открытие ранее подготовленного файла с токенами результатов и приведение к нижнему регистру"""

        df = np.array(pd.read_excel('слова приведенные к нормальному виду из файла 2.xlsx')['Токены'])
        lst_file_2 = [re.sub('\W+', ' ', text.lower()) for text in (df)]

        """Перевод текста показателей в векторное представление"""


        pb_100 = 100 / len(lst_file_1)
        self.label_1.setText('Создаем вектора слов файл 1')
        self.coin = 0
        data_to_vector_file_1 = []
        for i in range(len(lst_file_1)):
            predlojenie_to_vector_file_1 = []
            for word in lst_file_1[i].split():
                predlojenie_to_vector_file_1.append(main.text_to_vector(word))
                self.coin += pb_100

            data_to_vector_file_1.append(predlojenie_to_vector_file_1)
        vector_word_file_1 = data_to_vector_file_1

        pb_100 = 100 / len(lst_file_2)
        self.label_1.setText('Создаем вектора слов файл 2')
        self.coin = 0
        data_to_vector_file_2 = []
        for i in range(len(lst_file_2)):
            predlojenie_to_vector_file_2 = []
            for word in lst_file_2[i].split():
                predlojenie_to_vector_file_2.append(main.text_to_vector(word))
                self.coin += pb_100

            data_to_vector_file_2.append(predlojenie_to_vector_file_2)
        vector_word_file_2 = data_to_vector_file_2


        """Последовательная передача показателей и результатов на векторный анализ"""

        self.label_1.setText('Проводим векторный анализ')
        pb_100 = 100 / (len(vector_word_file_1) * len(vector_word_file_1))
        self.coin = 0
        coin_pokaz = 0
        total_lst = []
        for i in range(len(vector_word_file_1)):
            coin_pokaz += 1
            coin_rez = 0
            for j in range(len(vector_word_file_2)):
                coin_rez += 1
                total_lst.append(main.vector_analiz(vector_word_file_1[i], vector_word_file_2[j],
                                                    coin_pokaz, coin_rez, self.lineEdit.text().replace(',', '.')))
                self.coin += pb_100


        """Очистка списка данных векторного анализа от пустых значений"""

        self.label_1.setText('Готовим данные векторного анализа')
        pb_100 = 100 / len(total_lst)
        self.coin = 0
        self.total = []
        if total_lst[0] == None or total_lst[0] == None:
            self.Error('При данном коэффициенте вектора связей нет')

        else:
            for n in range(len(total_lst)):
                if isinstance(total_lst[n], list):
                    self.total.append(total_lst[n])
                    self.coin += pb_100

            self.otchet_vec()


        """Открытие изначальных файлов для получения данных (название показателей и результатов) в интересах формирования отчета
        векторного анализа """

    def otchet_vec(self):

        pb_100 = 100 / len(self.total)
        self.coin = 0
        df_id_file_1 = pd.read_excel(self.path_1)
        df_id_file_2 = pd.read_excel(self.path_2)

        dct = {}
        for t in self.total:
            a = np.array2string(np.array(df_id_file_1.iloc[t[1] - 1, :2]))
            b = np.array2string(np.array(df_id_file_2.iloc[t[3] - 1, :2]))
            c = t[4]
            if dct.get(a, False) is False:
                dct[a] = [str(b)]
                dct[a].append(c)
            else:
                dct[a].append(b)
                dct[a].append(c)
            self.coin += pb_100


        """Формирование DataFrame и запись итогов векторного анализа"""
        df = pd.DataFrame.from_dict(dct, orient='index').T
        df.to_excel('GASU_VECTOR_hudlit.xlsx', index=False)
        self.koef_jakkard()



    def koef_jakkard(self):

        df_rez = pd.read_excel('слова приведенные к нормальному виду из файла 2.xlsx')
        data = pd.read_excel('GASU_VECTOR_hudlit.xlsx')

        """Подготовка данных для сравнения методом Жаккара. Подготовка выходных данных для записи в итоговый DataFrame"""

        self.label_1.setText('Рассчет коэффециента Жаккарда')
        pb_100 = 50 / (data.shape[0] * len(data.columns))
        tttotal = []
        self.coin = 0
        for i in range(len(data.columns)):
            ttotal = []
            cur_col = data.columns[i]
            pokazatel = [(re.sub('\W+', ' ', text)) for text in data.columns[i]]
            pokazatel = ''.join(pokazatel).split()
            id_pok = ''.join(pokazatel[0])
            pokazatel_name = ' '.join(pokazatel[1:])
            pok_to_jakar = main.pok_id_val(data.columns[i])
            list1 = pok_to_jakar[1]
            cur_data = data[cur_col].dropna()

            rez_to_jakar = main.rez_vec_id_val(cur_data)
            for j in range(len(rez_to_jakar[1])):
                total = []
                list2 = rez_to_jakar[1][j]
                return_from_jacar = main.jaccard(list1, list2)
                ID = [rez_to_jakar[0][j][0]]

                index = df_rez.index[df_rez[self.comboBox_3.currentText()] == int(ID[0])]
                name = df_rez.iloc[index][self.comboBox_4.currentText()]
                return_from_jacar = [round(return_from_jacar[0], 3)]
                self.coin += pb_100

                if return_from_jacar[0] > float(self.lineEdit_2.text().replace(',', '.')):
                    total.append(id_pok)
                    total.append(pokazatel_name)
                    total.append(rez_to_jakar[0][j][0])
                    total.append(list(name)[0])
                    total.append(rez_to_jakar[2][j][0])
                    total.append(return_from_jacar[0])
                    ttotal.append(total)
            tttotal.extend(ttotal)

        """Запись итоговых результатов в DataFrame"""


        self.df_total = pd.DataFrame(tttotal)
        if self.df_total.empty:
            self.Error('При данном коэффициенте Жаккарда связей не обнаружено')
        else:
            self.df_total.columns = ['id файл_1', 'Данные файл_1', 'id файл_2', 'Данные файл_2', 'Вектор', 'Жаккард']

            self.label_14.setText(f"Построенно связей:   {self.df_total.shape[0]}")
            self.total_link = self.df_total.shape[0]
            self.label_10.setText(f"Не вошло:                   {self.over_link - self.total_link}")

            self.progressBar.setValue(100)
            self.label_1.setText('Завершено')
            self.pushButton_5.setEnabled(True)
        self.del_file()


        """Удаление временных файлов"""

    def del_file(self):

        if os.path.isfile(os.path.abspath('слова приведенные к нормальному виду из файла 1.xlsx')):
            os.remove(os.path.abspath('слова приведенные к нормальному виду из файла 1.xlsx'))
        if os.path.isfile(os.path.abspath('слова приведенные к нормальному виду из файла 2.xlsx')):
            os.remove(os.path.abspath('слова приведенные к нормальному виду из файла 2.xlsx'))
        if os.path.isfile(os.path.abspath('GASU_VECTOR_hudlit.xlsx')):
            os.remove(os.path.abspath('GASU_VECTOR_hudlit.xlsx'))




        """Выгрузка данных в итоговый файл"""

    def save_file(self):
        save_file_name = QFileDialog.getSaveFileName(None, "Save File", ".", "XSLX Files (*.xlsx)")

        if save_file_name[0] != '' and isinstance(self.df_total, pd.DataFrame):
            with pd.ExcelWriter(str(save_file_name[0])) as writer:
                self.df_total.to_excel(writer, index=False)
        else:
            self.Error('Нет данных')






    def Error(self, massage):
        self.label_14.setText(f"Построенно связей:   0")
        self.label_10.setText(f"Не вошло:                   0")
        self.label_1.setText(''), self.progressBar.setValue(0)
        msg = QMessageBox()
        msg.setInformativeText(massage)
        msg.setWindowTitle("Ошибка")
        msg.exec()




if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec())
