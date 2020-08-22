#!/usr/bin/env python3

import sys
from PyQt5 import QtWidgets
from gui import design
from modules import excel_file
import pprint


class CpromApp(QtWidgets.QMainWindow, design.Ui_MainWindow):
    def __init__(self):
        self.xl = excel_file.ExcelFile()
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.load_file)

    def load_file(self):
        self.listWidget.clear()
        self.listWidget.addItem("Загружаем таблицу. Ждите...")
        f = QtWidgets.QFileDialog.getOpenFileName(self, filter="Exel files (*.xlsx *.xls)")
        self.listWidget.addItem('Таблица из файла ' + f[0] + ' загружена.')
        tb_title, tb_raw = self.xl.load_table(f[0])
        tb_by_npp = self.xl.group_by_npp(tb_raw)
        # pprint.pprint(tb_by_npp)
        tb_by_npp_sorted = self.xl.sort_by_value(tb_by_npp)
        # pprint.pprint(tb_by_npp_sorted)
        cum_90_percent_tabl = self.xl.get_90_cum_percent(tb_by_npp_sorted)
        # pprint.pprint(cum_90_percent_tabl)
        self.xl.stat_char(cum_90_percent_tabl)
        self.listWidget.addItem('Таблица обработана.')


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = CpromApp()
    window.show()
    app.exec_()
