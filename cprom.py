#!/usr/bin/env python3

import sys
from PyQt5 import QtWidgets
from gui import design
from modules import excel_file
import re


class CpromApp(QtWidgets.QMainWindow, design.Ui_MainWindow, QtWidgets.QInputDialog):
    def __init__(self):
        self.xl = excel_file.ExcelFile()
        super().__init__()
        self.setupUi(self)
        self.start_text()
        self.pushButton.clicked.connect(self.load_file)

    def load_file(self):
        self.listWidget.clear()
        self.listWidget.addItem("Loading and processing the excel table.\nPlease be patient and wait...")
        f = QtWidgets.QFileDialog.getOpenFileName(self, filter="Exel files (*.xlsx *.xls)")
        if not f[0]:
            self.listWidget.clear()
            self.start_text()
            return
        self.listWidget.addItem('The table from file ' + f[0] + ' was loaded.')
        tb_title, tb_raw = self.xl.load_table_new(f[0])
        # group by NPP
        tb_by_npp = self.xl.group_by_npp(tb_raw)
        # sorted into NPP group
        tb_by_npp_sorted = self.xl.sort_by_value(tb_by_npp)
        # get 90 percent company
        cum_90_percent_table, not_selected_table = self.xl.get_90_cum_percent(tb_by_npp_sorted)
        # choose big company and get states and other
        chosen, strata, average, sigma, count, covar = self.xl.stat_char(cum_90_percent_table)
        # choose in stratas
        strata_final = self.xl.random_choose(strata, average)
        # make final table
        final_table = self.xl.make_final_table(tb_by_npp_sorted, chosen, strata_final, not_selected_table)
        # save to file
        f_w = re.sub(r'.xlsx|.xls', '_processed.xlsx', f[0])
        self.xl.write_to_file_new(f_w, tb_title, final_table)
        self.listWidget.addItem('Processing completed.')
        self.listWidget.addItem('File is saved.')

    def start_text(self):
        self.listWidget.addItem(" ")
        self.listWidget.addItem("\nPlease select Excel file")
        self.listWidget.addItem("IMPORTANT!\nThe first column is for grouping, the ninth column is for calculation. Otherwise, the program will not work correctly.")
        # inputD = QtWidgets.QInputDialog(self)  #.setGeometry.(self, 300, 300, 350, 250)
        # inputD.setGeometry(500, 500, 500, 500)
        # text, ok, = inputD.getText(self, 'Input Dialog', 'test')
        # print(text)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = CpromApp()
    window.show()
    app.exec_()
