#!/usr/bin/env python3

from PyQt5 import QtWidgets
import design


class CpromApp(QtWidgets.QMainWindow, design.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
