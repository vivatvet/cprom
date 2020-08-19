#!/usr/bin/env python3

import sys
from PyQt5 import QtWidgets
from gui import gui


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = gui.CpromApp()
    window.show()
    app.exec_()
