#!/usr/bin/python3
# -*-coding:utf-8 -*-

# Reference:**********************************************
# @Time    : 6/15/2021 10:53 PM
# @Author  : Gaopeng.Bai
# @File    : main.py.py
# @User    : gaope
# @Software: PyCharm
# @Description: 
# Reference:**********************************************
import sys
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

from GUI.mainGui import Ui_MainWindow


class GUI_main(QMainWindow, Ui_MainWindow):
    def __init__(self, main_window):
        super(GUI_main, self).__init__()
        self.setupUi(main_window)


if __name__ == '__main__':
    app = QApplication(sys.argv)  # initialize application
    MainWindow = QMainWindow()  # Create main window
    main_ui = GUI_main(MainWindow)  # Create UI window
    MainWindow.show()  # present window
    # It returns 0 after the message loop ends, and then calls sys.exit (0) to
    # exit the program
    sys.exit(app.exec_())