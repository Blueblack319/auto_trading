from kiwoom.kiwoom import Kiwoom
from PyQt5.QtWidgets import *
import sys


class UiClass:
    def __init__(self):
        print('Ui_class 입니다.')

        self.app = QApplication(sys.argv)

        self.kiwoom = Kiwoom()

        self.app.exec_()