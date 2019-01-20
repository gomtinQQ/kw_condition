# -*-coding: utf-8 -
import sys, os, re, datetime, copy, json
import xlwings as xw
import resource_rc

import util, kw_util

from PyQt5 import QtCore, QtWidgets
from PyQt5.QtCore import QObject, pyqtSlot, pyqtSignal, QUrl, QEvent
from PyQt5.QtCore import QStateMachine, QState, QTimer, QFinalState
from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget
from mainwindow_ui import Ui_MainWindow

class Stock(QObject):
    sigConditionOccur = pyqtSignal()
    sigBuy = pyqtSignal()
    sigNoBuy = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.ocx = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.fsm = QStateMachine()
        self.createState()

    def createState(self):
        mainState = QState(self.fsm)
        connectedState = QState(QtCore.QState.ParallelStates, mainState)
        processBuyState = QState(connectedState)
        determineBuyProcessBuyState = QState(processBuyState)
        waitingTRlimitProcessBuyState = QState(processBuyState)
        print(determineBuyProcessBuyState)
        determineBuyProcessBuyState.addTransition(self.sigNoBuy, waitingTRlimitProcessBuyState)
        determineBuyProcessBuyState.addTransition(self.sigBuy, waitingTRlimitProcessBuyState)
        self.test_buy()
        determineBuyProcessBuyState.entered.connect(self.determineBuyProcessBuyStateEntered)

    def test_buy(self):
        print("TEST")
        self.sigBuy.emit()

    @pyqtSlot()
    def determineBuyProcessBuyStateEntered(self):
        print("########### determineBuyProcessBuyStateEntered ###############")
        # jongmok_info_dict = []
        # is_log_print_enable = False
        # return_vals = []
        # printLog = ''

        # jongmok_info_dict = self.getConditionOccurList()
        # print(jongmok_info_dict)
        # result = self.sendOrder("buy_" + jongmokCode, kw_util.sendOrderScreenNo,objKiwoom.account_list[0], kw_util.dict_order["신규매수"], jongmokCode, qty, 0, kw_util.dict_order["시장가"], "")



if __name__ == "__main__":
    myApp = QtWidgets.QApplication(sys.argv)
    objKiwoom = Stock()

    sys.exit(myApp.exec_())
