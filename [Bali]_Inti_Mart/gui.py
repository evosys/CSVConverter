# -*- coding: utf-8 -*-
# @Author: ichadhr
# @Date:   2018-11-13 15:41:23
# @Last Modified by:   richard.hari@live.com
# @Last Modified time: 2018-11-13 15:54:59
# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'gui.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(328, 158)
        MainWindow.setMinimumSize(QtCore.QSize(1, 0))
        MainWindow.setMaximumSize(QtCore.QSize(1000, 1000))
        MainWindow.setAutoFillBackground(False)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.btPO = QtWidgets.QPushButton(self.centralwidget)
        self.btPO.setGeometry(QtCore.QRect(240, 10, 81, 22))
        self.btPO.setObjectName("btPO")
        self.edPO = QtWidgets.QLineEdit(self.centralwidget)
        self.edPO.setEnabled(False)
        self.edPO.setGeometry(QtCore.QRect(10, 10, 221, 22))
        self.edPO.setFrame(True)
        self.edPO.setReadOnly(True)
        self.edPO.setObjectName("edPO")
        self.btCnv = QtWidgets.QPushButton(self.centralwidget)
        self.btCnv.setGeometry(QtCore.QRect(90, 80, 151, 41))
        self.btCnv.setObjectName("btCnv")
        self.lbPO = QtWidgets.QLabel(self.centralwidget)
        self.lbPO.setGeometry(QtCore.QRect(20, 70, 47, 13))
        self.lbPO.setText("")
        self.lbPO.setObjectName("lbPO")
        self.edPD = QtWidgets.QLineEdit(self.centralwidget)
        self.edPD.setEnabled(False)
        self.edPD.setGeometry(QtCore.QRect(10, 45, 221, 22))
        self.edPD.setFrame(True)
        self.edPD.setReadOnly(True)
        self.edPD.setObjectName("edPD")
        self.btPD = QtWidgets.QPushButton(self.centralwidget)
        self.btPD.setGeometry(QtCore.QRect(240, 45, 81, 22))
        self.btPD.setObjectName("btPD")
        self.lbPD = QtWidgets.QLabel(self.centralwidget)
        self.lbPD.setGeometry(QtCore.QRect(20, 100, 47, 13))
        self.lbPD.setText("")
        self.lbPD.setObjectName("lbPD")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 328, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        font = QtGui.QFont()
        font.setPointSize(8)
        self.statusbar.setFont(font)
        self.statusbar.setAutoFillBackground(True)
        self.statusbar.setStyleSheet("color: #a5a5a5;")
        self.statusbar.setSizeGripEnabled(False)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.btPO.setText(_translate("MainWindow", "Open PO"))
        self.btCnv.setText(_translate("MainWindow", "Convert CSV"))
        self.btPD.setText(_translate("MainWindow", "Master Data"))

import res_rc

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

