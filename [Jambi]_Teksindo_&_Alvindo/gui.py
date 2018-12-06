# -*- coding: utf-8 -*-
# @Author: ichadhr
# @Date:   2018-10-27 21:13:08
# @Last Modified by:   richard.hari@live.com
# @Last Modified time: 2018-10-30 02:14:56
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
        MainWindow.resize(331, 165)
        MainWindow.setMinimumSize(QtCore.QSize(1, 0))
        MainWindow.setMaximumSize(QtCore.QSize(1000, 1000))
        MainWindow.setAutoFillBackground(False)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.btOpen = QtWidgets.QPushButton(self.centralwidget)
        self.btOpen.setGeometry(QtCore.QRect(240, 10, 81, 22))
        self.btOpen.setObjectName("btOpen")
        self.edFile = QtWidgets.QLineEdit(self.centralwidget)
        self.edFile.setEnabled(False)
        self.edFile.setGeometry(QtCore.QRect(10, 10, 221, 22))
        self.edFile.setFrame(True)
        self.edFile.setReadOnly(True)
        self.edFile.setObjectName("edFile")
        self.btCnv = QtWidgets.QPushButton(self.centralwidget)
        self.btCnv.setGeometry(QtCore.QRect(90, 104, 151, 41))
        self.btCnv.setObjectName("btCnv")
        self.lbPath = QtWidgets.QLabel(self.centralwidget)
        self.lbPath.setGeometry(QtCore.QRect(20, 40, 47, 13))
        self.lbPath.setObjectName("lbPath")
        self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox.setGeometry(QtCore.QRect(10, 40, 311, 55))
        self.groupBox.setObjectName("groupBox")
        self.cbOutlet = QtWidgets.QComboBox(self.groupBox)
        self.cbOutlet.setGeometry(QtCore.QRect(10, 20, 291, 22))
        self.cbOutlet.setObjectName("cbOutlet")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 331, 21))
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
        self.btOpen.setText(_translate("MainWindow", "Open"))
        self.btCnv.setText(_translate("MainWindow", "Convert CSV"))
        self.lbPath.setText(_translate("MainWindow", "TextLabel"))
        self.groupBox.setTitle(_translate("MainWindow", "Select store:"))

import res_rc

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

