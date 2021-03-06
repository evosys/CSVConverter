# -*- coding: utf-8 -*-
# @Author: ichadhr
# @Date:   2018-11-13 15:41:23
# @Last Modified by:   richard.hari@live.com
# @Last Modified time: 2018-11-14 13:50:59
import sys
import time
import os
import appinfo
import itertools
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from gui import Ui_MainWindow
from pathlib import Path
import xlrd
import distutils.dir_util
import webbrowser
from itertools import chain

# variabel for header CSV
HEAD_CODE_STORE = 'code_store'
HEAD_PO_NO      = 'po_no'
HEAD_BARCODE    = 'barcode'
HEAD_QTY        = 'qty'
HEAD_MODAL      = 'modal_karton'

NEWDIR     = 'CSV-output'
DELIM      = ';'

CODE_STORE = '050417'


# main class
class mainWindow(QMainWindow, Ui_MainWindow) :
    def __init__(self) :
        QMainWindow.__init__(self)
        self.setupUi(self)

        # app icon
        self.setWindowIcon(QIcon(':/resources/icon.png'))

        # centering app
        tr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        tr.moveCenter(cp)
        self.move(tr.topLeft())

        # button Open PO
        self.btPO.clicked.connect(self.openXLSPO)

        # button Open PO
        self.btPD.clicked.connect(self.openXLSPD)

        # button convert
        self.btCnv.clicked.connect(self.BtnCnv)

        # status bar
        self.statusBar().showMessage('v'+appinfo._version)

        # hide label Proucts Data
        self.lbPD.hide()
        self.lbPD.clear()


        # hide label Purchase Order
        self.lbPO.hide()
        self.lbPO.clear()


    # PATH FILE
    def openXLSPO(self) :
        fileName, _ = QFileDialog.getOpenFileName(self,"Open Purchase Order", "","XLS Files (*.xls)")
        if fileName:
            self.lbPO.setText(fileName)
            x = QUrl.fromLocalFile(fileName).fileName()
            self.edPO.setText(x)
            self.edPO.setStyleSheet("""QLineEdit { color: green }""")

    def openXLSPD(self) :
        fileName, _ = QFileDialog.getOpenFileName(self,"Open Products Data", "","XLS Files (*.xls)")
        if fileName:
            self.lbPD.setText(fileName)
            x = QUrl.fromLocalFile(fileName).fileName()
            self.edPD.setText(x)
            self.edPD.setStyleSheet("""QLineEdit { color: green }""")


    # function xlrd
    def funcXLRD(self, PurchaseOrder = False) :

        if PurchaseOrder :
            pathXLS = self.lbPO.text()

            if len(pathXLS) == 0:

                reply = QMessageBox.warning(self, "Warning", "Please select Purchase Order file first!", QMessageBox.Ok)

                if reply == QMessageBox.Ok :
                    return False

            else :
                try :

                    book = xlrd.open_workbook(pathXLS, ragged_rows=True)
                    sheet = book.sheet_by_index(0)

                    return sheet

                except xlrd.XLRDError as e:
                    msg = "The '.xls' file has been corrupted."
                    errorSrv = QMessageBox.critical(self, "Error", msg, QMessageBox.Abort)
                    sys.exit(0)
        else :
            pathXLS = self.lbPD.text()
            if len(pathXLS) == 0:

                reply = QMessageBox.warning(self, "Warning", "Please select Products Data file first!", QMessageBox.Ok)

                if reply == QMessageBox.Ok :
                    return False

            else :
                try :
                    book = xlrd.open_workbook(pathXLS, ragged_rows=True)
                    sheet = book.sheet_by_index(0)

                    return sheet

                except xlrd.XLRDError as e:
                    msg = "The '.xls' file has been corrupted."
                    errorSrv = QMessageBox.critical(self, "Error", msg, QMessageBox.Abort)
                    sys.exit(0)


    # get row range
    def get_cell_range(self, start_col, start_row, end_col, end_row, PurchaseOrder = False):
        if PurchaseOrder :
            sheet = self.funcXLRD(True)
        else :
            sheet = self.funcXLRD()

        return [sheet.row_values(row, start_colx=start_col, end_colx=end_col+1) for row in range(start_row, end_row+1)]


    # open file
    def open_file(self, filename):
        if sys.platform == "win32":
            os.startfile(filename)
        else:
            opener ="open" if sys.platform == "darwin" else "xdg-open"
            subprocess.call([opener, filename])


    # create directory if not exist
    def CreateDir(self, cDIR, nDir, filename) :

        resPathFile =  Path(os.path.abspath(os.path.join(cDIR, nDir, "{}.csv".format(filename))))

        if os.path.exists(resPathFile) :
            os.remove(resPathFile)
        else :
            # os.makedirs(os.path.dirname(resPathFile), exist_ok=True)
            distutils.dir_util.mkpath(os.path.dirname(resPathFile))

        return resPathFile


    # get Data Barcode from Master Data
    def dataBarcode(self) :

        sheet = self.funcXLRD()

        if sheet :

            totalrow = sheet.nrows - 1

            rh = self.get_cell_range(7, 0, 7, totalrow)

            return rh


    # get Data PLU from Master Data
    def dataPLU(self) :

        sheet = self.funcXLRD()

        if sheet :

            totalrow = sheet.nrows - 1

            rh = self.get_cell_range(3, 0, 3, totalrow)

            return rh


    # merge Barcode and PLU master data
    def MasterData(self) :
        result = []

        x = self.dataBarcode()
        z = self.dataPLU()

        brc = self.to1d(x)
        plu = self.to1d(z)

        tmpRes = [list(a) for a in zip(brc, plu)]

        for i in tmpRes :
            if i[0] and self.is_float(i[0]):
                i[0] = str(int(i[0]))
                i[1] = str(int(i[1]))

                result.append(i)

        return result


    # get Number PO
    def getPONO(self) :
        sheet = self.funcXLRD(True)

        if sheet :

            rh = self.get_cell_range(3, 2, 3, 2, True)

            return rh


    # get PLU PO
    def getPLU(self) :
        result = []

        sheet = self.funcXLRD(True)

        if sheet :

            totalrow = sheet.nrows - 1

            rh = self.get_cell_range(3, 0, 3, totalrow, True)

            for z in rh :
                filt = filter(None, z)

                for i in filt :
                    if self.is_float(i) :
                        result.append(str(int(float(i))))

            return result


    # get QTY PO
    def getQTY(self) :
        result = []

        sheet = self.funcXLRD(True)

        if sheet :

            totalrow = sheet.nrows - 1

            tmpPLU = self.get_cell_range(3, 0, 3, totalrow, True)
            tmpQTY = self.get_cell_range(7, 0, 7, totalrow, True)

            plu = self.to1d(tmpPLU)
            qty = self.to1d(tmpQTY)

            tmp = [list(a) for a in zip(plu, qty)]

            for i in tmp :
                if i[0] and self.is_float(i[0]):
                    if i[1] :
                        result.append(str(int(i[1])))

            return result



    # get Modal PO
    def getMDL(self, num) :
        result = []

        sheet = self.funcXLRD(True)

        if sheet :

            totalrow = sheet.nrows - 1

            rh = self.get_cell_range(13, 0, 13, totalrow, True)

            for z in rh :
                filt = filter(None, z)

                for i in filt :
                    if self.is_float(i) :
                        result.append(int(i))

            fin = result[:num]

            return fin



    # comparing PLU Master Data with PO
    def comparePLU(self) :
        result = []

        for i in self.getPLU() :
            x = self.search_nested(self.MasterData(), i)
            if x is not None :
                result.append(x)
            else :
                result.append([None, i])

        return result



    # final barcode for PO
    def finalBarcode(self) :
        result = []

        resData = self.comparePLU()

        if resData :
            for i in resData :
                result.append(i[0])

            return result
        else :
            return



    # final PLU for PO
    def finalPLU(self) :
        result = []

        resData = self.comparePLU()

        if resData :
            for i in resData :
                result.append(i[1])

            return result
        else :
            return



    # check if None
    def CheckNone(self, mylist, path) :
        result = any(item[0] == None for item in mylist)

        if result :
            outLog = open(str(path), 'w')
            outLog.write('-------------------------------------------------\n')
            outLog.write('|\tPLU tidak terdapat di Master Data\t|\n')
            outLog.write('-------------------------------------------------\n\n')
            for i in mylist :
                if i[0] == None :
                    outLog.write('Kode PLU\t: ' +i[1]+ '\n')
        return result


    # covert 2d to 1d list
    def to1d(self, aList) :
        result = []

        for i in aList :
            if not i :
                result.append('')
            else :
                for j in i :
                    result.append(j)

        return result


    # search
    def search_nested(self, mylist, val) :
        for i in range(len(mylist)) :
            for j in range(len(mylist[i])) :
            # print i,j
                if mylist[i][j] == val :
                    return mylist[i]


    # is float
    def is_float(self, value) :
        realFloat = 0.1

        if type(value) == type(realFloat):
            return True
        else:
            return False


    # result Path output
    def PathOut(self) :
        current_dir = os.getcwd()
        # PATH file
        pathXLS = self.lbPO.text()
        resPath, resFilename = os.path.split(os.path.splitext(pathXLS)[0])
        resPathFile = self.CreateDir(current_dir, NEWDIR, resFilename)
        resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

        return [resPathFile, resultPath, resFilename]


    # Button Convert
    def BtnCnv(self) :

        pathXLSPD = self.lbPD.text()

        if len(pathXLSPD) == 0:

            reply = QMessageBox.warning(self, "Warning", "Please select Master Data file first!", QMessageBox.Ok)

        else :

            # path file
            resPathFile, resultPath, resFilename = self.PathOut()

            ponum = self.getPONO()

            if ponum :
                responum = ponum[0][0]

            brc = self.finalBarcode()
            plu = self.finalPLU()
            qty = self.getQTY()
            mdl = '1'

            if brc :
                dataList = [list(z) for z in zip(itertools.repeat(CODE_STORE, len(brc)), itertools.repeat(responum, len(brc)), brc, qty, itertools.repeat(mdl, len(brc)))]

            filtered = [x for x in dataList if x[2] is not None]

            with open(resPathFile, "w+") as csv :

                # write first header
                csv.write(HEAD_CODE_STORE + DELIM + HEAD_PO_NO + DELIM + HEAD_BARCODE + DELIM + HEAD_QTY + DELIM + HEAD_MODAL)

                # write new line
                csv.write("\n")

                for i in filtered :

                    csv.write(
                        str(i[0])
                        +DELIM+
                        str(i[1])
                        +DELIM+
                        str(i[2])
                        +DELIM+
                        str(i[3])
                        +DELIM+
                        str(i[4])
                        +'\n')

                csv.close()

            p = Path(resFilename).stem + ".log"

            logFile = str(resultPath)+'\\'+str(p)

            if self.CheckNone(self.comparePLU(), logFile) :

                msg = "Terdapat PLU pada PO yang tidak terdapat pada Mater Data"

                reply = QMessageBox.warning(self, "Warning", msg, QMessageBox.Ok)

                if reply == QMessageBox.Ok :
                    p = Path(resFilename).stem + ".log"

                    self.open_file(str(resultPath))
                    time.sleep(1)
                    self.open_file(str(logFile))


            else :

                reply = QMessageBox.information(self, "Information", "Success!", QMessageBox.Ok)

                if reply == QMessageBox.Ok :
                    self.open_file(str(resultPath))


if __name__ == '__main__' :
    app = QApplication(sys.argv)

    # create splash screen
    splash_pix = QPixmap(':/resources/unilever_splash.png')

    splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
    splash.setWindowFlags(QtCore.Qt.FramelessWindowHint)
    splash.setEnabled(False)

    # adding progress bar
    progressBar = QProgressBar(splash)
    progressBar.setMaximum(10)
    progressBar.setGeometry(17, splash_pix.height() - 20, splash_pix.width(), 50)

    splash.show()

    for iSplash in range(1, 11) :
        progressBar.setValue(iSplash)
        t = time.time()
        while time.time() < t + 0.1 :
            app.processEvents()

    time.sleep(1)

    window = mainWindow()
    window.setWindowTitle(appinfo._appname)
    # window.setWindowFlags(QtCore.Qt.WindowCloseButtonHint)
    # window.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint)
    window.show()
    splash.finish(window)
    sys.exit(app.exec_())
