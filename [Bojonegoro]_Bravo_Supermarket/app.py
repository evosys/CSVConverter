# -*- coding: utf-8 -*-
# @Author: ichadhr
# @Date:   2018-08-13 13:20:49
# @Last Modified by:   richard.hari@live.com
# @Last Modified time: 2018-10-19 11:34:24
import sys
import time
import os
import appinfo
import itertools
import subprocess
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from gui import Ui_MainWindow
from pathlib import Path
import xlrd
import distutils.dir_util

# variabel for header CSV
HEAD_CODE_STORE = 'code_store'
HEAD_PO_NO      = 'po_no'
HEAD_BARCODE    = 'barcode'
HEAD_QTY        = 'qty'
HEAD_MODAL      = 'modal_karton'

NEWDIR     = 'CSV-output'
DELIM      = ';'

CODE_STORE = '043833'

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

        # button Open
        self.btOpen.clicked.connect(self.openXLS)

        # button convert
        self.btCnv.clicked.connect(self.BtnCnv)

        # status bar
        self.statusBar().showMessage('v'+appinfo._version)

        # hide label path
        self.lbPath.hide()
        self.lbPath.clear()


    # PATH FILE
    def openXLS(self) :
        pathFiles = []
        listFiles = []
        fileNames, _ = QFileDialog.getOpenFileNames(self,"Open File", "","XLS Files (*.xls)")
        if fileNames:
            for _i in fileNames :
                Fnames = QUrl.fromLocalFile(_i).fileName()
                listFiles.append(Fnames)
                pathFiles.append(_i)

            lblFiles = "✣ ".join(pathFiles)
            txtFiles = " | ".join(listFiles)

            self.lbPath.setText(lblFiles)
            self.edFile.setText(txtFiles)
            self.edFile.setStyleSheet("""QLineEdit { color: green }""")

    def combine_list(self, pathFiles) :
        result = []

        for _i in pathFiles :
            result.append(_i)

        return result



    def BtnCnv(self) :

        pathFile = self.lbPath.text()
        pathFiles = pathFile.split('✣')


        for _x in pathFiles :
            brc = self.getBRC(_x.tream())

        print(brc)


    # function xlrd
    def funcXLRD(self, pathXLS) :
        # PATH file
        # pathXLS = self.lbPath.text()

        if len(pathXLS) == 0:

            QMessageBox.warning(self, "Warning", "Please select XLS file first!", QMessageBox.Ok)

        else :
            try :
                book = xlrd.open_workbook(pathXLS, ragged_rows=True)
                sheet = book.sheet_by_index(0)

                return sheet

            except xlrd.XLRDError as e:
                msg = "Unsupported format, or corrupt file !"
                errorSrv = QMessageBox.critical(self, "Error", msg, QMessageBox.Abort)
                sys.exit(0)


    # function get cell range
    def get_cell_range(self, pathFile, start_col, start_row, end_col, end_row):
        sheet = self.funcXLRD(pathFile)
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


    # Get Barcode
    def getPONO(self, pathFile) :

        result = self.get_cell_range(pathFile, 1, 5, 1, 5)
        return result[0][0]


    # Get Barcode
    def getBRC(self, pathFile) :

        result = []

        sheet = self.funcXLRD(pathFile)
        totalrow = sheet.nrows-1

        rh = self.get_cell_range(pathFile, 2, 0, 2, totalrow)

        for z in rh :
            filt = filter(None, z)
            for i in filt :
                result.append(i)

        fin = [ x for x in result if x.isdigit() ]

        return fin


    # get Quantity
    def getQTY(self, pathFile) :
        result = []

        sheet = self.funcXLRD(pathFile)
        totalrow = sheet.nrows-1

        rh = self.get_cell_range(pathFile, 7, 0, 7, totalrow)

        for z in rh :
            filt = filter(None, z)
            for i in filt :
                if isinstance(i, float) :
                    x = int(float(i))
                    result.append(x)

        return result



    # Get Modal
    def getMDL(self, pathFile) :
        result = []

        sheet = self.funcXLRD(pathFile)
        totalrow = sheet.nrows-1

        rh = self.get_cell_range(pathFile, 9, 0, 9, totalrow)

        for z in rh :
            filt = filter(None, z)
            for i in filt :
                if isinstance(i, float) :
                    x = str(i)
                    v = x.replace('.', ',')
                    result.append(v)

        return result


    def BtnCnv1(self) :
        current_dir = os.getcwd()
        # PATH file
        pathXLS = self.lbPath.text()
        resPath, resFilename = os.path.split(os.path.splitext(pathXLS)[0])
        resPathFile = self.CreateDir(current_dir, NEWDIR, resFilename)
        resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

        # make as variabel
        ponum = self.getPONO()
        brc = self.getBRC()
        qty = self.getQTY()
        mdl = self.getMDL()

        # prepare write CSV
        with open(resPathFile, "w+") as csv :

            # write first header
            csv.write(HEAD_CODE_STORE + DELIM + HEAD_PO_NO + DELIM + HEAD_BARCODE + DELIM + HEAD_QTY + DELIM + HEAD_MODAL)

            # write new line
            csv.write("\n")

            for resCD, resPO, resBC, resQT, resMD in zip(itertools.repeat(CODE_STORE, len(brc)), itertools.repeat(ponum, len(brc)), brc, qty, mdl) :
                resQT = str(resQT)
                csv.write(resCD+DELIM+resPO+DELIM+resBC+DELIM+resQT+DELIM+resMD+'\n')


            csv.close()

        reply = QMessageBox.information(self, "Information", "Success!", QMessageBox.Ok)

        if reply == QMessageBox.Ok :
            self.open_file(str(resultPath))


    # button convert CSV
    def BtnCnv1(self) :
        current_dir = os.getcwd()
        # PATH file
        pathXLS = self.lbPath.text()
        resPath, resFilename = os.path.split(os.path.splitext(pathXLS)[0])
        resPathFile = self.CreateDir(current_dir, NEWDIR, resFilename)
        resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

        # make as variabel
        ponum = self.getPONO()
        brc = self.getBRC()
        qty = self.getQTY()
        mdl = self.getMODAL()

        # prepare write CSV
        with open(resPathFile, "w+") as csv :

            # write first header
            csv.write(HEAD_CODE_STORE + DELIM + HEAD_PO_NO + DELIM + HEAD_BARCODE + DELIM + HEAD_QTY + DELIM + HEAD_MODAL)

            # write new line
            csv.write("\n")

            for resCD, resPO, resBC, resQT, resMD in zip(itertools.repeat(CODE_STORE, len(brc)), itertools.repeat(ponum, len(brc)), brc, qty, mdl) :
                resPO = resPO[0][0]
                resQT = str(resQT)
                resMD = str(resMD)
                csv.write(resCD+DELIM+resPO+DELIM+resBC+DELIM+resQT+DELIM+resMD+'\n')


            csv.close()

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
