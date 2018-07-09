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

CODE_STORE = '358971'

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

    # get last row
    def getLastRow(self) :
        res = []
        sheet = self.funcXLRD()

        return sheet.nrows

    # get barcode
    def getPONOM(self, ROW) :
        col = 2
        res = []

        sheet = self.funcXLRD()

        if type(ROW) is list :

            for this in ROW :

                tmp = sheet.cell_value(this, col)
                res.append(tmp)

                # replace '/' with '-'
                res = [w.replace('/', '-') for w in res]

            return res

        else :
            QMessageBox.warning(self, "Error", "Can't get content file!", QMessageBox.Ok)

    # get ROW
    def GET_ROW_PO(self) :
        col = 2
        res = []

        sheet = self.funcXLRD()

        rows = sheet.nrows
        columns = sheet.ncols

        for row in range(rows):
            if sheet.cell_value(row, col) != '' :
                res.append(row)

        # remove 0 from list
        res.remove(0)

        # return
        return res


    # PATH FILE
    def openXLS(self) :
        fileName, _ = QFileDialog.getOpenFileName(self,"Open File", "","XLS Files (*.xls)")
        if fileName:
            self.lbPath.setText(fileName)
            x = QUrl.fromLocalFile(fileName).fileName()
            self.edFile.setText(x)
            self.edFile.setStyleSheet("""QLineEdit { color: green }""")


    # function xlrd
    def funcXLRD(self) :
        # PATH file
        pathXLS = self.lbPath.text()

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
    def get_cell_range(self, start_col, start_row, end_col, end_row):
        sheet = self.funcXLRD()
        return [sheet.row_values(row, start_colx=start_col, end_colx=end_col+1) for row in range(start_row, end_row+1)]

    def CreateDir(self, cDIR, nDir, filename) :

        resPathFile =  os.path.abspath(os.path.join(cDIR, nDir, "{}.csv".format(filename)))

        if os.path.exists(resPathFile) :
            os.remove(resPathFile)
        else :
            # os.makedirs(os.path.dirname(resPathFile), exist_ok=True)
            distutils.dir_util.mkpath(os.path.dirname(resPathFile))

        return resPathFile


    # formating
    def getFormating(self, ROW, colbar) :
        tmpArr = ROW
        tmpArr.append(self.getLastRow())
        res = []
        for (i, value) in enumerate(self.GET_ROW_PO()) :
            gto = i+1

            a = [colbar, value, colbar, tmpArr[gto]-1]

            res.append(a)

        return res

    # actual Value
    def resValue(self, ROW, colbar) :
        res = []
        tmcol = self.getFormating(ROW, colbar)
        start_col, start_row, end_col, end_row = zip(*tmcol)

        for a, b, c, d in zip(start_col, start_row, end_col, end_row) :
            tmp = self.get_cell_range(a, b , c , d)

            res.append(tmp)

        print(res)

        return res

    def open_file(self, filename):
        if sys.platform == "win32":
            os.startfile(filename)
        else:
            opener ="open" if sys.platform == "darwin" else "xdg-open"
            subprocess.call([opener, filename])

    # button convert CSV
    def BtnCnv(self) :

        current_dir = os.getcwd()

        # get path directory
        rowPO = self.GET_ROW_PO()
        filePO = self.getPONOM(rowPO)
        print(filePO)

        brcd = self.resValue(rowPO, 3)
        mdl = self.resValue(rowPO, 4)
        qty = self.resValue(rowPO, 5)

        resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

        for file, barcodes, quantity, modal in zip(filePO, brcd, qty, mdl) :
            resPathFile = self.CreateDir(current_dir, NEWDIR, file)
            with open(resPathFile, "w+") as csv :

                # Write CSV file

                # write first header
                csv.write(HEAD_CODE_STORE + DELIM + HEAD_PO_NO + DELIM + HEAD_BARCODE + DELIM + HEAD_QTY + DELIM + HEAD_MODAL)

                # write new line
                csv.write("\n")

                for br, qt, md in zip(barcodes, quantity, modal) :
                    for resCD, resPO, resBC, resQT, resMD in zip(itertools.repeat(CODE_STORE, len(br)), itertools.repeat(file, len(br)), br, qt, md) :
                        resQT = str(int(float(resQT)))
                        resMD = str(int(float(resMD)))
                        csv.write(resCD+DELIM+resPO+DELIM+resBC+DELIM+resQT+DELIM+resMD+'\n')

                # close
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