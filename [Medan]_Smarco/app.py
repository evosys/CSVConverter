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

# variabel for header CSV
HEAD_CODE_STORE = 'code_store'
HEAD_PO_NO      = 'po_no'
HEAD_BARCODE    = 'barcode'
HEAD_QTY        = 'qty'
HEAD_MODAL      = 'modal_karton'

NEWDIR     = 'CSV-output'
DELIM      = ';'

CODE_STORE = '571822'

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
        fileName, _ = QFileDialog.getOpenFileName(self,"Open File", "","XLS Files (*.xls *.xlsx)")
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
                msg = "The '.xls' file has been corrupted."
                errorSrv = QMessageBox.critical(self, "Error", msg, QMessageBox.Abort)
                sys.exit(0)


    # create directory if not exist
    def CreateDir(self, cDIR, nDir, filename) :

        resPathFile =  Path(os.path.abspath(os.path.join(cDIR, nDir, "{}.csv".format(filename))))

        if os.path.exists(resPathFile) :
            os.remove(resPathFile)
        else :
            # os.makedirs(os.path.dirname(resPathFile), exist_ok=True)
            distutils.dir_util.mkpath(os.path.dirname(resPathFile))

        return resPathFile


    # open file
    def open_file(self, filename):
        if sys.platform == "win32":
            os.startfile(filename)
        else:
            opener ="open" if sys.platform == "darwin" else "xdg-open"
            subprocess.call([opener, filename])


    # get cell range
    def get_cell_range(self, start_col, start_row, end_col, end_row):
        sheet = self.funcXLRD()
        return [sheet.row_values(row, start_colx=start_col, end_colx=end_col+1) for row in range(start_row, end_row+1)]


    # get PO number
    def get_ponum(self) :
        sheet = self.funcXLRD()

        result = sheet.cell_value(rowx=5, colx=2)
        # result = result.replace('/', '-')

        return result

    def get_ponum_custum(self) :
        sheet = self.funcXLRD()

        result = sheet.cell_value(rowx=5, colx=1)
        # result = result.replace('/', '-')

        return result


    # get Barcode
    def get_barcode(self) :
        sheet = self.funcXLRD()
        totalrow = sheet.nrows-12

        tmp = self.get_cell_range(1, 14, 1, totalrow)

        evenVal = tmp[1:][::2]

        return evenVal

    # get Barcode
    def get_barcode_custom(self) :
        sheet = self.funcXLRD()
        totalrow = sheet.nrows-12

        tmp = self.get_cell_range(0, 14, 0, totalrow)

        evenVal = tmp[1:][::2]

        return evenVal

    # get QTY
    def get_qty(self) :
        sheet = self.funcXLRD()
        totalrow = sheet.nrows-12

        tmp = self.get_cell_range(8, 14, 8, totalrow)

        evenVal = tmp[1:][::2]

        return evenVal


    # get modal karton
    def get_mdl(self) :
        sheet = self.funcXLRD()
        totalrow = sheet.nrows-12

        tmp = self.get_cell_range(7, 14, 7, totalrow)

        evenVal = tmp[1:][::2]

        return evenVal


    # button convert CSV
    def BtnCnv(self) :
        current_dir = os.getcwd()
        # PATH file
        pathXLS = self.lbPath.text()
        resPath, resFilename = os.path.split(os.path.splitext(pathXLS)[0])
        resPathFile = self.CreateDir(current_dir, NEWDIR, resFilename)
        resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

        # make as variabel
        TMPponum = self.get_ponum()

        if not TMPponum :
            ponum = self.get_ponum_custum()

        else :
            ponum = TMPponum


        TMPbrc = self.get_barcode()

        brcCHK = all(x == '' for i in TMPbrc for x in i)
        if brcCHK :
            brc = self.get_barcode_custom()
        else :
            brc = self.get_barcode()

        qty = self.get_qty()
        mdl = self.get_mdl()

        # prepare write CSV
        with open(resPathFile, "w+") as csv :

            # write first header
            csv.write(HEAD_CODE_STORE + DELIM + HEAD_PO_NO + DELIM + HEAD_BARCODE + DELIM + HEAD_QTY + DELIM + HEAD_MODAL)

            # write new line
            csv.write("\n")

            for br, qt, md in zip(brc, qty, mdl) :
                for resCD, resPO, resBC, resQT, resMD in zip(itertools.repeat(CODE_STORE, len(br)), itertools.repeat(ponum, len(br)), br, qt, md) :
                    resQT = str(int(float(resQT))) # format to string
                    resMD = str(float(resMD)) # format to string
                    resPO = resPO.strip()

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