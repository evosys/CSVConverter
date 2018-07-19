import os
import sys
import time
import re
import itertools
import appinfo
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from gui import Ui_MainWindow
from pathlib import Path
import string

import distutils.dir_util

import xlrd

# variabel for header CSV
HEAD_CODE_STORE = 'code_store'
HEAD_PO_NO      = 'po_no'
HEAD_BARCODE    = 'barcode'
HEAD_QTY        = 'qty'
HEAD_MODAL      = 'modal_karton'

NEWDIR     = 'CSV-output'
DELIM      = ';'

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

        # add item to combobox
        self.cbOutlet.addItem('TJ.MORAWA [IRIAN]', '571881')
        self.cbOutlet.addItem('TEMBUNG [IRIAN]', '571882')
        self.cbOutlet.addItem('AKSARA [IRIAN]', '571883')
        self.cbOutlet.addItem('BAHAGIA [IRIAN]', '571884')
        self.cbOutlet.addItem('MARELAN [IRIAN]', '571885')


    # get last row
    def getLastRow(self) :
        res = []
        sheet = self.funcXLRD()

        return sheet.nrows


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
            sys.exit(0)

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
    def getPONO(self) :

        result = self.get_cell_range(3, 4, 3, 4)
        return result

    # Get Barcode
    def getBRC(self) :

        endrow = self.getLastRow() - 1
        newlist = self.get_cell_range(2, 8, 2, endrow)

        flat_list = []
        res = []

        for sublist in newlist :
            for item in sublist:
                flat_list.append(item)

        result = filter(None, flat_list)

        for i in result :
            res.append(i)

        return res


    # Get Barcode
    def getQTY(self) :

        endrow = self.getLastRow() - 1
        newlist = self.get_cell_range(4, 0, 4, endrow)

        flat_list = []
        res = []

        for sublist in newlist :
            for item in sublist:
                if type(item) == float :
                    flat_list.append(item)

        result = filter(None, flat_list)

        for i in result :
            res.append(int(i))

        return res


    # Get Barcode
    def getMDL(self) :

        endrow = self.getLastRow() - 1
        newlist = self.get_cell_range(11, 0, 11, endrow)

        flat_list = []
        res = []

        for sublist in newlist :
            for item in sublist:
                if type(item) == float :
                    flat_list.append(item)

        result = filter(None, flat_list)

        for i in result :
            res.append(int(i))

        return res

    # format filename
    def format_filename(self, s) :
        valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
        filename = ''.join(c for c in s if c in valid_chars)
        filename = filename.replace(' ','_') # I don't like spaces in filenames.
        return filename


    # button convert CSV
    def BtnCnv(self) :
        current_dir = os.getcwd()
        # PATH file
        pathXLS = self.lbPath.text()
        resPath, resFilename = os.path.split(os.path.splitext(pathXLS)[0])
        resPathFile = self.CreateDir(current_dir, NEWDIR, resFilename)
        resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

        # make as variabel
        code_store = str(self.cbOutlet.itemData(self.cbOutlet.currentIndex()))
        ponum = self.getPONO()
        responum = self.format_filename(ponum[0][0])
        brc = self.getBRC()
        qty = self.getQTY()
        mdl = self.getMDL()

        # prepare write CSV
        with open((self.CreateDir(current_dir, NEWDIR, responum)), "w+") as csv :

            # write first header
            csv.write(HEAD_CODE_STORE + DELIM + HEAD_PO_NO + DELIM + HEAD_BARCODE + DELIM + HEAD_QTY + DELIM + HEAD_MODAL)

            # write new line
            csv.write("\n")

            for resCD, resPO, resBC, resQT, resMD in zip(itertools.repeat(code_store, len(brc)), itertools.repeat(ponum, len(brc)), brc, qty, mdl) :
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