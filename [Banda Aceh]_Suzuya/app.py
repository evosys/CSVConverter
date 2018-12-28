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

import csv
from collections import defaultdict

# variabel for header CSV
HEAD_CODE_STORE = 'code_store'
HEAD_PO_NO      = 'po_no'
HEAD_BARCODE    = 'barcode'
HEAD_QTY        = 'qty'
HEAD_MODAL      = 'modal_karton'
HEAD_UOM        = 'uom'

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
        self.cbOutlet.addItem('SUZUYA I', '549518')
        self.cbOutlet.addItem('SUZUYA II', '549519')


    def openXLS(self) :
        fileName, _ = QFileDialog.getOpenFileName(self,"Open File", "","CSV Files (*.csv)")
        if fileName:
            self.lbPath.setText(fileName)
            x = QUrl.fromLocalFile(fileName).fileName()
            self.edFile.setText(x)
            self.edFile.setStyleSheet("""QLineEdit { color: green }""")


    # get PO number
    def getPONO(self, pathXLS) :
        COLUMNS = defaultdict(list)

        with open(pathXLS, 'r') as f :
            reader = csv.DictReader(f) # read rows into a dictionary format
            for row in reader :
                for (k,v) in row.items() :
                    COLUMNS[k].append(v)

        result = COLUMNS['NoPO']
        return result;


    # get barcode
    def getBRC(self, pathXLS) :
        COLUMNS = defaultdict(list)

        with open(pathXLS, 'r') as f :
            reader = csv.DictReader(f) # read rows into a dictionary format
            for row in reader :
                for (k,v) in row.items() :
                    COLUMNS[k].append(v)

        result = COLUMNS['Barcode']
        return result;

    # get QTY
    def getQTY(self, pathXLS) :
        COLUMNS = defaultdict(list)

        with open(pathXLS, 'r') as f :
            reader = csv.DictReader(f) # read rows into a dictionary format
            for row in reader :
                for (k,v) in row.items() :
                    COLUMNS[k].append(v)

        result = COLUMNS['Qty']
        return result;


    # get QTY
    def getUOM(self, pathXLS) :
        COLUMNS = defaultdict(list)

        with open(pathXLS, 'r') as f :
            reader = csv.DictReader(f) # read rows into a dictionary format
            for row in reader :
                for (k,v) in row.items() :
                    COLUMNS[k].append(v)

        result = COLUMNS['Satuan']
        return result;


    # get MODAL
    def getMODAL(self, pathXLS) :
        COLUMNS = defaultdict(list)

        with open(pathXLS, 'r') as f :
            reader = csv.DictReader(f) # read rows into a dictionary format
            for row in reader :
                for (k,v) in row.items() :
                    COLUMNS[k].append(v)

        result = COLUMNS['H.Beli']
        return result;


    # button convert CSV
    def BtnCnv(self) :
        # get path directory
        pathXLS = self.lbPath.text()
        # set path to variabel
        resPath, resFilename = os.path.split(pathXLS)
        current_dir = os.getcwd()
        resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

        if len(pathXLS) == 0:

            QMessageBox.warning(self, "Warning", "Please select CSV file first!", QMessageBox.Ok)

        else :

            tempFile = self.getPONO(pathXLS)
            FileName = resFilename
            current_dir = os.getcwd()
            resPathFile = os.path.abspath(os.path.join(current_dir, NEWDIR, FileName))
            resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

            # create and open .csv file
            if os.path.exists(resPathFile) :
                os.remove(resPathFile)
            else :
                os.makedirs(os.path.dirname(resPathFile), exist_ok=True)


            csv = open(resPathFile, 'w+')

            # write first header
            csv.write(HEAD_CODE_STORE + DELIM + HEAD_PO_NO + DELIM + HEAD_BARCODE + DELIM + HEAD_UOM + DELIM + HEAD_QTY + DELIM + HEAD_MODAL)

            # write new line
            csv.write("\n")

            a = str(self.cbOutlet.itemData(self.cbOutlet.currentIndex()))
            b = self.getPONO(pathXLS)
            c = self.getBRC(pathXLS)
            d = self.getUOM(pathXLS)
            e = self.getQTY(pathXLS)
            f = self.getMODAL(pathXLS)

            for resA, resB, resC, resD, resE, resF in zip(itertools.repeat(a, len(c)), b, c, d, e, f) :
                resD = resD.split('/', 1)[0] # remove '/ 'on UOM
                resE = str(resE).split('.')[0] # remove decimal on QTY
                csv.write(resA+DELIM+resB+DELIM+resC+DELIM+resD+DELIM+resE+DELIM+resF)
                csv.write("\n")

            csv.close()

            reply = QMessageBox.information(self, "Information", "Success!", QMessageBox.Ok)
            if reply == QMessageBox.Ok :
                os.startfile(str(resultPath))


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
