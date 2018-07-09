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

from lxml import etree

# variabel for header CSV
HEAD_CODE_STORE = 'code_store'
HEAD_PO_NO      = 'po_no'
HEAD_BARCODE    = 'barcode'
HEAD_QTY        = 'qty'
HEAD_MODAL      = 'modal_karton'

NEWDIR     = 'CSV-output'
DELIM      = ';'
NSD        = {'Default':'urn:schemas-microsoft-com:office:spreadsheet', 'o': 'urn:schemas-microsoft-com:office:office', 'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}

# CODE_STORE = '548511'

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
        self.cbOutlet.addItem('MEGA MART KERTAPATI', '548511')
        self.cbOutlet.addItem('MEGA MART (KERTAPATI) 2', '964709')
        self.cbOutlet.addItem('MEGA MART KM 12', '548525')
        self.cbOutlet.addItem('MEGA MART KM 18', '548550')
        self.cbOutlet.addItem('MEGA MART PUSRI', '548512')
        self.cbOutlet.addItem('MEGA MART (PLAJU)', '964710')
        self.cbOutlet.addItem('MEGA MART (UTOYO)', '964711')


    def openXLS(self) :
        fileName, _ = QFileDialog.getOpenFileName(self,"Open File", "","XLS Files (*.xls)")
        if fileName:
            self.lbPath.setText(fileName)
            x = QUrl.fromLocalFile(fileName).fileName()
            self.edFile.setText(x)
            self.edFile.setStyleSheet("""QLineEdit { color: green }""")

    # get PO number
    def getPONO(self, pathXLS) :

        TREE = etree.parse(pathXLS)

        if len(pathXLS) == 0 :
            QMessageBox.warning(self, "Warning", "Please select XLS file first!", QMessageBox.Ok)
        else :
            x = TREE.find('/ss:Worksheet/ss:Table/ss:Row[@ss:Height="6.59"]/ss:Cell[@ss:MergeDown="2"]/ss:Data', namespaces=NSD)
            return x

    # get barcode
    def getBRC(self, pathXLS) :

        TREE = etree.parse(pathXLS)

        if len(pathXLS) == 0 :
            QMessageBox.warning(self, "Warning", "Please select XLS file first!", QMessageBox.Ok)
        else :
            x = TREE.xpath('.//ss:Worksheet/ss:Table/ss:Row[@ss:Height="10.99"]/ss:Cell[@ss:StyleID="s6"]/ss:Data[@ss:Type="String"]', namespaces=NSD)
            return x

    # get QTY
    def getQTY(self, pathXLS) :

        TREE = etree.parse(pathXLS)

        if len(pathXLS) == 0 :
            QMessageBox.warning(self, "Warning", "Please select XLS file first!", QMessageBox.Ok)
        else :
            x = TREE.xpath('.//ss:Worksheet/ss:Table/ss:Row[@ss:Height="10.99"]/ss:Cell[@ss:Index="18"]/ss:Data[@ss:Type="String"]', namespaces=NSD)

            new_results = []

            for e in x :
                if e.text not in('QTY') :
                    new_results.append(e)

            x = new_results

            return x

    # get MODAL
    def getMODAL(self, pathXLS) :

        TREE = etree.parse(pathXLS)

        if len(pathXLS) == 0 :
            QMessageBox.warning(self, "Warning", "Please select XLS file first!", QMessageBox.Ok)
        else :
            x = TREE.xpath('.//ss:Worksheet/ss:Table/ss:Row[@ss:Height="10.99"]/ss:Cell[@ss:MergeAcross="3"]/ss:Data[@ss:Type="String"]', namespaces=NSD)

            new_results = []

            for e in x :
                if e.text not in('H. BELI') :
                    new_results.append(e)

            x = new_results

            return x

    # button convert CSV
    def BtnCnv(self) :
        # get path directory
        pathXLS = self.lbPath.text()
        # set path to variabel
        resPath, resFilename = os.path.split(pathXLS)
        current_dir = os.getcwd()
        resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

        if len(pathXLS) == 0:

            QMessageBox.warning(self, "Warning", "Please select XLS file first!", QMessageBox.Ok)

        else :

            tempFile = self.getPONO(pathXLS)
            FileName = tempFile.text[4:] + '.csv'
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
            csv.write(HEAD_CODE_STORE + DELIM + HEAD_PO_NO + DELIM + HEAD_BARCODE + DELIM + HEAD_QTY + DELIM + HEAD_MODAL)
            # write new line
            csv.write("\n")

            a = str(self.cbOutlet.itemData(self.cbOutlet.currentIndex()))
            b = self.getPONO(pathXLS)
            c = self.getBRC(pathXLS)
            d = self.getQTY(pathXLS)
            e = self.getMODAL(pathXLS)

            for resA, resB, resC, resD, resE in zip(itertools.repeat(a,len(c)), itertools.repeat(b,len(c)), c, d, e) :
                # resA = resA
                resB = resB.text[4:]
                resC = resC.text
                resD = re.sub(r"\D", "", resD.text)
                resE = resE.text
                resE = resE.replace(",", "")
                # print(resA+DELIM+resB+DELIM+resC+DELIM+resD+DELIM+resE)
                csv.write(resA+DELIM+resB+DELIM+resC+DELIM+resD+DELIM+resE)
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