# -*- coding: utf-8 -*-
# @Author: ichadhr
# @Date:   2018-10-02 17:28:58
# @Last Modified by:   richard.hari@live.com
# @Last Modified time: 2018-10-08 14:33:56
import sys
import time
import os
import appinfo
import itertools
import subprocess
import re
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from gui import Ui_MainWindow
from pathlib import Path
import distutils.dir_util
# import warnings

# variabel for header CSV
HEAD_CODE_STORE = 'code_store'
HEAD_PO_NO      = 'po_no'
HEAD_BARCODE    = 'barcode'
HEAD_QTY        = 'qty'
HEAD_MODAL      = 'modal_karton'

NEWDIR     = 'CSV-output'
DELIM      = ';'

# CODE_STORE = '047361'

IS_WIN32 = 'win32' in str(sys.platform).lower()

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)

    return os.path.join(os.path.abspath("."), relative_path)

# warnings.filterwarnings("ignore", message="RuntimeWarning: numpy.dtype size changed, may indicate binary incompatibity. Expected 56, got 52")

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

        # path tabula
        self.tabulaexe = resource_path("tabula/tabula.exe")

        # button Open
        self.btOpen.clicked.connect(self.openPDF)

        # button convert
        self.btCnv.clicked.connect(self.BtnCnv)

        # add item to combobox
        self.cbOutlet.addItem('COCO MART DC', '058616')
        self.cbOutlet.addItem('COCO MART MUMBUL', '058605')
        self.cbOutlet.addItem('COCO MART JIMBARAN', '050018')
        self.cbOutlet.addItem('COCO MART BATUBULAN 13206/43006', '047375')
        self.cbOutlet.addItem('COCO SUPERMARKET UBUD 13208/43008', '047361')

        # status bar
        self.statusBar().showMessage('v'+appinfo._version)

        # hide label path
        self.lbPath.hide()
        self.lbPath.clear()


    # PATH FILE
    def openPDF(self) :
        fileName, _ = QFileDialog.getOpenFileName(self,"Open File", "","PDF Files (*.pdf)")
        if fileName:
            self.lbPath.setText(fileName)
            x = QUrl.fromLocalFile(fileName).fileName()
            self.edFile.setText(x)
            self.edFile.setStyleSheet("""QLineEdit { color: green }""")


    # Create Directory
    def CreateDir(self, cDIR, nDir, filename) :

        resPathFile =  os.path.abspath(os.path.join(cDIR, nDir, "{}.csv".format(filename)))

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

    # running tabula
    def tabula(self, coordinate, pathFile) :

        output = self.launchWithoutConsole(self.tabulaexe, ['-p', 'all', '-a', str(coordinate), str(pathFile)])

        return output


    def launchWithoutConsole(self, command, args):
        """Launches 'command' windowless and waits until finished"""
        startupinfo = subprocess.STARTUPINFO()
        stdin = subprocess.PIPE
        stdout = subprocess.PIPE
        stderr = subprocess.PIPE
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        tmpRes, err = subprocess.Popen([command] + args, startupinfo=startupinfo, stdin=stdin, stderr=stderr, stdout=stdout).communicate()

        result = tmpRes.decode('utf-8').splitlines()

        return result


    def PDFponum(self, pathPDF) :

        crdnt = "126.066,78.466,135.734,172.178"

        result = self.tabula(crdnt, pathPDF)

        return result


    def PDFbarcode(self, pathPDF) :

        crdnt = "180.359,30.866,840.809,107.472"

        tmpResult = self.tabula(crdnt, pathPDF)

        result = self.checkListFloat(tmpResult, True)

        return result


    def PDFqty(self, pathPDF) :
        crdnt = "180.359,382.659,839.322,426.541"

        tmpResult = self.tabula(crdnt, pathPDF)

        result = self.checkListFloat(tmpResult, True)

        return result


    def PDFmodal(self, pathPDF) :
        crdnt = "179.616,309.028,838.578,382.6598"

        tmpResult = self.tabula(crdnt, pathPDF)

        result = self.checkListFloat(tmpResult)

        return result


    # check a list for float type value
    def checkListFloat(self, arList, isfloat = False) :
        result = []

        if isfloat :
            for _i in arList:
                if self.checkFLoat(_i) :
                    result.append([int(float(_i))])
        else :
            for _i in arList:
                res = re.sub('[^\d\.,]', '', _i)
                if res :
                    result.append([res])

        return result


    # check float
    def checkFLoat(self, value) :
        try :
            return float(value).is_integer()
        except ValueError:
            return False


    # button convert CSV
    def BtnCnv(self) :

        current_dir = os.getcwd()

        # PATH file
        pathPDF = self.lbPath.text()
        resPath, resFilename = os.path.split(os.path.splitext(pathPDF)[0])
        resPathFile = self.CreateDir(current_dir, NEWDIR, resFilename)
        resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

        CODE_STORE = str(self.cbOutlet.itemData(self.cbOutlet.currentIndex()))
        tmpponum = self.PDFponum(pathPDF)
        ponum = tmpponum[0]
        brc = self.PDFbarcode(pathPDF)
        qty = self.PDFqty(pathPDF)
        mdl = self.PDFmodal(pathPDF)

        # prepare write CSV
        with open(resPathFile, "w+") as csv :

            # write first header
            csv.write(HEAD_CODE_STORE + DELIM + HEAD_PO_NO + DELIM + HEAD_BARCODE + DELIM + HEAD_QTY + DELIM + HEAD_MODAL)

            # write new line
            csv.write("\n")

            for br, qt, md in zip(brc, qty, mdl) :
                for resCD, resPO, resBC, resQT, resMD in zip(itertools.repeat(CODE_STORE, len(br)), itertools.repeat(ponum, len(br)), br, qt, md) :

                    resBC = str(resBC)
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
