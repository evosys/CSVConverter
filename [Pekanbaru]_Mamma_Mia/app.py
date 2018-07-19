import os
import sys
import time
import re
import itertools
import unicodedata
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

CODE_STORE = '645594'

NSD        = {'Default':'urn:schemas-microsoft-com:office:spreadsheet', 'o': 'urn:schemas-microsoft-com:office:office', 'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}

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
        # self.cbOutlet.addItem('[YOVA] Gunung Malang', '429703')
        # self.cbOutlet.addItem('[YOVA] Karang Jati', '429696')

    def clean_name(self, name, replace_space_with=None):
        """
        Remove invalid file name chars from the specified name

        :param name: the file name
        :param replace_space_with: if not none replace space with this string
        :return: a valid name for Win/Mac/Linux
        """

        # ref: https://en.wikipedia.org/wiki/Filename
        # ref: https://stackoverflow.com/questions/4814040/allowed-characters-in-filename
        # No control chars, no: /, \, ?, %, *, :, |, ", <, >

        # remove control chars
        name = ''.join(ch for ch in name if unicodedata.category(ch)[0] != 'C')

        cleaned_name = re.sub(r'[/\\?%*:|"<>]', '', name)
        if replace_space_with is not None:
            return cleaned_name.replace(' ', replace_space_with)
        return cleaned_name


    def openXLS(self) :
        fileName, _ = QFileDialog.getOpenFileName(self,"Open File", "","XLS Files (*.xls)")
        if fileName:
            self.lbPath.setText(fileName)
            x = QUrl.fromLocalFile(fileName).fileName()
            self.edFile.setText(x)
            self.edFile.setStyleSheet("""QLineEdit { color: green }""")


    # get PO number
    def getPONO(self, pathXLS) :
        magical_parser = etree.XMLParser(encoding='utf-8', recover=True)

        TREE = etree.parse(pathXLS, magical_parser)

        if len(pathXLS) == 0 :
            QMessageBox.warning(self, "Warning", "Please select XLS file first!", QMessageBox.Ok)
        else :
            x = TREE.xpath('.//ss:Worksheet/ss:Table/ss:Row[@ss:Height="13.73"]/ss:Cell[@ss:MergeAcross="5"]/ss:Data[@ss:Type="String"]', namespaces=NSD)

            res = re.sub(r'[^\w]', ' ', x[0].text).strip()

            return res


    # get Barcode
    def getBRC(self, pathXLS) :
        magical_parser = etree.XMLParser(encoding='utf-8', recover=True)

        TREE = etree.parse(pathXLS, magical_parser)

        if len(pathXLS) == 0 :
            QMessageBox.warning(self, "Warning", "Please select XLS file first!", QMessageBox.Ok)
        else :
            x = TREE.xpath('.//ss:Worksheet/ss:Table/ss:Row[@ss:Height="13.73"]/ss:Cell[@ss:Index="2"]/ss:Data[@ss:Type="String"]', namespaces=NSD)

            result = []

            for i in x :
                result.append(i.text);

            return result


    # get Barcode
    def getQTY(self, pathXLS) :
        magical_parser = etree.XMLParser(encoding='utf-8', recover=True)

        TREE = etree.parse(pathXLS, magical_parser)

        if len(pathXLS) == 0 :
            QMessageBox.warning(self, "Warning", "Please select XLS file first!", QMessageBox.Ok)
        else :
            x = TREE.xpath('.//ss:Worksheet/ss:Table/ss:Row[@ss:Height="13.73"]/ss:Cell[@ss:Index="4"]/ss:Data[@ss:Type="Number"]', namespaces=NSD)

            result = []

            for i in x :
                result.append(i.text);

            return result


    # get Barcode
    def getMDL(self, pathXLS) :
        magical_parser = etree.XMLParser(encoding='utf-8', recover=True)

        TREE = etree.parse(pathXLS, magical_parser)

        if len(pathXLS) == 0 :
            QMessageBox.warning(self, "Warning", "Please select XLS file first!", QMessageBox.Ok)
        else :
            x = TREE.xpath('.//ss:Worksheet/ss:Table/ss:Row[@ss:Height="13.73"]/ss:Cell[@ss:Index="12"]/ss:Data[@ss:Type="Number"]', namespaces=NSD)

            result = []

            for i in x :
                result.append(i.text);

            return result


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
            tempFile = str(self.getPONO(pathXLS))
            FileName = self.clean_name(tempFile) + '.csv'
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

            a = CODE_STORE
            b = self.getPONO(pathXLS)
            c = self.getBRC(pathXLS)
            d = self.getQTY(pathXLS)
            e = self.getMDL(pathXLS)

            for resA, resB, resC, resD, resE in zip(itertools.repeat(a, len(c)), itertools.repeat(b, len(c)), c, d, e) :
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