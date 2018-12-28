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

# variabel for header CSV
HEAD_CODE_STORE = 'code_store'
HEAD_PO_NO      = 'po_no'
HEAD_BARCODE    = 'barcode'
HEAD_QTY        = 'qty'
HEAD_MODAL      = 'modal_karton'

NEWDIR     = 'CSV-output'
DELIM      = ';'

# CODE_STORE = '598998'


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

        # add item to combobox
        self.cbOutlet.addItem('BENGAWAN MULTI TRADING / LNSK, CV.', '378977')
        self.cbOutlet.addItem('BENGAWAN RETAIL MANDIRI, PT.', '378966')
        self.cbOutlet.addItem('LUWES GADING', '378938')
        self.cbOutlet.addItem('LUWES GENTAN', '378988')
        self.cbOutlet.addItem('LUWES KESTALAN', '378983')
        self.cbOutlet.addItem('LUWES DELANGGU', '378976')

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
        fileName, _ = QFileDialog.getOpenFileName(self,"Open Products Data", "","XLSX Files (*.xlsx)")
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
                    msg = "The '.xlsx' file has been corrupted."
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
                    msg = "The '.xlsx' file has been corrupted."
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

        if os.path.exists(str(resPathFile)) :
            os.remove(str(resPathFile))
        else :
            # os.makedirs(os.path.dirname(resPathFile), exist_ok=True)
            distutils.dir_util.mkpath(os.path.dirname(str(resPathFile)))

        return str(resPathFile)


    # get Data Barcode from Master Data
    def dataBarcode(self) :
        result = []

        sheet = self.funcXLRD()

        if sheet :

            totalrow = sheet.nrows - 1

            rh = self.get_cell_range(1, 0, 1, totalrow)

            for z in rh :
                filt = filter(None, z)

                for i in filt :
                    result.append(i)

            # fin = [ x for x in result if x.isdigit() ]

            return result



    # get Data PLU from Master Data
    def dataPLU(self) :
        result = []

        sheet = self.funcXLRD()

        if sheet :

            totalrow = sheet.nrows - 1

            rh = self.get_cell_range(2, 0, 2, totalrow)

            for z in rh :
                filt = filter(None, z)

                for i in filt :
                    result.append(i)

            # fin = [ x for x in result if x.isdigit() ]

            return result


    # get Number PO
    def getPONO(self) :
        sheet = self.funcXLRD(True)

        if sheet :

            rh = self.get_cell_range(3, 8, 3, 8, True)

            return rh


    # get PLU PO
    def getPLU(self) :
        result = []

        sheet = self.funcXLRD(True)

        if sheet :

            totalrow = sheet.nrows - 1

            rh = self.get_cell_range(1, 0, 1, totalrow, True)

            for z in rh :
                filt = filter(None, z)

                for i in filt :
                    if not self.is_float(i) :
                        result.append(i)

            fin = [ x for x in result if x.isdigit() ]

            return fin


    # get QTY PO
    def getQTY(self, num) :
        result = []

        sheet = self.funcXLRD(True)

        if sheet :

            totalrow = sheet.nrows - 1

            rh = self.get_cell_range(5, 0, 5, totalrow, True)

            for z in rh :
                filt = filter(None, z)

                for i in filt :
                    if self.is_float(i) :
                        result.append(int(i))

            fin = result[:num]

            return fin


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


    # combine dataPLU and dataBarcode
    def resDataBCD(self) :
        result = []
        if self.dataBarcode() :
            dataList = [list(z) for z in zip(self.dataBarcode(), self.dataPLU())]
        else :
            sys

        # b = [x for x in dataList if x[0] in self.getPLU()]
        if self.dataPLU() :
            for v in self.getPLU() :
                # c = list(filter(lambda x:x[0]==v, dataList))
                x = self.search_nested(dataList, v)
                if x is not None :
                    result.append(x)
                else :
                    result.append([None, v])

            return result
        else :
            return


    # final barcode for PO
    def finalBarcode(self) :
        result = []

        resData = self.resDataBCD()

        if resData :
            for i in resData :
                result.append(i[0])

            return result
        else :
            return


    # final PLU for PO
    def finalPLU(self) :
        result = []

        resData = self.resDataBCD()

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

            reply = QMessageBox.warning(self, "Warning", "Please select Products Data file first!", QMessageBox.Ok)

        else :

            # path file
            resPathFile, resultPath, resFilename = self.PathOut()

            # make as variabel
            CODE_STORE = str(self.cbOutlet.itemData(self.cbOutlet.currentIndex()))
            ponum = self.getPONO()

            if ponum :
                responum = ponum[0][0]

            brc = self.finalBarcode()
            plu = self.finalPLU()

            if brc :
                qty = self.getQTY(len(brc))
                mdl = self.getMDL(len(brc))

            if brc :
                dataList = [list(z) for z in zip(itertools.repeat(CODE_STORE, len(brc)), itertools.repeat(responum, len(brc)), brc, qty, mdl)]

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

            if self.CheckNone(self.resDataBCD(), logFile) :

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

            # z = self.CheckNone(brc)

            # prepare write CSV



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
