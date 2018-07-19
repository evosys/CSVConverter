import sys
import time
import os
import appinfo
import csv
import datetime
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from gui import Ui_MainWindow
from pathlib import Path
import distutils.dir_util
import string

NOWDATE = datetime.datetime.now()
NEWDIR = 'CSV-output'

HEADER = ['code_store', 'po_no', 'barcode', 'qty', 'modal_karton']

CODESTORE = (
    {'id':'MAXI', 'val':'571799'},
    {'id':'MAXI02', 'val':'571800'},
    {'id':'MXH', 'val':'571805'},
    {'id':'MBJ', 'val':'571806'},
    {'id':'MBT', 'val':'571798'},
    {'id':'MBS', 'val':'571794'},
    {'id':'MBAM', 'val':'571796'},
    {'id':'MBP', 'val':'571803'},
    {'id':'MBL', 'val':'571801'},
    {'id':'MBK', 'val':'571797'},
    {'id':'MBA', 'val':'983934'},
    {'id':'MBB', 'val':'946945'},
    {'id':'MBG', 'val':'571795'},
    {'id':'MBD', 'val':'571802'}
)

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
        fileName, _ = QFileDialog.getOpenFileName(self,"Open File", "","CSV Files (*.csv)")
        if fileName:
            self.lbPath.setText(fileName)
            x = QUrl.fromLocalFile(fileName).fileName()
            self.edFile.setText(x)
            self.edFile.setStyleSheet("""QLineEdit { color: green }""")


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


    # make a list
    def makeList(self) :
        pathXLS = self.lbPath.text()

        result = []
        res = []

        with open(pathXLS, 'r') as f :
            reader = csv.reader(f, delimiter=";")
            for row in reader :
                res.append(row)

        res.remove(res[0])

        x = [list(filter(None, lst)) for lst in res]

        for i in x :
            result.append(i)

        return result


    def filterList(self) :
        result = []

        makeLst = self.makeList()

        for item in makeLst :
            for subItem in item :
                for d in CODESTORE :
                    if d['id'] == subItem :
                        item[0] = d['val']
                        result.append(item)

        return result


    # get uniq list
    def getUniq(self, lst) :
        res = []
        for i in lst :
            res.append([i[0], i[1]])

        r = list(set(map(tuple, res)))
        return r


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

        if len(pathXLS) == 0 :

            QMessageBox.warning(self, "Warning", "Please select CSV file first!", QMessageBox.Ok)

        else :

            resPath, resFilename = os.path.split(os.path.splitext(pathXLS)[0])
            resPathFile = self.CreateDir(current_dir, NEWDIR, resFilename)
            resultPath = Path(os.path.abspath(os.path.join(current_dir, NEWDIR)))

            final_list = self.filterList()
            isUniq = self.getUniq(final_list)

            # create file and generate header
            for fname in isUniq :
                fileName = self.format_filename(fname[1])
                with open(self.CreateDir(current_dir, NEWDIR, fileName), "w+", newline='') as f :
                    writer = csv.writer(f, delimiter=';')
                    writer.writerow(HEADER)
                f.close()


            # write content to file
            for i in final_list :
                for fname in isUniq :
                    if i[0] == fname[0] :
                        fileName = self.format_filename(fname[1])
                        with open(Path(os.path.abspath(os.path.join(current_dir, NEWDIR, "{}.csv".format(fileName)))), "a", newline='') as f :
                            writer = csv.writer(f, delimiter=';')
                            # writer.writerow('ini')
                            writer.writerow(i)
                        f.close()


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