import sys
from PyQt5 import uic
from PyQt5.QtWidgets import QMainWindow, QApplication
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import QThreadPool, QObject, pyqtSignal, pyqtSlot, QRunnable

from constants import MSG_SUCCESS, MSG_WARNING, MSG_ERROR, showMessageBox
from generate import create_pdf, load_spreadsheet_data

try:
    from ctypes import windll  # Only exists on Windows.
    myappid = 'Brittany PDF Creator'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

class Signals(QObject):
    completed = pyqtSignal(dict)

####### Thread for generate pdf report file with sheetId and sheetName ########
class Gen(QRunnable):
    def __init__(self, msheetId, wsheetId, sheetName, genType):
        super().__init__()
        self.msheetId = msheetId
        self.wsheetId = wsheetId
        self.sheetName = sheetName
        self.genType = genType
        self.signal = Signals()

    @pyqtSlot()
    def run(self):
        data = create_pdf(self.wsheetId, self.sheetName, self.msheetId, self.genType)
        self.signal.completed.emit(data)

####### Thread for load spreadsheet data by sheetId ########
class LoadSS(QRunnable):
    def __init__(self, wsheetId, msheetId):
        super().__init__()
        self.wsheetId = wsheetId
        self.msheetId = msheetId
        self.signal = Signals()

    @pyqtSlot()
    def run(self):
        data = load_spreadsheet_data(self.wsheetId, self.msheetId)
        self.signal.completed.emit(data)

####### Main Window ########
class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        try:
            uic.loadUi('assets/main.ui', self)
            self.loading_sheet = QMovie("assets/images/loading.gif")
            self.loading_gen = QMovie("assets/images/loading.gif")
            self.edit_wsheetid.setText("1ta2t7cQarx6dzacKHLevtYYOsnV_gei9N2z7lHvlhLs")
            self.edit_msheetid.setText("18wZ_UlyQKmhzygdb8nk8I6xAyIPvxJm3Ofh58d1NKZs")

            self.btn_gen.clicked.connect(self.performGeneration)
            self.btn_load_sheet.clicked.connect(self.performLoadSheet)
        except:
            showMessageBox("Please check <assets/main.ui> file.", MSG_ERROR)

    def performLoadSheet(self):
        self.btn_load_sheet.setEnabled(False)
        self.combo_sheet_names.setEnabled(False)
        self.combo_gen_type.setEnabled(False)
        self.btn_gen.setEnabled(False)
        self.combo_sheet_names.clear()

        self.movie_sheet_loading.setMovie(self.loading_sheet)
        self.loading_sheet.start()

        self.wsheetId = self.edit_wsheetid.text()
        self.msheetId = self.edit_msheetid.text()
        if self.wsheetId.strip() == "":
            showMessageBox("Input or copy and paste your work spreadsheet ID.", MSG_WARNING)
            return
        if self.msheetId.strip() == "":
            showMessageBox("Input or copy and paste your master spreadsheet ID.", MSG_WARNING)
            return

        pool = QThreadPool.globalInstance()
        loadSS = LoadSS(self.wsheetId, self.msheetId)
        loadSS.signal.completed.connect(self.updateLoadingSS)
        pool.start(loadSS)

    def performGeneration(self):
        self.genType = self.combo_gen_type.currentText()
        if self.genType == "Top10":
            gen_type = 0
        else:
            gen_type = 1
        self.sheetName = self.combo_sheet_names.currentText()
        self.movie_gen_loading.setMovie(self.loading_gen)
        self.loading_gen.start()

        self.btn_load_sheet.setEnabled(False)
        self.edit_wsheetid.setEnabled(False)
        self.edit_msheetid.setEnabled(False)
        self.combo_sheet_names.setEnabled(False)
        self.combo_gen_type.setEnabled(False)
        self.btn_gen.setEnabled(False)

        pool = QThreadPool.globalInstance()
        gen = Gen(self.msheetId, self.wsheetId, self.sheetName, gen_type)
        gen.signal.completed.connect(self.updateLoadingGen)
        pool.start(gen)

    def updateLoadingSS(self, res):
        self.loading_sheet.stop()
        self.movie_sheet_loading.clear()
        self.btn_load_sheet.setEnabled(True)
        if res["status"] == MSG_SUCCESS:
            self.combo_sheet_names.addItems(res["data"])
            self.combo_sheet_names.setEnabled(True)
            self.combo_gen_type.setEnabled(True)
            self.btn_gen.setEnabled(True)
        else:
            self.combo_sheet_names.setEnabled(False)
            self.combo_gen_type.setEnabled(False)
            self.btn_gen.setEnabled(False)
            showMessageBox(res["msg"], res["status"])

    def updateLoadingGen(self, res):
        self.loading_gen.stop()
        self.movie_gen_loading.clear()
        showMessageBox(res["msg"], res["status"])

        self.btn_load_sheet.setEnabled(True)
        self.edit_wsheetid.setEnabled(True)
        self.edit_msheetid.setEnabled(True)
        self.combo_sheet_names.setEnabled(True)
        self.combo_gen_type.setEnabled(True)
        self.btn_gen.setEnabled(True)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mw = MainWindow()
    mw.show()
    sys.exit(app.exec_())