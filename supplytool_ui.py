__author__ = 'Mark Mon Monteros'

from PyQt5 import QtCore, QtGui, QtWidgets
import os, shutil, time
import pandas as pd

class Ui_SupplyTool(object):

    clock_start = time.time() #Time before the operations start

    def __init__(self):
        self.path = os.path.dirname(os.path.realpath(__file__))
        self.dump = self.path + '\\' + '.csv\\'
        self.forF = None
        self.hcF = None
        self.supF = None
        self.hc_dest = None
        self.sup_dest = None
        self.draft = None
        self.fn = None

    def setupUi(self, SupplyTool):
        SupplyTool.setObjectName("SupplyTool")
        SupplyTool.setFixedSize(388, 620)
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Window, brush)
        SupplyTool.setPalette(palette)
        SupplyTool.setAcceptDrops(True)
        SupplyTool.setTabShape(QtWidgets.QTabWidget.Rounded)
        self.centralwidget = QtWidgets.QWidget(SupplyTool)
        self.centralwidget.setObjectName("centralwidget")
        self.forLbl = QtWidgets.QLabel(self.centralwidget)
        self.forLbl.setGeometry(QtCore.QRect(9, 9, 371, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.forLbl.setFont(font)
        self.forLbl.setObjectName("forLbl")
        self.forBtn = QtWidgets.QPushButton(self.centralwidget)
        self.forBtn.setGeometry(QtCore.QRect(9, 60, 75, 23))
        self.forBtn.setObjectName("forBtn")
        self.forBtn.setCheckable(True)
        self.hcLbl = QtWidgets.QLabel(self.centralwidget)
        self.hcLbl.setGeometry(QtCore.QRect(9, 90, 371, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.hcLbl.setFont(font)
        self.hcLbl.setObjectName("hcLbl")
        self.hcBtn = QtWidgets.QPushButton(self.centralwidget)
        self.hcBtn.setGeometry(QtCore.QRect(9, 140, 75, 23))
        self.hcBtn.setObjectName("hcBtn")
        self.hcBtn.setCheckable(True)
        self.supLbl = QtWidgets.QLabel(self.centralwidget)
        self.supLbl.setGeometry(QtCore.QRect(9, 170, 371, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setItalic(False)
        font.setWeight(75)
        self.supLbl.setFont(font)
        self.supLbl.setObjectName("supLbl")
        self.supBtn = QtWidgets.QPushButton(self.centralwidget)
        self.supBtn.setGeometry(QtCore.QRect(9, 220, 75, 23))
        self.supBtn.setObjectName("supBtn")
        self.supBtn.setCheckable(True)
        self.console = QtWidgets.QTextEdit(self.centralwidget)
        self.console.setGeometry(QtCore.QRect(0, 250, 391, 261))
        self.console.setReadOnly(True)
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(85, 255, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(85, 255, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(240, 240, 240))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Base, brush)
        self.console.setPalette(palette)
        font = QtGui.QFont()
        font.setBold(False)
        font.setWeight(50)
        self.console.setFont(font)
        self.console.setAutoFillBackground(False)
        self.console.setInputMethodHints(QtCore.Qt.ImhMultiLine)
        self.console.setTextInteractionFlags(QtCore.Qt.NoTextInteraction)
        self.console.setObjectName("console")
        self.hcEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.hcEdit.setGeometry(QtCore.QRect(9, 110, 371, 20))
        self.hcEdit.setObjectName("hcEdit")
        self.hcEdit.setReadOnly(True)
        self.supEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.supEdit.setGeometry(QtCore.QRect(9, 190, 371, 20))
        self.supEdit.setObjectName("supEdit")
        self.supEdit.setReadOnly(True)
        self.forEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.forEdit.setGeometry(QtCore.QRect(9, 30, 371, 20))
        self.forEdit.setObjectName("forEdit")
        self.forEdit.setReadOnly(True)
        self.startBtn = QtWidgets.QPushButton(self.centralwidget)
        self.startBtn.setGeometry(QtCore.QRect(10, 520, 371, 51))
        font = QtGui.QFont()
        font.setFamily("Showcard Gothic")
        font.setPointSize(15)
        font.setBold(True)
        font.setItalic(False)
        font.setWeight(75)
        self.startBtn.setFont(font)
        self.startBtn.setObjectName("startBtn")
        self.startBtn.setCheckable(True)
        self.author = QtWidgets.QLabel(self.centralwidget)
        self.author.setGeometry(QtCore.QRect(200, 580, 201, 20))
        self.author.setObjectName("author")
        SupplyTool.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(SupplyTool)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 388, 26))
        self.menubar.setObjectName("menubar")
        SupplyTool.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(SupplyTool)
        self.statusbar.setObjectName("statusbar")
        SupplyTool.setStatusBar(self.statusbar)

        self.retranslateUi(SupplyTool)
        QtCore.QMetaObject.connectSlotsByName(SupplyTool)

        self.console.append('\nINSTRUCTIONS :')
        self.console.append('\n[1] Manually convert Consolidated for reporting xls/xlsx file to CSV UTF-8 format -- can\'t convert password protected files.')
        self.console.append('\n[2] Input converted Consolidated for reporting csv file to the field Consolidated for reporting. « FIRST FIELD »')
        self.console.append('\n[3] Input Excel file (.xls or .xlsx format to the field Consolidated HC report. « SECOND FIELD »')
        self.console.append('\n[4] Input Excel file (.xls or .xlsx format to the field Supply To Be Deleted. « LAST FIELD »')
        self.console.append('\nNote: Supply Tool program will exit automatically if input sheets are incorrect.')
        self.startBtn.clicked.connect(self.start)
        self.forBtn.clicked.connect(self.set_xls)
        self.hcBtn.clicked.connect(self.set_xls)
        self.supBtn.clicked.connect(self.set_xls)

    def retranslateUi(self, SupplyTool):
        _translate = QtCore.QCoreApplication.translate
        SupplyTool.setWindowTitle(_translate("SupplyTool", "Supply Tool"))
        SupplyTool.setWindowIcon(QtGui.QIcon('./img/icon.ico'))
        self.forLbl.setText(_translate("SupplyTool", "Consolidated for reporting sheet :"))
        self.forBtn.setText(_translate("SupplyTool", "Browse..."))
        self.hcLbl.setText(_translate("SupplyTool", "Consolidated HC sheet :"))
        self.hcBtn.setText(_translate("SupplyTool", "Browse..."))
        self.supLbl.setText(_translate("SupplyTool", "Supply To Be Deleted sheet:"))
        self.supBtn.setText(_translate("SupplyTool", "Browse..."))
        self.console.setHtml(_translate("SupplyTool", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:8.25pt; font-weight:400; font-style:normal;\">\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p></body></html>"))
        self.startBtn.setText(_translate("SupplyTool", "S T A R T"))
        self.author.setText(_translate("SupplyTool", "created by: Mark Mon Monteros"))

    def set_xls(self):
        if self.forBtn.isChecked():
            self.forF = None
            self.forF, _ = QtWidgets.QFileDialog.getOpenFileName(None, 'Select File', '', 'Excel Files (*.csv)')
            self.forEdit.setText(self.forF)
            self.forBtn.toggle()
        if self.hcBtn.isChecked():
            self.hcF = None
            self.hcF, _ = QtWidgets.QFileDialog.getOpenFileName(None, 'Select File', '', 'Excel Files (*.xls *.xlsx)')
            self.hcEdit.setText(self.hcF)
            self.hcBtn.toggle()
        if self.supBtn.isChecked():
            self.supF = None
            self.supF, _ = QtWidgets.QFileDialog.getOpenFileName(None, 'Select File', '', 'Excel Files (*.xls *.xlsx)')
            self.supEdit.setText(self.supF)
            self.supBtn.toggle()

    def start(self):
        self.console.setText('Initializing...')
        if self.forF == '' or self.forF is None or self.hcF == '' or self.hcF is None or self.supF == '' or self.supF is None:
            _errorMsg = QtWidgets.QMessageBox()
            _errorMsg.setWindowIcon(QtGui.QIcon('./img/icon.ico'))
            _errorMsg.setWindowTitle('Warning !')
            _errorMsg.setIcon(QtWidgets.QMessageBox.Warning)
            _errorMsg.setText('All fields must contain an item.')
            _errorMsg.exec_()
            self.error()
        else:
            self.run()

    def run(self):
        self.console.append('\n+-----------+-----------+')
        self.console.append(' Supply Automation Tool')
        self.console.append('   by Mark Mon Monteros')
        self.console.append('+-----------+-----------+')
        self.console.append(' Coded in Python ver 3.7*\n')

        self.convert()

    def error(self):
        self.console.append('\n>>> WARNING !!! <<<')
        self.console.append('\nPlease verify if you have a file inserted on each field.')
 
    def convert(self):
        self.console.append('\nConverting excel files to CSV UTF-8 format...')
        try:
            os.mkdir(self.dump)
        except FileExistsError:
            shutil.rmtree(self.dump)
            os.mkdir(self.dump)

        self.console.append('\n [*] ' + str(self.hcF))
        df = pd.read_excel(self.hcF, 'Sheet1', index_col=None)
        df.to_csv(self.hcF.replace(self.hcF[self.hcF.rindex('.'):], '') + '.csv', sep=',', encoding='UTF-8')
        hc_xls = self.hcF[self.hcF.rindex('/'):]
        hc_csv = hc_xls.replace(self.hcF[self.hcF.rindex('.'):], '') + '.csv'
        self.hc_dest = self.dump + str(hc_csv.replace(hc_csv[hc_csv.rindex('/')], ''))
        shutil.move(self.hcF.replace(self.hcF[self.hcF.rindex('.'):], '') + '.csv', self.hc_dest)
        
        self.console.append('\n [*] ' + str(self.supF))
        df = pd.read_excel(self.supF, 'Sheet1', index_col=None)
        df.to_csv(self.supF.replace(self.supF[self.supF.rindex('.'):], '') + '.csv', sep=',', encoding='UTF-8')
        sup_xls = self.supF[self.supF.rindex('/'):]                            #/Supply To Be Deleted.xlsx
        sup_csv = sup_xls.replace(self.supF[self.supF.rindex('.'):], '') + '.csv' #/Supply To Be Deleted.csv
        self.sup_dest = self.dump + str(sup_csv.replace(sup_csv[sup_csv.rindex('/')], ''))   #self.dump + Supply To Be Deleted.csv
        shutil.move(self.supF.replace(self.supF[self.supF.rindex('.'):], '') + '.csv', self.sup_dest)

        self.del_col()

    def del_col(self):
        self.console.append('\nDeleting columns from Supply To Be Deleted headers...')
        with open(self.forF, 'r', encoding='UTF-8') as for_csv:
            df = pd.read_csv(for_csv, low_memory=False)
            with open(self.sup_dest, 'r', encoding='UTF-8') as supply_csv:
                df2 = pd.read_csv(supply_csv)
                for i in df2.columns:
                    try:
                        del df[i]
                    except KeyError:
                        pass
                self.draft = self.sup_dest.replace(self.sup_dest[self.sup_dest.rindex('\\'):], '') + '\\draft.csv'
                df.to_csv(self.draft, index=False)

        self.filter_rows()
      
    def filter_rows(self):
        #latest sheet(#FILTER 1,511 records)
        self.console.append('\nFiltering Rows...')
        with open(self.draft, 'r', encoding='UTF-8') as draft_csv:
            df = pd.read_csv(draft_csv, low_memory=False)
            df = df.loc[df['IG'].isin(['SFDC IPS', 'Oracle IPS', 'Workday IPS']) | df['Resources Reqd From'].isin(['Salesforce IPS', 'Oracle IPS', 'Workday IPS'])]
            output = df.drop('Technology', axis=1)
            output.to_csv(self.draft, index=False)

        self.vlookup()

    def vlookup(self):
        self.console.append('\nVLOOKUP Consolidated HC report & Consolidated for reporting sheets...')
        with open(self.draft, 'r', encoding='UTF-8') as draft_csv:
            df = pd.read_csv(draft_csv)
            with open(self.hc_dest, 'r', encoding='UTF-8') as hc_csv:
                df2 = pd.read_csv(hc_csv)
                df2 = df2[['Name', 'Technology']].drop_duplicates()
                output = df.merge(df2, on=['Name'], how='left')
                output.to_csv(self.draft, index=False)
        #take not null values or drop null values and 
        #replace all NULL to Non-Cloud in Technology Column
        with open(self.draft, 'r', encoding='UTF-8') as draft_csv:
            df = pd.read_csv(draft_csv)
            df['Technology'] = df['Technology'].fillna('Non-Cloud')
            df.to_csv(self.draft, index=False)

        self.save_output()

    def save_output(self):
            _saveFile = QtWidgets.QFileDialog()
            _saveFile.setWindowIcon(QtGui.QIcon('./img/icon.ico'))
            _saveFile.setWindowTitle('Save File As')
            self.fn, _ok = QtWidgets.QFileDialog.getSaveFileName(_saveFile, 'Save File As', 'Enter filename', 'Comma Delimited File (*.csv)')
            if _ok and self.fn != '':
                try:
                    with open(self.draft, 'r', encoding='UTF-8') as draft_csv:
                        df = pd.read_csv(draft_csv)
                        df.to_csv(self.fn, index=False)
                        self.console.append('\nSaving output file as :\n')
                        self.console.append(str(self.fn))
                    self.end()
                except PermissionError:
                    self.end()
            else:
                _errorMsg = QtWidgets.QMessageBox()
                _errorMsg.setWindowIcon(QtGui.QIcon('./img/icon.ico'))
                _errorMsg.setWindowTitle('Error !')
                _errorMsg.setIcon(QtWidgets.QMessageBox.Critical)
                _errorMsg.setText('Filename should not be blank.')
                _errorMsg.exec_()
                self.save_output()

    def end(self):
        try:
            shutil.rmtree(self.dump)
        except FileNotFoundError:
            pass

        self.console.append('\nPlease find and replace all characters "Ã±" to "ñ" manually...') #QMessage here choice
        _end = QtWidgets.QMessageBox()
        _end.setWindowIcon(QtGui.QIcon('./img/icon.ico'))
        _end.setWindowTitle('Finished')
        choice = QtWidgets.QMessageBox.question(_end, 'Finished',
                                            '\nPlease find and replace all characters "Ã±" to "ñ" manually...\nDo you want to open the output file now?',
                                            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if choice == QtWidgets.QMessageBox.Yes:
            self.console.append('Opening output file :')
            self.console.append('\n[*]' + str(self.fn))
            os.startfile(self.fn)

        clock_end = time.time() #Time after it finished
        self.console.append('\nD O N E !!! \n[ Finished in ' + str(self.clock_start-clock_end) + ' seconds ]')

if __name__ == '__main__':

    app = QtWidgets.QApplication([])
    SupplyTool = QtWidgets.QMainWindow()
    ui = Ui_SupplyTool()
    ui.setupUi(SupplyTool)
    SupplyTool.show()
    app.exec_()

