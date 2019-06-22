__author__ = 'Mark Mon Monteros'

from PyQt5 import QtCore, QtGui, QtWidgets
import os, shutil, time
import pandas as pd
import numpy as np

class Ui_DemandTool(object):

    clock_start = time.time() #Time before the operations start

    def __init__(self):
        self.path = os.path.dirname(os.path.realpath(__file__))
        self.dump = self.path + '\\dump'
        self.dcsopath = self.path + '\\dcso'
        self.atcp = None
        self.datal = None
        self.demandcol =  None
        self.dcso = list()
        self.dcso_edited = list()
        self.dcso_error = list()
        self.dcso_error_substr = list()
        self.dcsocsv = list()
        self.rhlyesdrop = list()
        self.rhlnodrop = list()
        self.du = list()
        self.ps = list()
        self.fe = list()
        self.rhlrrd = list()
        self.rrdskills = dict()
        self.fn = None

    def setupUi(self, DemandTool):       
        DemandTool.setObjectName("DemandTool")
        DemandTool.setFixedSize(391, 750)
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Light, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 127, 63))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Midlight, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 42, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Dark, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 56, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Mid, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.BrightText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Shadow, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.AlternateBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 220))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ToolTipBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ToolTipText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Light, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 127, 63))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Midlight, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 42, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Dark, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 56, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Mid, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.BrightText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Shadow, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.AlternateBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 220))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ToolTipBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ToolTipText, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 42, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 170, 127))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Light, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 127, 63))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Midlight, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 42, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Dark, brush)
        brush = QtGui.QBrush(QtGui.QColor(170, 56, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Mid, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 42, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.BrightText, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 42, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Window, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Shadow, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 85, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.AlternateBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 220))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ToolTipBase, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ToolTipText, brush)
        DemandTool.setPalette(palette)
        self.centralwidget = QtWidgets.QWidget(DemandTool)
        self.centralwidget.setObjectName("centralwidget")
        self.atcpLbl = QtWidgets.QLabel(self.centralwidget)
        self.atcpLbl.setGeometry(QtCore.QRect(10, 10, 371, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.atcpLbl.setFont(font)
        self.atcpLbl.setObjectName("atcpLbl")
        self.atcpBtn = QtWidgets.QPushButton(self.centralwidget)
        self.atcpBtn.setGeometry(QtCore.QRect(10, 61, 75, 23))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 42, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 42, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        self.atcpBtn.setPalette(palette)
        self.atcpBtn.setObjectName("atcpBtn")
        self.atcpBtn.setCheckable(True)
        self.atcpEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.atcpEdit.setGeometry(QtCore.QRect(10, 31, 371, 20))
        self.atcpEdit.setObjectName("atcpEdit")
        self.atcpEdit.setReadOnly(True)
        self.datalBtn = QtWidgets.QPushButton(self.centralwidget)
        self.datalBtn.setGeometry(QtCore.QRect(10, 140, 75, 23))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 42, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        self.datalBtn.setPalette(palette)
        self.datalBtn.setObjectName("datalBtn")
        self.datalBtn.setCheckable(True)
        self.datalLbl = QtWidgets.QLabel(self.centralwidget)
        self.datalLbl.setGeometry(QtCore.QRect(10, 89, 371, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.datalLbl.setFont(font)
        self.datalLbl.setObjectName("datalLbl")
        self.datalEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.datalEdit.setGeometry(QtCore.QRect(10, 110, 371, 20))
        self.datalEdit.setObjectName("datalEdit")
        self.datalEdit.setReadOnly(True)
        self.demandcolEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.demandcolEdit.setGeometry(QtCore.QRect(10, 190, 371, 20))
        self.demandcolEdit.setObjectName("demandcolEdit")
        self.demandcolEdit.setReadOnly(True)
        self.demandcolLbl = QtWidgets.QLabel(self.centralwidget)
        self.demandcolLbl.setGeometry(QtCore.QRect(10, 169, 371, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.demandcolLbl.setFont(font)
        self.demandcolLbl.setObjectName("demandcolLbl")
        self.demandcolBtn = QtWidgets.QPushButton(self.centralwidget)
        self.demandcolBtn.setGeometry(QtCore.QRect(10, 220, 75, 23))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 42, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        self.demandcolBtn.setPalette(palette)
        self.demandcolBtn.setObjectName("demandcolBtn")
        self.demandcolBtn.setCheckable(True)
        self.startBtn = QtWidgets.QPushButton(self.centralwidget)
        self.startBtn.setGeometry(QtCore.QRect(10, 650, 371, 51))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 42, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        self.startBtn.setPalette(palette)
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
        self.author.setGeometry(QtCore.QRect(190, 710, 201, 20))
        self.author.setObjectName("author")
        self.console = QtWidgets.QTextEdit(self.centralwidget)
        self.console.setGeometry(QtCore.QRect(0, 380, 391, 261))
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
        self.dcsoBtn = QtWidgets.QPushButton(self.centralwidget)
        self.dcsoBtn.setGeometry(QtCore.QRect(10, 350, 75, 23))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 85, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.ButtonText, brush)
        brush = QtGui.QBrush(QtGui.QColor(127, 42, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText, brush)
        self.dcsoBtn.setPalette(palette)
        self.dcsoBtn.setObjectName("dcsoBtn")
        self.dcsoBtn.setCheckable(True)
        self.dcsoLbl = QtWidgets.QLabel(self.centralwidget)
        self.dcsoLbl.setGeometry(QtCore.QRect(10, 250, 371, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.dcsoLbl.setFont(font)
        self.dcsoLbl.setObjectName("dcsoLbl")
        # self.dcsoEdit = QtWidgets.QLabel(self.centralwidget)
        # self.dcsoEdit.setEnabled(False)
        # self.dcsoEdit.setGeometry(QtCore.QRect(304, 350, 71, 20))
        font = QtGui.QFont()
        font.setKerning(True)
        # self.dcsoEdit.setFont(font)
        # self.dcsoEdit.setInputMethodHints(QtCore.Qt.ImhHiddenText)
        # self.dcsoEdit.setObjectName("label")
        self.dcsoWidget = QtWidgets.QTextEdit(self.centralwidget)
        #self.dcsoWidget = QtWidgets.QListWidget(self.centralwidget)
        self.dcsoWidget.setGeometry(QtCore.QRect(10, 270, 371, 71))
        self.dcsoWidget.setAcceptDrops(True)
        self.dcsoWidget.setInputMethodHints(QtCore.Qt.ImhMultiLine)
        self.dcsoWidget.setObjectName("dcsoWidget")
        self.dcsoWidget.setReadOnly(True)
        DemandTool.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(DemandTool)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 391, 26))
        self.menubar.setObjectName("menubar")
        DemandTool.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(DemandTool)
        self.statusbar.setObjectName("statusbar")
        DemandTool.setStatusBar(self.statusbar)

        self.retranslateUi(DemandTool)
        QtCore.QMetaObject.connectSlotsByName(DemandTool)

        self.console.append('\nINSTRUCTIONS :')
        self.console.append('\n[1] Input Excel file (.xls, .xlsx or .xlsm format to the field ATCP and Overdue Demand Report sheet. « FIRST FIELD »')
        self.console.append('\n[2] Input Excel file (.xls, .xlsx or .xlsm format to the field Data loader - Salesforce sheet. « SECOND FIELD »')
        self.console.append('\n[3] Input Excel file (.xls, .xlsx or .xlsm format to the field Demand Columns To Be Included sheet. « THIRD FIELD »')
        self.console.append('\n[4] Input DCSO Excel files (.xls, .xlsx or .xlsm format to the field DCSO sheets. « LAST FIELD »')
        self.console.append('\nNote: Demand Tool program will exit automatically if input sheets are incorrect.')

        self.startBtn.clicked.connect(self.start)
        self.atcpBtn.clicked.connect(self.set_xls)
        self.datalBtn.clicked.connect(self.set_xls)
        self.demandcolBtn.clicked.connect(self.set_xls)
        self.dcsoBtn.clicked.connect(self.set_xls)

    def retranslateUi(self, DemandTool):
        _translate = QtCore.QCoreApplication.translate
        DemandTool.setWindowTitle(_translate("DemandTool", "DemandTool"))
        DemandTool.setWindowIcon(QtGui.QIcon('./img/icon.ico'))
        self.atcpLbl.setText(_translate("DemandTool", "ATCP Open and Overdue Demand Report sheet :"))
        self.atcpBtn.setText(_translate("DemandTool", "Browse..."))
        self.datalBtn.setText(_translate("DemandTool", "Browse..."))
        self.datalLbl.setText(_translate("DemandTool", "Data loader - Salesforce sheet :"))
        self.demandcolLbl.setText(_translate("DemandTool", "Demand Columns To Be Included sheet :"))
        self.demandcolBtn.setText(_translate("DemandTool", "Browse..."))
        self.startBtn.setText(_translate("DemandTool", "S T A R T"))
        self.author.setText(_translate("DemandTool", "created by: Mark Mon Monteros"))
        self.console.setHtml(_translate("DemandTool", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:7.8pt; font-weight:400; font-style:normal;\">\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p></body></html>"))
        self.dcsoBtn.setText(_translate("DemandTool", "Browse..."))
        self.dcsoLbl.setText(_translate("DemandTool", "DCSO sheets :"))
        #self.dcsoWidget.setText(_translate("DemandTool", ""))

    def set_xls(self):
        if self.atcpBtn.isChecked():
            self.atcp = None
            self.atcp, _ = QtWidgets.QFileDialog.getOpenFileName(None, 'Select File', '', 'Excel Files (*.xls *.xlsx)')
            self.atcpEdit.setText(self.atcp)
            self.atcpBtn.toggle()
        if self.datalBtn.isChecked():
            self.datal = None
            self.datal, _ = QtWidgets.QFileDialog.getOpenFileName(None, 'Select File', '', 'Excel Files (*.xls *.xlsx)')
            self.datalEdit.setText(self.datal)
            self.datalBtn.toggle()
        if self.demandcolBtn.isChecked():
            self.demandcol = None
            self.demandcol, _ = QtWidgets.QFileDialog.getOpenFileName(None, 'Select File', '', 'Excel Files (*.xls *.xlsx)')
            self.demandcolEdit.setText(self.demandcol)
            self.demandcolBtn.toggle()
        if self.dcsoBtn.isChecked():
            del self.dcso[:]
            del self.dcsocsv[:]
            del self.dcso_edited[:]
            self.dcso, _ = QtWidgets.QFileDialog.getOpenFileNames(None, 'Select File/s', '', 'Excel Files (*.xls *.xlsx *.xlsm)')
            for f in self.dcso:
                self.dcsoWidget.append(os.path.basename(f))
            self.dcsoWidget.append('\nTotal DCSO files added: ' + str(len(self.dcso)))
            self.dcsoBtn.toggle()

    def start(self):
        self.console.setText('Initializing...')
        if self.atcp == '' or self.atcp is None or self.datal == '' or self.datal is None or self.demandcol == '' or self.demandcol is None or self.dcso == [] or len(self.dcso) == 0:
            _errorMsg = QtWidgets.QMessageBox()
            _errorMsg.setWindowIcon(QtGui.QIcon('./img/icon.ico'))
            _errorMsg.setWindowTitle('Warning !')
            _errorMsg.setIcon(QtWidgets.QMessageBox.Warning)
            _errorMsg.setText('All fields must contain an item.')
            _errorMsg.exec_()
            self._error_()
        else:
            self.run()

    def run(self):
        self.console.append('\n+-----------+-----------+')
        self.console.append(' Demand Automation Tool')
        self.console.append('   by Mark Mon Monteros')
        self.console.append('+-----------+-----------+')
        self.console.append(' Coded in Python ver 3.7*\n')

        self.convert()

    def _error_(self):
        self.console.append('\n>>> WARNING !!! <<<')
        self.console.append('\nPlease verify if you have a file inserted on each field.')

    def convert(self):
        self.console.append('\nConverting excel files to CSV UTF-8 format...')
        try:
            os.mkdir(self.dump)
            os.mkdir(self.dcsopath)
        except FileExistsError:
            shutil.rmtree(self.dump)
            shutil.rmtree(self.dcsopath)
            os.mkdir(self.dump)
            os.mkdir(self.dcsopath)
        try:
            os.remove(os.path.join(self.path, r'output.csv'))
            os.remove(os.path.join(self.path, r'final.csv'))
        except FileNotFoundError:
            pass

        self.console.append('\n [*] ' + str(os.path.basename(self.atcp)))
        rhl_yes = pd.read_excel(self.atcp, 'RHL=Yes', index_col=None)
        rhl_no = pd.read_excel(self.atcp, 'RHL=No', index_col=None)
        rhl_yes.to_csv(os.path.join(self.dump, r'RHL Yes.csv'), sep=',', index=False, encoding='UTF-8')
        rhl_no.to_csv(os.path.join(self.dump, r'RHL No.csv'), sep=',', index=False, encoding='UTF-8')

        self.console.append('\n [*] ' + str(os.path.basename(self.datal)))
        basename = os.path.basename(self.datal)
        data_loader = pd.read_excel(self.datal, 'Data loader - Salesforce', index_col=None)
        data_loader.to_csv(os.path.join(self.dump, basename.replace(basename[basename.rindex('.'):], '') + '.csv'), sep=',', encoding='UTF-8')

        self.console.append('\n [*] ' + str(os.path.basename(self.demandcol)))
        basename = os.path.basename(self.demandcol)
        demand_cols = pd.read_excel(self.demandcol, 'Sheet1', index_col=None)
        demand_cols.to_csv(os.path.join(self.dump, basename.replace(basename[basename.rindex('.'):], '') + '.csv'), sep=',', encoding='UTF-8')

        for f in self.dcso:
            basename = os.path.basename(f)
            self.console.append('\n [*] ' + str(basename))
            dcso_csv = pd.read_excel(f, 'Resource Requirements Detail', index_col=None)
            dcso_csv.to_csv(os.path.join(self.dcsopath, basename.replace(basename[basename.rindex('.'):], '') + '.csv'), sep=',', encoding='UTF-8')
        self.console.append('\nCSV UTF-8 Conversion Completed...')

        self.get_dcso_csv()

    def get_dcso_csv(self):
        for f in os.listdir(self.dcsopath):
            if f.endswith('.csv'):
                self.dcsocsv.append(f)

        self.cut_dcso_rows()

    def cut_dcso_rows(self):
        for dcso in self.dcsocsv:
            with open(os.path.join(self.dcsopath, dcso), 'r', encoding='UTF-8') as dcso_csv:
                df = pd.read_csv(dcso_csv, low_memory=False)
                df.drop(df.index[:28], inplace=True)
                df.to_csv(os.path.join(self.dcsopath, r'#edited - ' + dcso), header=None, index=False)

        for f in os.listdir(self.dcsopath):
            if f.startswith('#edited'):
                self.dcso_edited.append(f)
            else:
                os.remove(os.path.join(self.dcsopath, f))

        self.cut_dcso_columns()

    def cut_dcso_columns(self):
        self.console.append('\nDeleting Columns...')
        basename = os.path.basename(self.demandcol)
        with open(os.path.join(self.dump, basename.replace(basename[basename.rindex('.'):], '') + '.csv'), 'r', encoding='UTF-8') as del_colums:
            df = pd.read_csv(del_colums, low_memory=False)
            df.columns = df.columns.str.upper() 
        with open(os.path.join(self.dump, r'RHL Yes.csv'), 'r', encoding='UTF-8') as rhl_yes:
            df2 = pd.read_csv(rhl_yes, low_memory=False)
            df2.columns = df2.columns.str.upper()
        with open(os.path.join(self.dump, r'RHL No.csv'), 'r', encoding='UTF-8') as rhl_no:
            df3 = pd.read_csv(rhl_no, low_memory=False)
            df3.columns = df3.columns.str.upper()
        #drop RHL YES columns
        for col in list(df2.columns):
            if col not in list(df.columns) or not list(df3.columns):
                self.rhlyesdrop.append(col)
        #drop RHL NO columns
        for col in list(df3.columns):
            if col not in list(df.columns) or not list(df2.columns):
                self.rhlnodrop.append(col)

        df2.drop(self.rhlyesdrop, axis=1, inplace=True)
        df2.to_csv(os.path.join(self.dump, r'RHL-Y.csv'), index=False, encoding='UTF-8')

        df3.drop(self.rhlnodrop, axis=1, inplace=True)
        df3.to_csv(os.path.join(self.dump, r'RHL-N.csv'), index=False, encoding='UTF-8')

        self.combine_rhl()

    def combine_rhl(self):
        with open(os.path.join(self.dump, r'RHL-Y.csv'), 'r', encoding='UTF-8') as rhl_yes:
            df = pd.read_csv(rhl_yes, low_memory=False)
        with open(os.path.join(self.dump, r'RHL-N.csv'), 'r', encoding='UTF-8') as rhl_no:
            df2 = pd.read_csv(rhl_no, low_memory=False)
            output = pd.concat([df, pd.DataFrame(df2)], join='outer')
            output.to_csv(os.path.join(self.dump, r'RHL combined.csv'), index=False, encoding='UTF-8')

        for f in os.listdir(self.dump):
            if not f.endswith('combined.csv') and not f.startswith('Data loader'):
                os.remove(os.path.join(self.dump, f))

        self.filter_rows()

    def filter_rows(self):
        self.console.append('\nFiltering Rows...')
        with open(os.path.join(self.dump, r'Data loader - Salesforce.csv'), 'r', encoding='UTF-8') as data_loader:
            df = pd.read_csv(data_loader, low_memory=False)
            du = df['Delivery Unit/DU'].dropna()
            ps = df['Primary Skill'].dropna()
            fe = df['Fulfillment Entity'].dropna()

        for i in du:
            self.du.append(i)
        for i in ps:
            self.ps.append(i)
        for i in fe:
            self.fe.append(i)

        with open(os.path.join(self.dump, r'RHL combined.csv'), 'r', encoding='UTF-8') as rhl:
            df = pd.read_csv(rhl, low_memory=False)
            output = df.loc[df['DU'].isin(self.du) | df['PRIMARY SKILL'].isin(self.ps) | df['FULFILLMENT ENTITY'].isin(self.fe)]
            output.to_csv(os.path.join(self.dump, r'#edited - RHL combined.csv'), index=False, encoding='UTF-8')

        for f in os.listdir(self.dump):
            if not f.startswith('#edited'):
                os.remove(os.path.join(self.dump, f))

        self.get_rhl_rrd()

    def get_rhl_rrd(self):
        with open(os.path.join(self.dump, '#edited - RHL combined.csv'), 'r', encoding='UTF-8') as rhl:
            df = pd.read_csv(rhl, low_memory=False)
            df = df['RRD NUMBER']

            for i in df:
                self.rhlrrd.append(i)


        self.get_dcso_rrd_skills()

    def get_dcso_rrd_skills(self):
        for dcso in self.dcso_edited:
            try:
                with open(os.path.join(self.dcsopath, dcso), 'r', encoding='UTF-8') as dcso_csv:
                    df = pd.read_csv(dcso_csv, low_memory=False, index_col='RRD\nNumber')
            except ValueError:
                self.dcso_error.append(dcso.split('- ', 1)[1])

        for i in self.dcso_error:
            self.dcso_error_substr.append(i.rsplit('.', 1)[0])

        if not self.dcso_error_substr:
            pass
        else:
            self.error()
                    
        for i in self.rhlrrd:
            try:
                self.rrdskills[i] = df.loc[i, 'Skill']
            except KeyError:
                pass

        with open(os.path.join(self.dump, r'rrd_skills.csv'), 'w', encoding='UTF-8') as rrdskills_csv:
            for i in self.rrdskills.keys():
                rrdskills_csv.write("%s,%s\n"%(i,self.rrdskills[i]))

        with open(os.path.join(self.dump, r'rrd_skills.csv'), 'r', encoding='UTF-8') as rrdskills_csv:
            df = pd.read_csv(rrdskills_csv, low_memory=False, names=['RRD NUMBER', 'ADDITIONAL SKILLS (NICE TO HAVE THIS)'])
            df.to_csv(os.path.join(self.dump, r'#rrd_skills.csv'), index=False, encoding='UTF-8')

        self.add_skills_column()

    def add_skills_column(self):
        with open(os.path.join(self.dump, r'#edited - RHL combined.csv'), 'r', encoding='UTF-8') as rhl:
            df = pd.read_csv(rhl, low_memory=False)
            #FILL NULL VALUES TO COLUMN
            df['ADDITIONAL SKILLS (NICE TO HAVE THIS)'] = [np.nan for _ in range(len(df))]
            df.to_csv(os.path.join(self.dump, r'#final.csv'), index=False, encoding='UTF-8')

        self.append_skills()

    def append_skills(self):
        with open(os.path.join(self.dump, r'#final.csv'), 'r', encoding='UTF-8') as rhl:
            df = pd.read_csv(rhl, low_memory=False)
            df = df[['RRD NUMBER', 'ADDITIONAL SKILLS (NICE TO HAVE THIS)']]
        with open(os.path.join(self.dump, r'#rrd_skills.csv'), 'r', encoding='UTF-8') as rrdskills_csv:
            df2 = pd.read_csv(rrdskills_csv, low_memory=False)

        output = df.merge(df2, on=['RRD NUMBER'], how='left')
        output = output.drop('ADDITIONAL SKILLS (NICE TO HAVE THIS)_x', axis=1)
        output.rename(columns={'RRD NUMBER':'RRD NUMBER', 'ADDITIONAL SKILLS (NICE TO HAVE THIS)_y':'ADDITIONAL SKILLS (NICE TO HAVE THIS)'}, inplace=True)
        output.to_csv(os.path.join(self.dump, r'#map.csv'), index=False, encoding='UTF-8')

        with open(os.path.join(self.dump, r'#final.csv'), 'r', encoding='UTF-8') as rhl:
            df3 = pd.read_csv(rhl, low_memory=False)
        with open(os.path.join(self.dump, r'#map.csv'), 'r', encoding='UTF-8') as map_csv:
            df4 = pd.read_csv(map_csv, low_memory=False)
            df4 = df4['ADDITIONAL SKILLS (NICE TO HAVE THIS)']

        df3 = df3.drop('ADDITIONAL SKILLS (NICE TO HAVE THIS)', axis=1)
        output = pd.concat([df3, df4], axis=1)
        output.to_csv(os.path.join(self.path, r'output.csv'), index=False, encoding='UTF-8')

        self.case_time_format()

    def case_time_format(self):
        with open(os.path.join(self.path, r'output.csv'), 'r', encoding='UTF-8') as output_csv:
            df = pd.read_csv(output_csv, low_memory=False)
            df.columns = df.columns.str.title()

            changeUpper = ['Id', 'Sg', 'Du', 'Og', 'Dcso', 'Rrd', 'Dg', 'Sid', 'Sr', 
            'Ig', 'Drd', 'Rsd', 'Gu', 'Dcn', 'Gcp', 'Ou', 'Csg', 'It', 'Rhly']
            
            #ID, #SG, #DU, #OG, #DCSO, #RRD, #WBSe, #DG, #SID, #SR, #IG,
            #DRD, #RSD, #GU, #DCN, #GCP, #To, #Of, #OU, #CSG, #IT, #RHLY

            for i in changeUpper:
                df.rename(columns=lambda x: x.replace(i, i.upper()), inplace=True)
            df.rename(columns=lambda x: x.replace('Wbse', 'WBSe'), inplace=True)
            df.rename(columns=lambda x: x.replace('To', 'to'), inplace=True)
            df.rename(columns=lambda x: x.replace('Of', 'of'), inplace=True)
    
            for i in df.columns:
                if not i.startswith('No of Times') and not i.startswith('Abacus') and 'Date' in i or 'Created' in i or i.endswith('Changed On'):
                    df[i] = pd.to_datetime(df[i])
                    df[i] = df[i].dt.strftime('%m/%d/%Y')

            df = df.replace('NaT', np.nan, regex=True)
            df.to_csv(os.path.join(self.path, r'final.csv'), index=False, encoding='UTF-8')
            
        os.remove(os.path.join(self.path, r'output.csv'))

        self.save_output()

    def save_output(self):
        _saveFile = QtWidgets.QFileDialog()
        _saveFile.setWindowIcon(QtGui.QIcon('./img/icon.ico'))
        _saveFile.setWindowTitle('Save File As')
        self.fn, _ok = QtWidgets.QFileDialog.getSaveFileName(_saveFile, 'Save File As', 'Enter filename', 'Comma Delimited File (*.csv)')
        if _ok and self.fn != '':
            try:
                with open(os.path.join(self.path, r'final.csv'), 'r', encoding='UTF-8') as draft_csv:
                    df = pd.read_csv(draft_csv)
                    df.to_csv(self.fn, index=False)
                    self.console.append('\nSaving output file as :\n')
                    self.console.append(str(self.fn))
                os.remove(os.path.join(self.path, r'final.csv'))
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
            shutil.rmtree(self.dcsopath)
        except FileNotFoundError:
            pass

        _end = QtWidgets.QMessageBox()
        _end.setWindowIcon(QtGui.QIcon('./img/icon.ico'))
        _end.setWindowTitle('Finished')
        choice = QtWidgets.QMessageBox.question(_end, 'Finished',
                                            '\nDo you want to open the output file now?',
                                            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if choice == QtWidgets.QMessageBox.Yes:
            self.console.append('Opening output file :')
            self.console.append('\n[*]' + str(self.fn))
            os.startfile(self.fn)

        clock_end = time.time() #Time after it finished
        self.console.append('\nD O N E !!! \n[ Finished in ' + str(self.clock_start-clock_end) + ' seconds ]')

    def error(self):
        self.console.setText('\nERROR: Please check/review "Resource Requirement Detail" column for these DCSO files..')
        for i in self.dcso_error_substr:
            self.console.append('   [*]', i)
        self.console.append('     -- Column header should start at row #30 or DCSO sheet template is mismatch  --\n')
        shutil.rmtree(self.dump)
        shutil.rmtree(self.dcsopath)
        sys.exit()

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    DemandTool = QtWidgets.QMainWindow()
    ui = Ui_DemandTool()
    ui.setupUi(DemandTool)
    DemandTool.show()
    app.exec()