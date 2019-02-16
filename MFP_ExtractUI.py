# UI Created by: PyQt5 UI code generator 5.5.1
#
# Adam T. Cuellar
# My 'Hello World' program for python :)

import myfitnesspal
import re
from openpyxl import load_workbook
import easygui
from easygui import multpasswordbox
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QWidget, QProgressBar, QPushButton, QApplication)
from datetime import date
from dateutil.rrule import rrule, DAILY

# some global variables
startText = ""
finishText = ""
username = ""
password = ""

# create login GUI
def loginGUI():
        global username
        global password
        msg = "Enter MyFitnessPal Login Information"
        title = "Log In"
        fieldNames = ["Username", "Password"]
        fieldValues = []
        fieldValues = multpasswordbox(msg,title, fieldNames)

        # make sure each field gets an input
        while 1:
            if fieldValues is None:
                break
            errmsg = ""
            for i in range(len(fieldNames)):
              if fieldValues[i].strip() == "":
                errmsg = errmsg + ('"%s" is a required field.\n\n' % fieldNames[i])
            if errmsg == "": break # no problems found
            fieldValues = multpasswordbox(errmsg, title, fieldNames, fieldValues)

        username = fieldValues[0]
        password = fieldValues[1]


# adds information to excel sheet
def parseForExcel(dictionary, line, sh, date):

    # get each part of the dictionary
    carbs = dictionary.get("carbohydrates")
    pro = dictionary.get("protein")
    fat = dictionary.get("fat")
    cals = dictionary.get("calories")
    sodium = dictionary.get("sodium")
    sugar = dictionary.get("sugar")
    potassium = dictionary.get("potassium")

    # write info to respective column + row on sheet
    sh.cell(line,1).value = date
    sh.cell(line,2).value = pro
    sh.cell(line,3).value = carbs
    sh.cell(line,4).value = fat
    # row reserved for fiber, currently not working
    sh.cell(line,5).value = 0;
    sh.cell(line,6).value = cals
    sh.cell(line,7).value = sodium
    # riw reserved for potassium, currently not working
    sh.cell(line,8).value = sodium

# extracts proper info from mfp
def extractInfo(StartDate,EndDate, fileName,self):

    completion = 0

    # open existing excel file and retrieve first sheet
    book = load_workbook(fileName)
    macros_sheet = book['Macros']

    # get last written in line number (skips 2 to leave extra space for in between weeks)
    line = macros_sheet.max_row + 2

    # open login UI
    loginGUI()

    client = None
    loginAttempts = 0;

    # make sure the user logs in correctly, if not give them 3 attempts
    while client is None:
        if loginAttempts is 3:
            easygui.msgbox("You've exceeded the maximum Log In attempts", "Warning")
            return

        try:
            # assign account name to client and track for invalid credentials
            client = myfitnesspal.Client(username, password)
            self.progress.setValue(completion)
        except:
            loginAttempts += 1
            easygui.msgbox("Invalid Credentials. Please Try Again", "Warning")
            loginGUI()

    # split inputted dates
    StartDateArr = StartDate.split("/")
    EndDateArr = EndDate.split("/")

    # form date from input
    a = date(int(StartDateArr[2]),int(StartDateArr[0]),int(StartDateArr[1]))
    b = date(int(EndDateArr[2]),int(EndDateArr[0]),int(EndDateArr[1]))

    # calucate total number of days for progress bar
    numDays = (b-a).days

    # iterate through each day
    for dt in rrule(DAILY, dtstart = a, until = b):
        year = int(dt.strftime("%Y"))
        month = int(dt.strftime("%m"))
        day = int(dt.strftime("%d"))
        dictionary = client.get_date(year, month, day).totals
        line += 1
        parseForExcel(dictionary,line, macros_sheet, dt.strftime("%-m/%-d/%Y"))
        completion += 100/numDays
        self.progress.setValue(completion)

    book.save(fileName)

    # show user we've completed the task
    easygui.msgbox("Task Completed!", "Success")
    # reset progress bar
    self.progress.setValue(0)
    return

# GUI setup
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(612, 409)
        self.centralWidget = QtWidgets.QWidget(MainWindow)
        self.centralWidget.setObjectName("centralWidget")
        self.calendarWidget = QtWidgets.QCalendarWidget(self.centralWidget)
        self.calendarWidget.setGeometry(QtCore.QRect(20, 20, 221, 200))
        self.progress = QProgressBar(self.centralWidget)
        self.progress.setGeometry(QtCore.QRect(18, 255, 221, 20))
        self.calendarWidget.setObjectName("calendarWidget")
        self.pushButton = QtWidgets.QPushButton(self.centralWidget)
        self.pushButton.setGeometry(QtCore.QRect(400, 310, 131, 27))
        self.pushButton.setObjectName("pushButton")
        self.lineEdit = QtWidgets.QLineEdit(self.centralWidget)
        self.lineEdit.setGeometry(QtCore.QRect(470, 40, 113, 27))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralWidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(470, 80, 113, 27))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.label = QtWidgets.QLabel(self.centralWidget)
        self.label.setGeometry(QtCore.QRect(300, 40, 171, 20))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralWidget)
        self.label_2.setGeometry(QtCore.QRect(320, 80, 151, 20))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralWidget)
        self.label_3.setObjectName("label_3")
        self.label_3.setGeometry(QtCore.QRect(18, 235, 170, 20))
        MainWindow.setCentralWidget(self.centralWidget)
        self.mainToolBar = QtWidgets.QToolBar(MainWindow)
        self.mainToolBar.setObjectName("mainToolBar")
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.mainToolBar)
        self.statusBar = QtWidgets.QStatusBar(MainWindow)
        self.statusBar.setObjectName("statusBar")
        MainWindow.setStatusBar(self.statusBar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MFP Extractor"))
        self.pushButton.setText(_translate("MainWindow", "Edit Spreadsheet"))
        self.label.setText(_translate("MainWindow", "Starting Date (MM/DD/YYYY): "))
        self.label_2.setText(_translate("MainWindow", "End Date (MM/DD/YYYY): "))
        self.label_3.setText(_translate("MainWindow", "Progress:"))
        self.pushButton.clicked.connect(self.clickedOpenSpreadsheet)


    # read and redirect inputted dates
    def clickedOpenSpreadsheet(self):
        global startText
        global finishText
        startText = self.lineEdit.text()
        finishText = self.lineEdit_2.text()
        x = re.search("[0-9]{2}/[0-9]{2}/[0-9]{4}",startText);
        y = re.search("[0-9]{2}/[0-9]{2}/[0-9]{4}",finishText);

        # warn user if date format is incorrect, else open find file dialog
        if x is None or y is None:
            easygui.msgbox("Incorrect Date Format", "Warning")
        else:
            fileName = easygui.fileopenbox()
            extractInfo(startText,finishText,fileName,self)

        return


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())


