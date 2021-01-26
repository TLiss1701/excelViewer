import pandas
import sys
#pylint: disable=no-name-in-module
from PyQt5.QtWidgets import QApplication, QGridLayout, QVBoxLayout, QHBoxLayout, QPushButton, QWidget, QLabel, QLineEdit, QFileDialog, QStackedLayout, QCompleter, QComboBox, QErrorMessage, QMessageBox
from PyQt5 import QtCore
#For Export
#from fbs_runtime.application_context.PyQt5 import ApplicationContext

class Window2(QWidget):
    def __init__(self, fileName):
        super().__init__()
        self.df = pandas.read_excel(fileName)
        self.df = self.df.applymap(str)

        self.setWindowTitle("Excel Viewer")
        #Second Page Layout
        self.layout = QVBoxLayout()
        #Pick What Column to Filter Off Of
        self.columnPicker = QHBoxLayout()
        self.columnPicker.addWidget(QLabel("Column:"))
        self.columns = list(self.df.columns)
        self.selectColumn = QComboBox()
        self.selectColumn.setEditable(True)
        self.selectColumn.addItems(self.columns)
        self.oldSelectColumnText = self.selectColumn.currentText()
        self.columnPicker.addWidget(self.selectColumn)
        self.columnInit = True
        self.selectColumn
        self.selectColumn.setEditText("")
        self.selectColumn.activated.connect(self.setColumn)
        self.layout.addLayout(self.columnPicker)
        #Pick What Row in Column to Show Data Off Of
        self.indexPicker = QHBoxLayout()
        self.indexPicker.addWidget(QLabel("Element:"))
        self.selectIndex = QComboBox()
        self.selectIndex.setEditable(True)
        self.indexPicker.addWidget(self.selectIndex)
        self.selectIndex.activated.connect(self.setIndex)
        self.layout.addLayout(self.indexPicker)
        #Show all Data
        self.dataTable = QGridLayout()
        self.rowTitles = []
        self.rowData = []
        for x in range(len(self.df.columns)-1):
            self.rowTitles.append(QLabel(""))
            self.rowData.append(QLabel(""))
        for x in range(len(self.df.columns)-1):
            self.dataTable.addWidget(self.rowTitles[x], x, 1)
            self.dataTable.addWidget(self.rowData[x], x, 2)
        self.layout.addLayout(self.dataTable)
        self.setLayout(self.layout)

    def setColumn(self):
        if self.selectColumn.currentText() == self.oldSelectColumnText and not self.columnInit:
            return
        else:
            if self.columnInit:
                self.columnInit = False
            if self.selectColumn.currentText() in self.columns:
                if any(list(self.df[self.selectColumn.currentText()]).count(x) > 1 for x in list(self.df[self.selectColumn.currentText()])):
                    self.dupeWarning = QMessageBox()
                    self.dupeWarning.setIcon(QMessageBox.Warning)
                    self.dupeWarning.setWindowTitle("Duplicate Warning")
                    self.dupeWarning.setText('<h4>The Chosen Column contains duplicates, which cannot be differentiated when rows are picked.</h4><h3>Do you wish to continue?</h3>')
                    self.dupeWarning.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                    if(self.dupeWarning.exec_() == QMessageBox.Yes):
                        self.indicies = list(self.df[self.selectColumn.currentText()])
                        self.selectIndex.clear()
                        self.selectIndex.addItems(self.indicies)
                        self.oldSelectColumnText = self.selectColumn.currentText()
                        self.selectIndex.setEditText("")
                    else:
                        self.selectColumn.setEditText(self.oldSelectColumnText)
                        return
                else:
                    self.indicies = list(self.df[self.selectColumn.currentText()])
                    self.selectIndex.clear()
                    self.selectIndex.addItems(self.indicies)
                    self.oldSelectColumnText = self.selectColumn.currentText()
                    self.selectIndex.setEditText("")
            else:
                self.columnError = QMessageBox()
                self.columnError.setIcon(QMessageBox.Critical)
                self.columnError.setWindowTitle("Column Select Error")
                self.columnError.setText('<h3>The Chosen Column does not Exist.</h3>')
                self.columnError.setStandardButtons(QMessageBox.Ok)
                self.columnError.exec_()

    def setIndex(self):
        if self.selectColumn.currentText() != "" and self.selectIndex.currentText() in self.indicies:
            try:
                self.dispdf = self.df.loc[self.df[self.selectColumn.currentText()] == self.selectIndex.currentText()]
                self.dispdf = self.dispdf.drop(columns=[self.selectColumn.currentText()])
                for x in range(len(self.df.columns)-1):
                    self.rowTitles[x].setText(self.dispdf.columns[x] + ":")
                    self.rowData[x].setText(self.dispdf.iat[0, x])
            except:
                self.uncaughtError()
        else:
            self.indexError = QMessageBox()
            self.indexError.setIcon(QMessageBox.Critical)
            self.indexError.setWindowTitle("Row Select Error")
            self.indexError.setText('<h3>The Chosen Row does not Exist.</h3>')
            self.indexError.setStandardButtons(QMessageBox.Ok)
            self.indexError.exec_()

    def uncaughtError(self):
        self.generalError = QMessageBox()
        self.generalError.setIcon(QMessageBox.Critical)
        self.generalError.setWindowTitle("Unexpected Error")
        self.generalError.setText('<h3>An Error Occured.</h3><h4>Please Try Again.</h4>')
        self.generalError.setStandardButtons(QMessageBox.Ok)
        self.generalError.exec_()

class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Viewer")
        #Stacked Layout
        self.stackedLayout = QStackedLayout()
        self.w = None  # No external window yet.
        #First Page Layout
        self.page1 = QWidget()
        self.page1Layout = QVBoxLayout()
        self.page1Layout.addWidget(QLabel("<h1>Excel Viewer</h1>"))
        self.page1Layout.addWidget(QLabel("Please choose the File to View (or drag and drop onto the box):"))
        self.fileSelect = QHBoxLayout()
        self.fileSelect.addWidget(QLabel("File:"))
        self.fileLineEdit = QLineEdit()
        self.fileSelect.addWidget(self.fileLineEdit)
        self.fileButton = QPushButton("Browse")
        self.fileName = ""
        self.fileButton.clicked.connect(self.getFiles)
        self.fileSelect.addWidget(self.fileButton)
        self.page1Layout.addLayout(self.fileSelect)
        self.nextButton = QPushButton("Next")
        self.nextButton.clicked.connect(self.switchPage)
        self.page1Layout.addWidget(self.nextButton)
        self.page1.setLayout(self.page1Layout)
        self.stackedLayout.addWidget(self.page1)
        #Set Object Layout
        self.setLayout(self.stackedLayout)

    def getFiles(self):
        self.fileName, _ = QFileDialog.getOpenFileName(self, 'Single File', QtCore.QDir.rootPath() , '*.xlsx')
        self.fileLineEdit.setText(self.fileName)

    def switchPage(self):
        if self.fileName != "" and self.fileName[-5:] == ".xlsx":
            if self.w is None:
                self.w = Window2(self.fileName)
                self.hide()
                self.w.show()
            else:
                self.w.close()  # Close window.
                self.w = None  # Discard reference.
        else:
            self.fileError = QMessageBox()
            self.fileError.setIcon(QMessageBox.Critical)
            self.fileError.setWindowTitle("File Error")
            self.fileError.setText('<h3>Your selected file is not an Excel File.</h3><h4> Please ensure the selected file ends with ".xlsx" and try again.</h4>')
            self.fileError.setStandardButtons(QMessageBox.Ok)
            self.fileError.exec_()


if __name__ == "__main__":
    #For Export
    #appctxt = ApplicationContext()
    #For Testing
    app = QApplication([])
    window = Window()
    window.show()
    #For Export
    #sys.exit(appctxt.app.exec_())
    #For Testing
    sys.exit(app.exec_())
