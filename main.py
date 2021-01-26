import pandas
import sys
#pylint: disable=no-name-in-module
from PyQt5.QtWidgets import QApplication, QGridLayout, QVBoxLayout, QHBoxLayout, QPushButton, QWidget, QLabel, QLineEdit, QFileDialog, QStackedLayout, QCompleter, QComboBox
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
        self.columnPicker.addWidget(self.selectColumn)
        self.selectColumn.activated.connect(self.setColumn)
        self.layout.addLayout(self.columnPicker)
        #Pick What Row in Column to Show Data Off Of
        self.indexPicker = QHBoxLayout()
        self.indexPicker.addWidget(QLabel("Element:"))
        self.selectIndex = QComboBox()
        self.indexPicker.addWidget(self.selectIndex)
        self.selectIndex.activated.connect(self.setIndex)
        self.layout.addLayout(self.indexPicker)
        #Show all Data
        self.dataTable = QGridLayout()
        self.rowTitles = []
        self.rowData = []
        for x in range(len(self.df.columns)-1):
            self.rowTitles.append(QLabel("Row Title " + str(x)))
            self.rowData.append(QLabel("Row Data " + str(x)))
        for x in range(len(self.df.columns)-1):
            self.dataTable.addWidget(self.rowTitles[x], x, 1)
            self.dataTable.addWidget(self.rowData[x], x, 2)
        self.layout.addLayout(self.dataTable)
        self.setLayout(self.layout)

    def setColumn(self):
        self.indicies = list(self.df[self.selectColumn.currentText()])
        self.selectIndex.clear()
        self.selectIndex.addItems(self.indicies)

    def setIndex(self):
        self.dispdf = self.df.loc[self.df[self.selectColumn.currentText()] == self.selectIndex.currentText()]
        self.dispdf = self.dispdf.drop(columns=[self.selectColumn.currentText()])
        for x in range(len(self.df.columns)-1):
            self.rowTitles[x].setText(self.dispdf.columns[x] + ":")
            self.rowData[x].setText(self.dispdf.iat[0, x])

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
        self.fileButton.clicked.connect(lambda: self.getFiles())
        self.fileSelect.addWidget(self.fileButton)
        self.page1Layout.addLayout(self.fileSelect)
        self.nextButton = QPushButton("Next")
        self.nextButton.clicked.connect(lambda: self.switchPage())
        self.page1Layout.addWidget(self.nextButton)
        self.page1.setLayout(self.page1Layout)
        self.stackedLayout.addWidget(self.page1)
        #Set Object Layout
        self.setLayout(self.stackedLayout)

    def getFiles(self):
        self.fileName, _ = QFileDialog.getOpenFileName(self, 'Single File', QtCore.QDir.rootPath() , '*.xlsx')
        self.fileLineEdit.setText(self.fileName)

    def switchPage(self):
        if self.w is None:
            self.w = Window2(self.fileName)
            self.hide()
            self.w.show()

        else:
            self.w.close()  # Close window.
            self.w = None  # Discard reference.


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
