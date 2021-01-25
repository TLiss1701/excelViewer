import pandas
import sys
from PyQt5.QtWidgets import QApplication, QGridLayout, QVBoxLayout, QHBoxLayout, QPushButton, QWidget, QLabel, QLineEdit, QFileDialog, QStackedLayout, QCompleter
from PyQt5 import QtCore

class Window2(QWidget):
    def __init__(self, fileName):
        super().__init__()
        self.df = pandas.read_excel(fileName)

        self.setWindowTitle("Excel Viewer")
        #Second Page Layout
        self.layout = QVBoxLayout()
        #Pick What Column to Filter Off Of
        self.columnPicker = QHBoxLayout()
        self.columnPicker.addWidget(QLabel("Column:"))
        self.columns = list(self.df.columns)
        self.colComp = QCompleter(self.columns)
        self.selectColumn = QLineEdit()
        self.selectColumn.setCompleter(self.colComp)
        self.columnPicker.addWidget(self.selectColumn)
        self.colButton = QPushButton("Set Column")
        self.columnPicker.addWidget(self.colButton)
        self.colButton.clicked.connect(self.setColumn)
        self.layout.addLayout(self.columnPicker)
        #Pick What Row in Column to Show Data Off Of
        self.indexPicker = QHBoxLayout()
        self.indexPicker.addWidget(QLabel("Element:"))
        self.idxComp = QCompleter([])
        self.selectIndex = QLineEdit()
        self.selectIndex.setCompleter(self.idxComp)
        self.indexPicker.addWidget(self.selectIndex)
        self.idxButton = QPushButton("Set Index")
        self.indexPicker.addWidget(self.idxButton)
        self.idxButton.clicked.connect(self.setIndex)
        self.layout.addLayout(self.indexPicker)
        #Show all Data
        self.dataTable = QGridLayout()
        for x in range(len(self.df.columns)-1):
            self.dataTable.addWidget(QLabel("Row Title " + str(x)), x, 1)
            self.dataTable.addWidget(QLabel("Row Data " + str(x)), x, 2)
        self.layout.addLayout(self.dataTable)
        self.setLayout(self.layout)

    def setColumn(self):
        print(list(self.df[self.selectColumn.text()]))
        self.indicies = list(self.df[self.selectColumn.text()])
        self.idxComp = QCompleter(self.indicies)
        self.selectIndex.setCompleter(self.idxComp)


    def setIndex(self):
        print(self.selectColumn.text())
        self.dispdf = self.df.drop(columns=[self.selectColumn.text()])
        print(self.dispdf)

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
            self.w.show()

        else:
            self.w.close()  # Close window.
            self.w = None  # Discard reference.


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window2(r"C:\Users\TLiss\Desktop\test_data.xlsx") #Remove Later
    window.show()
    sys.exit(app.exec_())
