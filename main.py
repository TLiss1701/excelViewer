import pandas
import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QGridLayout
from PyQt5.QtWidgets import QVBoxLayout
from PyQt5.QtWidgets import QHBoxLayout
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QLabel
from PyQt5.QtWidgets import QLineEdit
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QStackedLayout
from PyQt5 import QtCore

class Window2(QWidget):
    def __init__(self, fileName):
        super().__init__()

        df = pandas.read_excel(fileName)
        print(df)

        self.setWindowTitle("Excel Viewer")
        #Second Page Layout
        self.page2Layout = QVBoxLayout()
        #TODO - Make 2nd Page#
        self.fileLabel = QLabel("Test" + fileName)
        self.page2Layout.addWidget(self.fileLabel)
        ###
        self.setLayout(self.page2Layout)

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
