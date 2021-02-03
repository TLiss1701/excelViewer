# Excel Viewer

Excel Viewer is a small local aplication that displays a selected row of data from a given excel spreadsheet. The inspiration behind the application was the need to view very wide excel sheets (ie: Sheets of Client Information) on a small portion of the screen. With this application, a filter column can be set (ie: Client Name), and then the row can be selected (ie: Doe, John) and all of that information will then appear on a single screen for easy readability. In addition, the rows and columns selected can be changed to view the spreadsheet in any number of ways.

### Installing

In order to Install Excel Viewer on your local device, download <a id="raw-url" href="https://github.com/TLiss1701/excelViewer/raw/master/Excel%20ViewerSetup.exe">this file</a>. Then, just proceed with the standard installation process.

## Getting Started with Development

The entire program is built in ```PyQt5``` and is stored in one python file, ```main.py```. 
In order to get the python file running out of the context of the fbs project structure, comment out all lines directly below ```#For Export``` and uncomment all lines directly below ```#For Testing```.

### Packaging into a .exe and installer

The fbs project structure is used to package the pythong file into an executable. Follow <a id="raw-url" href="https://build-system.fman.io/pyqt-exe-creation/">this tutorial</a> in order to package the file into an executable. The packaged latest version along with an installer is included in the repo.

## Built With

* [PyQt5](https://doc.qt.io/qtforpython/) - The application framework used

## Authors

* **Trevor Liss**
