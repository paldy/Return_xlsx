import os
import sys
from PyQt5 import QtCore, QtGui, QtWidgets

from handler import ThreadHandler  # Importing custom ThreadHandler class
from view import *  # Importing custom UI definition


class Interface(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_MainWindow()  # Creating an instance of the UI
        self.ui.setupUi(self)  # Setting up the UI elements

        self.excel_file = None  # Variable to store selected Excel file

        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)  # Setting window flags to make it frameless
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)  # Setting attribute for translucent background
        self.center()  # Centering the window on the screen

        # Connecting buttons to their respective functions
        self.ui.pushButton_3.clicked.connect(self.close)
        self.ui.pushButton.clicked.connect(self.choose_file)
        self.ui.pushButton_2.clicked.connect(self.start_process)

        self.handler = ThreadHandler()  # Creating an instance of ThreadHandler
        self.handler.signal.connect(self.signal_handler)  # Connecting signal from ThreadHandler to signal_handler method


    def center(self):
        qr = self.frameGeometry()
        cp = QtWidgets.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())  # Moving the window to the centered position


    def mousePressEvent(self, event):
        self.oldPos = event.globalPos()


    def mouseMoveEvent(self, event):
        try:
            delta = QtCore.QPoint(event.globalPos() - self.oldPos)
            self.move(self.x() + delta.x(), self.y() + delta.y())
            self.oldPos = event.globalPos()
        except AttributeError:
            pass  # Handling an attribute error if it occurs


    def choose_file(self):
        file = QtWidgets.QFileDialog.getOpenFileName(self)[0]
        if file:
            self.excel_file = file
            self.ui.label_4.setText(os.path.basename(self.excel_file))  # Displaying the selected file name


    def start_process(self):
        conf_list = [
            self.ui.checkBox.isChecked(),
            self.ui.checkBox_2.isChecked(),
            self.ui.checkBox_3.isChecked(),
        ]  # Creating a list of configuration options based on checkbox states

        if self.excel_file:  # Checking if a file is selected
            self.handler.filepath = self.excel_file
            self.handler.config = conf_list

            if any(conf_list):  # Checking if any configuration option is selected
                self.handler.start()  # Starting the thread handler
                self.ui.pushButton.setDisabled(True)  # Disabling buttons during the process
                self.ui.pushButton_2.setDisabled(True)
            else:
                message = "Need to configure"  # Error message if no configuration option is selected
                QtWidgets.QMessageBox.warning(self, "Error", message)
        else:
            message = "You should select a file"  # Error message if no file is selected
            QtWidgets.QMessageBox.warning(self, "Error", message)


    def signal_handler(self, value):
        if value[0] == "result":  # Handling the result signal
            self.ui.label_3.setText(value[1])  # Displaying the result
            message = f"Password recovered: {value[1]}\n"
            message += "Created new file without password: decrypted.xlsx"
            QtWidgets.QMessageBox.about(self, "result", message)  # Showing a message box with the result
            self.ui.pushButton.setDisabled(False)  # Re-enabling buttons after the process is completed
            self.ui.pushButton_2.setDisabled(False)
        elif value[0] == "fail":  # Handling the failure signal
            self.ui.label_3.setText(value[1])  # Displaying the failure message


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = Interface()
    window.show()
    sys.exit(app.exec_())

