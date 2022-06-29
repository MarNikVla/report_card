

import sys
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QLineEdit, QFileDialog


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName('MainWindow')
        MainWindow.resize(235, 180)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName('centralwidget')
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setObjectName('pushButton')
        self.line1 = QLineEdit(self.centralwidget)
        self.line1.resize(200, 32)
        self.line1.move(0, 50)
        self.path = None
        MainWindow.setCentralWidget(self.centralwidget)
        self.retranslateUi(MainWindow)

    def open_dialog_box(self):
        filename = QFileDialog.getOpenFileName()
        self.path = filename[0]
        self.line1.setText(self.path)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate('MainWindow', 'MainWindow'))
        self.pushButton.setText(_translate('MainWindow', 'button'))
        # self.pushButton.clicked.connect(self.open_dialog_box)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

