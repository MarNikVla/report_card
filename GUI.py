import sys

from PyQt5 import QtCore, QtWidgets
from PyQt5.QtCore import Qt

from main import save_file


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(674, 238)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.choiceButton = QtWidgets.QPushButton(self.centralwidget)
        self.choiceButton.setGeometry(QtCore.QRect(50, 110, 171, 40))
        self.choiceButton.setObjectName("choiceButton")
        self.choiceButton.clicked.connect(self.open_file)

        self.fillButton = QtWidgets.QPushButton(self.centralwidget)
        self.fillButton.setGeometry(QtCore.QRect(470, 110, 171, 40))
        self.fillButton.setObjectName("fillButton")
        self.fillButton.setDisabled(True)
        self.fillButton.clicked.connect(self.fill_file)

        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(20, 30, 621, 50))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit.setReadOnly(True)

        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(50, 160, 590, 40))
        self.label.setObjectName("label")
        self.label.setAlignment(Qt.AlignRight)
        self.label.setHidden(True)
        self.label.setWordWrap(True)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 674, 36))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Заполнение табелей"))
        self.choiceButton.setText(_translate("MainWindow", "Выберите файл"))
        self.fillButton.setText(_translate("MainWindow", "Заполнить"))
        self.label.setText(_translate("MainWindow", "TextLabel"))

    def open_file(self):
        self.file_name = QtWidgets.QFileDialog.getOpenFileName(None, "Open", "",
                                                               "Excel (*.xls *.xlsx)")
        if self.file_name[0] != '':
            self.lineEdit.setText(self.file_name[0])
            self.fillButton.setDisabled(False)

    def fill_file(self):
        try:
            if self.file_name[0] != '':
                save_file(self.file_name[0])
                self.label.setHidden(False)
                self.label.setStyleSheet("color: green")
                self.label.setText('Done! Окно можно закрыть')
        except Exception as exc:
            self.label.setHidden(False)
            self.label.setStyleSheet(
                "background-color: yellow; color: red; border: 1px solid black;")
            if exc.__class__.__name__ == 'PermissionError':
                self.label.setText(
                    f'Файл используется другой программой Ошибка:{exc.__class__.__name__}: {exc}')
            else:
                self.label.setText(f'Ошибка:{exc.__class__.__name__}: {exc}')


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
