# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Email_Handling_window.ui'
#
# Created by: PyQt5 UI code generator 5.15.10
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(391, 272)
        MainWindow.setMinimumSize(QtCore.QSize(391, 272))
        MainWindow.setMaximumSize(QtCore.QSize(391, 272))
        MainWindow.setAutoFillBackground(False)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.start_b = QtWidgets.QPushButton(self.centralwidget)
        self.start_b.setGeometry(QtCore.QRect(50, 180, 93, 28))
        self.start_b.setObjectName("start_b")
        self.Pause_b = QtWidgets.QPushButton(self.centralwidget)
        self.Pause_b.setGeometry(QtCore.QRect(260, 180, 93, 28))
        self.Pause_b.setObjectName("Pause_b")
        self.groupBox_4 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_4.setGeometry(QtCore.QRect(20, 0, 361, 80))
        self.groupBox_4.setTitle("")
        self.groupBox_4.setObjectName("groupBox_4")
        self.Status_label = QtWidgets.QLabel(self.groupBox_4)
        self.Status_label.setGeometry(QtCore.QRect(50, 20, 261, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.Status_label.setFont(font)
        self.Status_label.setFocusPolicy(QtCore.Qt.NoFocus)
        self.Status_label.setText("")
        self.Status_label.setAlignment(QtCore.Qt.AlignCenter)
        self.Status_label.setObjectName("Status_label")
        self.groupBox_5 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_5.setGeometry(QtCore.QRect(20, 90, 361, 71))
        self.groupBox_5.setTitle("")
        self.groupBox_5.setObjectName("groupBox_5")
        self.Tmerlable = QtWidgets.QLabel(self.groupBox_5)
        self.Tmerlable.setEnabled(True)
        self.Tmerlable.setGeometry(QtCore.QRect(40, 30, 141, 20))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.Tmerlable.setFont(font)
        self.Tmerlable.setFocusPolicy(QtCore.Qt.NoFocus)
        self.Tmerlable.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.Tmerlable.setFrameShadow(QtWidgets.QFrame.Plain)
        self.Tmerlable.setMidLineWidth(9)
        self.Tmerlable.setTextFormat(QtCore.Qt.RichText)
        self.Tmerlable.setScaledContents(True)
        self.Tmerlable.setAlignment(QtCore.Qt.AlignCenter)
        self.Tmerlable.setTextInteractionFlags(QtCore.Qt.NoTextInteraction)
        self.Tmerlable.setObjectName("Tmerlable")
        self.Timer = QtWidgets.QLineEdit(self.groupBox_5)
        self.Timer.setGeometry(QtCore.QRect(180, 30, 51, 22))
        self.Timer.setObjectName("Timer")
        self.Tmerlable_2 = QtWidgets.QLabel(self.groupBox_5)
        self.Tmerlable_2.setEnabled(True)
        self.Tmerlable_2.setGeometry(QtCore.QRect(230, 32, 41, 20))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.Tmerlable_2.setFont(font)
        self.Tmerlable_2.setFocusPolicy(QtCore.Qt.NoFocus)
        self.Tmerlable_2.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.Tmerlable_2.setFrameShadow(QtWidgets.QFrame.Plain)
        self.Tmerlable_2.setMidLineWidth(9)
        self.Tmerlable_2.setTextFormat(QtCore.Qt.RichText)
        self.Tmerlable_2.setScaledContents(True)
        self.Tmerlable_2.setAlignment(QtCore.Qt.AlignCenter)
        self.Tmerlable_2.setTextInteractionFlags(QtCore.Qt.NoTextInteraction)
        self.Tmerlable_2.setObjectName("Tmerlable_2")
        self.groupBox_6 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_6.setGeometry(QtCore.QRect(20, 170, 361, 51))
        self.groupBox_6.setTitle("")
        self.groupBox_6.setObjectName("groupBox_6")
        self.groupBox_4.raise_()
        self.groupBox_6.raise_()
        self.groupBox_5.raise_()
        self.start_b.raise_()
        self.Pause_b.raise_()
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 391, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Offline Orders"))
        self.start_b.setText(_translate("MainWindow", "Run"))
        self.Pause_b.setText(_translate("MainWindow", "Stopl"))
        self.Tmerlable.setText(_translate("MainWindow", "Operating timing"))
        self.Tmerlable_2.setText(_translate("MainWindow", "Sec"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
