# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Main_window.ui'
#
# Created by: PyQt5 UI code generator 5.15.10
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWin(object):
    def setupUi(self, MainWin):
        MainWin.setObjectName("MainWin")
        MainWin.setWindowModality(QtCore.Qt.ApplicationModal)
        MainWin.setEnabled(True)
        MainWin.resize(582, 432)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWin.sizePolicy().hasHeightForWidth())
        MainWin.setSizePolicy(sizePolicy)
        MainWin.setMinimumSize(QtCore.QSize(582, 432))
        MainWin.setMaximumSize(QtCore.QSize(582, 432))
        MainWin.setSizeIncrement(QtCore.QSize(582, 432))
        MainWin.setWindowOpacity(1.0)
        MainWin.setLayoutDirection(QtCore.Qt.LeftToRight)
        MainWin.setAutoFillBackground(False)
        MainWin.setLocale(QtCore.QLocale(QtCore.QLocale.English, QtCore.QLocale.UnitedStates))
        MainWin.setInputMethodHints(QtCore.Qt.ImhNone)
        self.centralwidget = QtWidgets.QWidget(MainWin)
        self.centralwidget.setObjectName("centralwidget")
        self.refreshchrome = QtWidgets.QCheckBox(self.centralwidget)
        self.refreshchrome.setEnabled(True)
        self.refreshchrome.setGeometry(QtCore.QRect(60, 170, 131, 30))
        font = QtGui.QFont()
        font.setFamily("Arial")
        self.refreshchrome.setFont(font)
        self.refreshchrome.setObjectName("refreshchrome")
        self.usersIDs = QtWidgets.QCheckBox(self.centralwidget)
        self.usersIDs.setGeometry(QtCore.QRect(60, 210, 151, 30))
        font = QtGui.QFont()
        font.setFamily("Arial")
        self.usersIDs.setFont(font)
        self.usersIDs.setObjectName("usersIDs")
        self.timer = QtWidgets.QLineEdit(self.centralwidget)
        self.timer.setGeometry(QtCore.QRect(440, 190, 61, 31))
        self.timer.setObjectName("timer")
        self.start = QtWidgets.QPushButton(self.centralwidget)
        self.start.setGeometry(QtCore.QRect(77, 330, 111, 31))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.start.sizePolicy().hasHeightForWidth())
        self.start.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(10)
        font.setItalic(False)
        font.setKerning(True)
        self.start.setFont(font)
        self.start.setObjectName("start")
        self.pause = QtWidgets.QPushButton(self.centralwidget)
        self.pause.setGeometry(QtCore.QRect(392, 330, 111, 31))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pause.sizePolicy().hasHeightForWidth())
        self.pause.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(10)
        font.setItalic(False)
        self.pause.setFont(font)
        self.pause.setObjectName("pause")
        self.UsersCodes = QtWidgets.QCheckBox(self.centralwidget)
        self.UsersCodes.setGeometry(QtCore.QRect(60, 250, 171, 30))
        font = QtGui.QFont()
        font.setFamily("Arial")
        self.UsersCodes.setFont(font)
        self.UsersCodes.setObjectName("UsersCodes")
        self.title = QtWidgets.QLabel(self.centralwidget)
        self.title.setGeometry(QtCore.QRect(30, 40, 511, 81))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setItalic(False)
        font.setUnderline(False)
        font.setWeight(50)
        font.setStrikeOut(False)
        font.setKerning(False)
        self.title.setFont(font)
        self.title.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.title.setAutoFillBackground(False)
        self.title.setTextFormat(QtCore.Qt.AutoText)
        self.title.setAlignment(QtCore.Qt.AlignCenter)
        self.title.setWordWrap(False)
        self.title.setObjectName("title")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(321, 191, 121, 31))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(507, 190, 31, 31))
        self.label_2.setObjectName("label_2")
        self.groupBox_4 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_4.setGeometry(QtCore.QRect(20, 10, 541, 141))
        font = QtGui.QFont()
        font.setPointSize(8)
        font.setItalic(False)
        self.groupBox_4.setFont(font)
        self.groupBox_4.setTitle("")
        self.groupBox_4.setObjectName("groupBox_4")
        self.groupBox_5 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_5.setGeometry(QtCore.QRect(20, 160, 541, 131))
        self.groupBox_5.setTitle("")
        self.groupBox_5.setObjectName("groupBox_5")
        self.label_4 = QtWidgets.QLabel(self.groupBox_5)
        self.label_4.setGeometry(QtCore.QRect(484, 78, 41, 31))
        self.label_4.setObjectName("label_4")
        self.groupBox_6 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_6.setGeometry(QtCore.QRect(20, 298, 541, 81))
        self.groupBox_6.setTitle("")
        self.groupBox_6.setObjectName("groupBox_6")
        self.Batch = QtWidgets.QLineEdit(self.centralwidget)
        self.Batch.setGeometry(QtCore.QRect(439, 239, 61, 31))
        self.Batch.setObjectName("Batch")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(320, 240, 121, 31))
        self.label_3.setObjectName("label_3")
        self.groupBox_6.raise_()
        self.groupBox_5.raise_()
        self.groupBox_4.raise_()
        self.refreshchrome.raise_()
        self.usersIDs.raise_()
        self.timer.raise_()
        self.start.raise_()
        self.pause.raise_()
        self.UsersCodes.raise_()
        self.title.raise_()
        self.label.raise_()
        self.label_2.raise_()
        self.Batch.raise_()
        self.label_3.raise_()
        MainWin.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWin)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 582, 26))
        self.menubar.setObjectName("menubar")
        MainWin.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWin)
        self.statusbar.setObjectName("statusbar")
        MainWin.setStatusBar(self.statusbar)

        self.retranslateUi(MainWin)
        QtCore.QMetaObject.connectSlotsByName(MainWin)

    def retranslateUi(self, MainWin):
        _translate = QtCore.QCoreApplication.translate
        MainWin.setWindowTitle(_translate("MainWin", "New Codinig System"))
        self.refreshchrome.setText(_translate("MainWin", "Refresh Chrome"))
        self.usersIDs.setText(_translate("MainWin", "Request Users IDs"))
        self.start.setText(_translate("MainWin", "Start"))
        self.pause.setText(_translate("MainWin", "pause"))
        self.UsersCodes.setText(_translate("MainWin", "Download Users Codes"))
        self.title.setText(_translate("MainWin", "Ready to Start"))
        self.label.setText(_translate("MainWin", "Set Execution Rate"))
        self.label_2.setText(_translate("MainWin", "Sec"))
        self.label_4.setText(_translate("MainWin", "Users"))
        self.label_3.setText(_translate("MainWin", "Set Batch Size"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWin = QtWidgets.QMainWindow()
    ui = Ui_MainWin()
    ui.setupUi(MainWin)
    MainWin.show()
    sys.exit(app.exec_())
