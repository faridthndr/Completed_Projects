# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Booking_window3.ui'
#
# Created by: PyQt5 UI code generator 5.15.10
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(335, 515)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        MainWindow.setMinimumSize(QtCore.QSize(335, 515))
        MainWindow.setMaximumSize(QtCore.QSize(335, 515))
        MainWindow.setSizeIncrement(QtCore.QSize(4, 0))
        MainWindow.setContextMenuPolicy(QtCore.Qt.DefaultContextMenu)
        MainWindow.setAcceptDrops(False)
        MainWindow.setWindowTitle("Booking Observer")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("../../Program/Python/Project/logo1.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setIconSize(QtCore.QSize(30, 28))
        MainWindow.setAnimated(True)
        MainWindow.setTabShape(QtWidgets.QTabWidget.Rounded)
        MainWindow.setUnifiedTitleAndToolBarOnMac(False)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox.setGeometry(QtCore.QRect(10, 390, 315, 75))
        self.groupBox.setTitle("")
        self.groupBox.setFlat(False)
        self.groupBox.setObjectName("groupBox")
        self.puase_b = QtWidgets.QPushButton(self.groupBox)
        self.puase_b.setGeometry(QtCore.QRect(200, 30, 93, 28))
        self.puase_b.setObjectName("puase_b")
        self.start_b = QtWidgets.QPushButton(self.groupBox)
        self.start_b.setGeometry(QtCore.QRect(30, 30, 93, 28))
        self.start_b.setObjectName("start_b")
        self.groupBox_2 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_2.setGeometry(QtCore.QRect(10, 270, 315, 111))
        self.groupBox_2.setTitle("")
        self.groupBox_2.setFlat(False)
        self.groupBox_2.setObjectName("groupBox_2")
        self.label = QtWidgets.QLabel(self.groupBox_2)
        self.label.setGeometry(QtCore.QRect(30, 10, 81, 21))
        self.label.setFocusPolicy(QtCore.Qt.NoFocus)
        self.label.setObjectName("label")
        self.Seconds_num = QtWidgets.QLineEdit(self.groupBox_2)
        self.Seconds_num.setGeometry(QtCore.QRect(130, 10, 41, 22))
        self.Seconds_num.setObjectName("Seconds_num")
        self.label_2 = QtWidgets.QLabel(self.groupBox_2)
        self.label_2.setGeometry(QtCore.QRect(200, 10, 81, 21))
        self.label_2.setFocusPolicy(QtCore.Qt.NoFocus)
        self.label_2.setObjectName("label_2")
        self.T0 = QtWidgets.QCheckBox(self.groupBox_2)
        self.T0.setGeometry(QtCore.QRect(20, 50, 51, 20))
        self.T0.setLocale(QtCore.QLocale(QtCore.QLocale.English, QtCore.QLocale.TrinidadAndTobago))
        self.T0.setObjectName("T0")
        self.T1 = QtWidgets.QCheckBox(self.groupBox_2)
        self.T1.setGeometry(QtCore.QRect(126, 50, 51, 20))
        self.T1.setLocale(QtCore.QLocale(QtCore.QLocale.English, QtCore.QLocale.TrinidadAndTobago))
        self.T1.setObjectName("T1")
        self.T2 = QtWidgets.QCheckBox(self.groupBox_2)
        self.T2.setGeometry(QtCore.QRect(203, 50, 91, 20))
        self.T2.setLocale(QtCore.QLocale(QtCore.QLocale.English, QtCore.QLocale.TrinidadAndTobago))
        self.T2.setObjectName("T2")
        self.T2_unbooking = QtWidgets.QCheckBox(self.groupBox_2)
        self.T2_unbooking.setGeometry(QtCore.QRect(203, 80, 111, 20))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.T2_unbooking.sizePolicy().hasHeightForWidth())
        self.T2_unbooking.setSizePolicy(sizePolicy)
        self.T2_unbooking.setLocale(QtCore.QLocale(QtCore.QLocale.English, QtCore.QLocale.TrinidadAndTobago))
        self.T2_unbooking.setObjectName("T2_unbooking")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(40, 220, 81, 21))
        self.label_3.setFocusPolicy(QtCore.Qt.NoFocus)
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(40, 190, 81, 21))
        self.label_4.setFocusPolicy(QtCore.Qt.NoFocus)
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(40, 130, 81, 21))
        self.label_5.setFocusPolicy(QtCore.Qt.NoFocus)
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(40, 160, 81, 21))
        self.label_6.setFocusPolicy(QtCore.Qt.NoFocus)
        self.label_6.setObjectName("label_6")
        self.PW_input = QtWidgets.QLineEdit(self.centralwidget)
        self.PW_input.setGeometry(QtCore.QRect(130, 219, 121, 22))
        self.PW_input.setObjectName("PW_input")
        self.User_N_input = QtWidgets.QLineEdit(self.centralwidget)
        self.User_N_input.setGeometry(QtCore.QRect(130, 190, 121, 22))
        self.User_N_input.setObjectName("User_N_input")
        self.Db_input = QtWidgets.QLineEdit(self.centralwidget)
        self.Db_input.setGeometry(QtCore.QRect(130, 160, 121, 22))
        self.Db_input.setObjectName("Db_input")
        self.Server_input = QtWidgets.QLineEdit(self.centralwidget)
        self.Server_input.setGeometry(QtCore.QRect(130, 130, 121, 22))
        self.Server_input.setObjectName("Server_input")
        self.groupBox_3 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_3.setGeometry(QtCore.QRect(10, 110, 315, 151))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.groupBox_3.sizePolicy().hasHeightForWidth())
        self.groupBox_3.setSizePolicy(sizePolicy)
        self.groupBox_3.setMaximumSize(QtCore.QSize(16777211, 16777215))
        self.groupBox_3.setBaseSize(QtCore.QSize(0, 0))
        self.groupBox_3.setTabletTracking(False)
        self.groupBox_3.setContextMenuPolicy(QtCore.Qt.NoContextMenu)
        self.groupBox_3.setTitle("")
        self.groupBox_3.setFlat(False)
        self.groupBox_3.setCheckable(False)
        self.groupBox_3.setChecked(False)
        self.groupBox_3.setObjectName("groupBox_3")
        self.groupBox_4 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_4.setGeometry(QtCore.QRect(10, 10, 315, 80))
        self.groupBox_4.setTitle("")
        self.groupBox_4.setObjectName("groupBox_4")
        self.Status_label = QtWidgets.QLabel(self.groupBox_4)
        self.Status_label.setGeometry(QtCore.QRect(10, 20, 291, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.Status_label.setFont(font)
        self.Status_label.setFocusPolicy(QtCore.Qt.NoFocus)
        self.Status_label.setText("")
        self.Status_label.setObjectName("Status_label")
        self.groupBox_4.raise_()
        self.groupBox.raise_()
        self.groupBox_2.raise_()
        self.groupBox_3.raise_()
        self.label_3.raise_()
        self.label_4.raise_()
        self.label_5.raise_()
        self.label_6.raise_()
        self.PW_input.raise_()
        self.User_N_input.raise_()
        self.Db_input.raise_()
        self.Server_input.raise_()
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 335, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setStatusTip("")
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionQuery_Rate = QtWidgets.QAction(MainWindow)
        self.actionQuery_Rate.setObjectName("actionQuery_Rate")

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        self.puase_b.setText(_translate("MainWindow", "Pause"))
        self.start_b.setText(_translate("MainWindow", "Run"))
        self.label.setText(_translate("MainWindow", "Query every"))
        self.label_2.setText(_translate("MainWindow", "Seconds"))
        self.T0.setText(_translate("MainWindow", "T+0"))
        self.T1.setText(_translate("MainWindow", "T+1"))
        self.T2.setText(_translate("MainWindow", "T2 Booking"))
        self.T2_unbooking.setText(_translate("MainWindow", "T2 Un-Booking"))
        self.label_3.setText(_translate("MainWindow", "Password"))
        self.label_4.setText(_translate("MainWindow", "User name"))
        self.label_5.setText(_translate("MainWindow", "Server"))
        self.label_6.setText(_translate("MainWindow", "Database"))
        self.actionQuery_Rate.setText(_translate("MainWindow", "Query Rate"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
