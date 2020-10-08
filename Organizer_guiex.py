# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Organizer_gui.ui'
#
# Created by: PyQt5 UI code generator 5.14.2
#
# WARNING! All changes made in this file will be lost!


import os.path
import os
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QDialog
from PyQt5.QtGui import QIcon


class Ui_Dialog(QDialog):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(521, 300)
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(9, 9, 9))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(9, 9, 9))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(9, 9, 9))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Button, brush)
        Dialog.setPalette(palette)
        self.buttonBox = QtWidgets.QDialogButtonBox(Dialog)
        self.buttonBox.setGeometry(QtCore.QRect(430, 250, 81, 32))
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.textBrowser = QtWidgets.QTextBrowser(Dialog)
        self.textBrowser.setGeometry(QtCore.QRect(30, 40, 471, 16))
        self.textBrowser.setToolTip("")
        self.textBrowser.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.textBrowser.setObjectName("textBrowser")
        self.pushButton = QtWidgets.QPushButton(Dialog)
        self.pushButton.setGeometry(QtCore.QRect(30, 10, 161, 23))
        self.pushButton.setObjectName("pushButton")
        self.groupBox = QtWidgets.QGroupBox(Dialog)
        self.groupBox.setGeometry(QtCore.QRect(20, 80, 301, 201))
        self.groupBox.setObjectName("groupBox")
        self.textBrowser_2 = QtWidgets.QTextBrowser(self.groupBox)
        self.textBrowser_2.setGeometry(QtCore.QRect(10, 20, 281, 171))
        self.textBrowser_2.setObjectName("textBrowser_2")
        self.lcdNumber = QtWidgets.QLCDNumber(Dialog)
        self.lcdNumber.setGeometry(QtCore.QRect(440, 60, 64, 23))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(198, 198, 198))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Light, brush)
        brush = QtGui.QBrush(QtGui.QColor(198, 198, 198))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Light, brush)
        brush = QtGui.QBrush(QtGui.QColor(198, 198, 198))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Button, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Light, brush)
        self.lcdNumber.setPalette(palette)
        self.lcdNumber.setDigitCount(5)
        self.lcdNumber.setObjectName("lcdNumber")
        self.progressBar = QtWidgets.QProgressBar(Dialog)
        self.progressBar.setGeometry(QtCore.QRect(360, 223, 151, 20))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setTextVisible(True)
        self.progressBar.setInvertedAppearance(False)
        self.progressBar.setObjectName("progressBar")
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(360, 200, 47, 13))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(450, 90, 51, 31))
        self.label_2.setWordWrap(True)
        self.label_2.setObjectName("label_2")

        self.retranslateUi(Dialog)
        self.buttonBox.accepted.connect(Dialog.accept)
        self.buttonBox.rejected.connect(Dialog.reject)
        QtCore.QMetaObject.connectSlotsByName(Dialog)
        

#------------Funçoes e ações -------------------------------------------       
        self.pushButton.clicked.connect(self.abrir_arquivo)
       
        
    
##================================================================================




#----------------- Definindo as ações ----------------------------      
    def abrir_arquivo(self):
        # self.getfile = QFileDialog.getOpenFileName(self,"Open Dir")
        # get = QFileDialog.getOpenFileName(self,"Open Dir")
        get = QFileDialog.getExistingDirectory(self,"Open Dir")
        # print(self.getfile)
        print(get)
        self.textBrowser.setText(get)
        path_file = os.listdir(get)
        print(path_file)
        arq = ("\n").join(path_file)
        self.textBrowser_2.setText(arq)
        cont_file = len(path_file)
        self.lcdNumber.display(cont_file)
        self.progressBar.setValue(50)
    
    def teste(self):
        print("Teste ok")
        
        
        
##=================================================================================





    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Organizador de arquivos v1.5"))
        self.pushButton.setText(_translate("Dialog", "Selecionar Pasta"))
        self.groupBox.setTitle(_translate("Dialog", "Arquivos"))
        self.label.setText(_translate("Dialog", "Processo"))
        self.label_2.setText(_translate("Dialog", "Contagem Arquivos"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
