# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Organizer_gui.ui'
#
# Created by: PyQt5 UI code generator 5.14.2
#
# WARNING! All changes made in this file will be lost!


import os.path
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QDialog
from PyQt5.QtGui import QIcon


path_file = []
path_name = []
extensao=[]
files=[]
print_file = []

meses = {1:'janeiro',2:'fevereiro',3:'março',
4:'abril',5:'maio',6:'junho',
7:'julho',8:'agosto',9:'setembro',
10:'outubro',11:'novembro',12:'dezembro'}

tipos = {"Atraso":"OTD e Baixa de entregas","DOCK":"Dock scheduling","Evento":"SIPA",
         "Rejeitado":"Disponibilidade","Responsabilidade":"Complaint","STATUS":"Monitoramento",
         }


ext_verific = [".xlsx",".xls"]

##=======================================================================

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
        self.progressBar.setProperty("value", 24)
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
        self.tabWidget = QtWidgets.QTabWidget(Dialog)
        self.tabWidget.setEnabled(True)
        self.tabWidget.setGeometry(QtCore.QRect(20, 70, 321, 221))
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.textBrowser_2 = QtWidgets.QTextBrowser(self.tab)
        self.textBrowser_2.setGeometry(QtCore.QRect(10, 10, 291, 181))
        self.textBrowser_2.setObjectName("textBrowser_2")
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setEnabled(False)
        self.tab_2.setObjectName("tab_2")
        self.textBrowser_3 = QtWidgets.QTextBrowser(self.tab_2)
        self.textBrowser_3.setEnabled(False)
        self.textBrowser_3.setGeometry(QtCore.QRect(10, 10, 291, 181))
        self.textBrowser_3.setObjectName("textBrowser_3")
        self.tabWidget.addTab(self.tab_2, "")

        self.retranslateUi(Dialog)
        self.tabWidget.setCurrentIndex(0)
        self.buttonBox.accepted.connect(Dialog.accept)
        self.buttonBox.rejected.connect(Dialog.reject)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

#------------Funçoes e ações -------------------------------------------       
        self.pushButton.clicked.connect(self.abrir_arquivo)
       
        
    
##================================================================================




#----------------- Definindo as ações ----------------------------      
    def abrir_arquivo(self):
        global path_file
        global path_name
        global files
        global extensao
        global print_file
        # self.getfile = QFileDialog.getOpenFileName(self,"Open Dir")
        # get = QFileDialog.getOpenFileName(self,"Open Dir")
        path_name = QFileDialog.getExistingDirectory(self,"Open Dir")
        # print(self.getfile)
        print(path_name)
        
        self.textBrowser.setText(path_name)
        path_file = os.listdir(path_name)
        print(path_file)
        
        # arq = ("\n").join(path_file)
        # self.textBrowser_2.setText(arq)
        
        cont_file = len(path_file)
        self.lcdNumber.display(cont_file)
        self.progressBar.setValue(50)
        
        
        for i in path_file: #Path_file = arquivos que estão na pasta
            dir_check = os.path.join(path_name,i)
            if os.path.isdir(dir_check) is False:
                    files.append(os.path.join(path_name,i))
            else:
                pass
        
        # for i in files:
        #     extensao.append(os.path.splitext(i)[1])  
        
        # for i in files:
        #    if ext_verific[0] in i:
        #        print_file.append(i)
        #    elif ext_verific[1] in i:
        #        print_file.append(i)
        #    else:
        #        pass 
           
        for i in path_file:
           if ext_verific[0] in i:
               print_file.append(i)
           elif ext_verific[1] in i:
               print_file.append(i)
           else:
               pass    
       
        print_file = "\n".join(print_file)
        
        self.textBrowser_2.setText(print_file)
        
    def transforme(self):
        global path_file
        
        
            
    
    
    def teste(self):
        print("Teste ok")
        
        
        
##=================================================================================













    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Organizador de arquivos v1.5"))
        self.pushButton.setText(_translate("Dialog", "Selecionar Pasta"))
        self.label.setText(_translate("Dialog", "Processo"))
        self.label_2.setText(_translate("Dialog", "Contagem Arquivos"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("Dialog", "Arquivos Input"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("Dialog", "Arquivos Output"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
