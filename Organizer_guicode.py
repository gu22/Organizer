# -*- coding: utf-8 -*-
"""
Created on Tue Oct  6 11:25:05 2020

@author: glpyz
"""

from PyQt5 import QtCore, QtGui, QtWidgets
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QDialog
from PyQt5.QtGui import QIcon


import Organizer_guidsg
from Organizer_guidsg import *




def texto():
    print("teste")
def abrir_arquivo(self):
        # self.getfile = QFileDialog.getOpenFileName(self,"Open Dir")
        # get = QFileDialog.getOpenFileName(self,"Open Dir")
        get = gui.QFileDialog.getExistingDirectory(self,"Open Dir")
        # print(self.getfile)
        print(get)
        self.gui.textBrowser.setText(get)
        path_file = os.listdir(get)
        print(path_file)
        arq = ("\n").join(path_file)
        self.gui.textBrowser_2.setText(arq)
        cont_file = len(path_file)
        self.gui.lcdNumber.display(cont_file)
        self.gui.progressBar.setValue(50)







gui = Ui_Dialog()

gui.setupUi(QtWidgets.QPushButton(Dialog))



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())



 