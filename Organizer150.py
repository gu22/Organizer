# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Organizer_gui.ui'
#
# Created by: PyQt5 UI code generator 5.14.2
#
# WARNING! All changes made in this file will be lost!

import datetime
import os.path
import PyQt5
from PyQt5 import *
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QDialog, QMessageBox
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QCoreApplication

try:
    from PyQt5 import sip
except ImportError:
    import sip

import openpyxl
from openpyxl import load_workbook
import pandas as pd

import shutil
from shutil import copyfile

from datetime import date
import sys
#EMAIL
import win32com.client



path_file = []
path_name = []
extensao=[]
files=[]
print_file = []
lista_pastas=[]
arq_dup=[]
verificador = "nada"
c = 0
data = datetime.datetime.now()
cont_prog = 0


meses = {1:'janeiro',2:'fevereiro',3:'março',
4:'abril',5:'maio',6:'junho',
7:'julho',8:'agosto',9:'setembro',
10:'outubro',11:'novembro',12:'dezembro'}

tipos = {"Atraso":"OTD e Baixa de entregas","DOCK":"Dock scheduling","Evento":"SIPA",
         "Rejeitado":"Disponibilidade","Responsabilidade":"Complaint","STATUS":"Monitoramento",
         }


ext_verific = [".xlsx",".xls"]

##=======================================================================

# --------------------------- LOG  ---------------------------------


class Logger(object):
    def __init__(self):
        self.terminal = sys.stdout

    def write(self, message):
        with open ("logfile.log", "a", encoding = 'utf-8') as self.log:            
            self.log.write(message)
        self.terminal.write(message)

    def flush(self):
        #this flush method is needed for python 3 compatibility.
        #this handles the flush command by doing nothing.
        #you might want to specify some extra behavior here.
        pass  

sys.stdout = Logger()

##==========================================================================


class Ui_Dialog(QDialog):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(521, 300)
        Dialog.setWindowIcon(QtGui.QIcon("or.ico"))
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
        self.textBrowser = QtWidgets.QTextBrowser(Dialog)
        self.textBrowser.setGeometry(QtCore.QRect(30, 40, 471, 16))
        self.textBrowser.setToolTip("")
        self.textBrowser.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.textBrowser.setObjectName("textBrowser")
        self.pushButton = QtWidgets.QPushButton(Dialog)
        self.pushButton.setGeometry(QtCore.QRect(30, 10, 161, 23))
        self.pushButton.setObjectName("pushButton")
        self.lcdNumber = QtWidgets.QLCDNumber(Dialog)
        self.lcdNumber.setGeometry(QtCore.QRect(370, 60, 64, 23))
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
        self.label.setGeometry(QtCore.QRect(360, 200, 141, 16))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(380, 80, 51, 31))
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
        self.label_3 = QtWidgets.QLabel(Dialog)
        self.label_3.setGeometry(QtCore.QRect(360, 130, 151, 61))
        self.label_3.setText("")
        self.label_3.setObjectName("label_3")
        self.lcdNumber_2 = QtWidgets.QLCDNumber(Dialog)
        self.lcdNumber_2.setGeometry(QtCore.QRect(440, 60, 64, 23))
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
        self.lcdNumber_2.setPalette(palette)
        self.lcdNumber_2.setDigitCount(5)
        self.lcdNumber_2.setObjectName("lcdNumber_2")
        self.label_4 = QtWidgets.QLabel(Dialog)
        self.label_4.setGeometry(QtCore.QRect(450, 80, 51, 31))
        self.label_4.setWordWrap(True)
        self.label_4.setObjectName("label_4")
        self.widget = QtWidgets.QWidget(Dialog)
        self.widget.setGeometry(QtCore.QRect(350, 260, 158, 25))
        self.widget.setObjectName("widget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.widget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.pushButton_2 = QtWidgets.QPushButton(self.widget)
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.setEnabled(False)
        self.horizontalLayout.addWidget(self.pushButton_2)
        self.buttonBox = QtWidgets.QDialogButtonBox(self.widget)
        self.buttonBox.setContextMenuPolicy(QtCore.Qt.NoContextMenu)
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.horizontalLayout.addWidget(self.buttonBox)

        self.retranslateUi(Dialog)
        self.tabWidget.setCurrentIndex(0)
        self.buttonBox.accepted.connect(Dialog.accept)
        self.buttonBox.rejected.connect(Dialog.reject)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

#------------Funçoes e ações -------------------------------------------       
        self.pushButton.clicked.connect(self.abrir_arquivo)
        self.pushButton_2.clicked.connect(self.transforme)
        
    
##================================================================================




#----------------- Definindo as ações ----------------------------      
    def abrir_arquivo(self):
        global path_file
        global path_name
        global files
        global extensao
        global print_file
        global data
        global cont_prog
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
        
        
        
        
        for i in path_file: #Path_file = arquivos que estão na pasta
            dir_check = os.path.join(path_name,i)
            if os.path.isdir(dir_check) is False:
                    files.append(os.path.join(path_name,i))
            else:
                pass
        
        for i in files:
            extensao.append(os.path.splitext(i)[1])  
        
        # for i in files:
        #     if ext_verific[0] in i:
        #         print_file.append(i)
        #     elif ext_verific[1] in i:
        #         print_file.append(i)
        #     else:
        #         pass 
           
        for i in path_file:
           if ext_verific[0] in i:
               print_file.append(i)
           elif ext_verific[1] in i:
               print_file.append(i)
           else:
               pass    
        
        cont_file = len(print_file) 
        cont_prog = cont_file
        print_file = "\n".join(print_file)
        
        nomediretorio = ".BASE_Original"
        
                # criando diretorio com nome Base_Original
        diretorio_baseoriginal = (os.path.join(path_name,nomediretorio))
        teste_base = os.path.exists(diretorio_baseoriginal)
        
        if teste_base is False:
            os.mkdir(diretorio_baseoriginal)
            print(" Base Original - Criada com Sucesso")
            
        #criando diretorio de armazenamento de acordo com data 
        data_pasta = (str(data.strftime("%y.%m.%d_%H.%M.%S")))
        data_pasta = (os.path.join(path_name,data_pasta))   
        os.mkdir(data_pasta)
        print("PASTA DATA CRIADA")
        print("")
        
        
        # Criando copy dos arquivos originais para a pasta com data >>> base_original
        c=0
        for g in files:
           shutil.copy(g, data_pasta)
           c +=1
        
        # movendo pasta
                
        shutil.move(data_pasta, diretorio_baseoriginal)
                
        
        self.lcdNumber.display(cont_file)
        self.progressBar.setValue(40)
        
        self.textBrowser_2.setText(print_file)
        self.pushButton.setEnabled(False)
        self.pushButton_2.setEnabled(True)
        self.buttonBox.setEnabled(False)
        self.label_3.setText("Arquivos Carregados\n\nPressione Prosseguir")
        self.label.setText("Em Andamento")
        
        
        
        
    def transforme(self):
        global path_file
        global path_name
        global files
        global extensao
        global print_file
        global c
        global verificador
        global lista_pastas
        global arq_dup
        global meses
        global tipos
        global cont_prog
        
        output=[]
        vbase = 40
        text_d = ""
        progresso = (vbase/cont_prog)
        lib = 0
        
        for i in files:
    
            ext_check = extensao[c]
            verificador = ''
            if ext_check in ext_verific:
            
                base = pd.read_excel(i)
            
                head =base.columns[0]
                ncolunas = len(base.columns)
            
                if ncolunas >= 5 :
                    verificador = base.iloc[1][4]
            
                # loop para capturar o mes
                for x in meses:
                    if meses[x] in head:
                        mes = meses[x]
                        print(mes)
                    else:
                        pass
             
                # loop para identificar tipo das informações
                for y in tipos:
                    if y in head:
                        tipo = tipos[y]
                        break
                    elif y in verificador:
                        tipo = tipos[y]
                        break
                    # else:
                    #     tipo = tipos[y]
                  
                #definindo a transportadorra    
                transportadora = base.iloc[2][0]
                
                # Atribuindo nome original
                arquivo = files[c]
                
                # Pegando extensão
                ext = extensao[c]
                
                #Renomeando arquivo
                os.rename(arquivo,(os.path.join(path_name,((tipo+" - "+transportadora+"_").upper())+(mes)+ext)))
                print((path_name+(tipo+" - "+transportadora+"_").upper()+(mes)+ext))
                print('<<<<<<<<<<<<<>>>>>>>>>>>>>')
                print(arquivo,(os.path.join(path_name,((tipo+" - "+transportadora+"_").upper())+(mes)+ext)))
        
                
                #definido arquivo que sera movido
                arq_move = (os.path.join(path_name,((tipo+" - "+transportadora+"_").upper()+(mes)+ext)))
        
                # Transferindo arquivo para a pasta correta (TESTE e transferencia)
                #centro_catch = str(base.iloc[2][1])
                centro = str(base.iloc[2][1])
                pasta_centro = os.path.join(path_name,centro)
                teste_pastacentro = os.path.exists(pasta_centro)
                if teste_pastacentro is True:
                        print(pasta_centro,"<<<<<>>>>>")
                        print("")
                        # if centro_catch == i:
                        #     move(arq_move, pasta_centro)
                else:
                      os.mkdir(pasta_centro)
                      
                # Realizando  e criando pasta/diretorio mes
                pasta_mes = os.path.join(pasta_centro,mes.upper())
                teste_pasta = os.path.exists(pasta_mes)
                if teste_pasta is True:
                    print ("Mes ja existe")
                    print ("---------------------")
                    print("")
                else:
                    os.mkdir(pasta_mes)
                
                #Criando pasta Transportadora
                #separando por transportadora
                pasta_tsp = os.path.join(pasta_mes,str(transportadora))
                teste_tsp = os.path.exists(pasta_tsp)
                if teste_tsp is False:
                    os.mkdir(pasta_tsp)
                
                arq_move_tst = arq_move
                arq_move_tst = os.path.split(arq_move)
                arq_move_tst = arq_move_tst[-1]
                if not arq_move_tst in os.listdir(pasta_tsp):     
                    shutil.move(arq_move,pasta_tsp) 
                else:
                    entrada_d = i.split("\\")[-1]
                    arq_dup.append(entrada_d)
                    print("Arquivo Duplicado\n",i,"\n",arq_move,"\n_______________________\n")
                    lib +=1
                 
                if not pasta_tsp in lista_pastas:
                    lista_pastas.append(pasta_tsp)
                
                
                c+=1
                output.append("{} : {} - {},{}".format(transportadora,tipo,mes,centro))
                vbase = vbase + progresso
                self.progressBar.setValue(vbase)
                print(vbase)
            else:
                vbase = vbase + progresso
                self.progressBar.setValue(vbase)
                c+=1
        if len(arq_dup) > 0:
            text_d = ("\n").join(arq_dup)
            text_d = "Arquivos duplicados\n{}".format(text_d)
     
        output.sort()
        cont_file = len(output) - lib
        output = ("\n").join(output)
        
        self.tab_2.setEnabled(True)
        self.textBrowser_3.setEnabled(True)
        self.tabWidget.setCurrentIndex(1)
        self.textBrowser_3.setText(output)
        self.lcdNumber_2.display(cont_file)
        self.pushButton_2.setEnabled(False)    
        self.label_3.setText(text_d)
        
        msg = QMessageBox()
        msg.setText("Gostaria de enviar os email agora?\nOK para sim")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        retval = msg.exec_()
        
        teste = int(retval)
        print(teste)
        if teste == 1024:
            qsend = True
        else:
            qsend = False
        # qsend = ""
        # msg.buttonClicked(qsend = True)        
        # msg.rejected(qsend = False)        
        
        # def email(self):
        #     global data
        #     global meses
            
        #     nonlocal lista_pastas
            
        for item in lista_pastas:
    
            orientacao_dia = int(data.strftime("%H"))
            saudacao=["Bom dia, ","Boa Tarde, ","Boa noite, "]    
            
            if orientacao_dia <=11:
                saudacao_email = saudacao[0]
            elif orientacao_dia <=17:
                saudacao_email = saudacao[1]
            else:
                saudacao_email = saudacao[2]
               
               
            email_path = os.listdir(item)
            
            for x in meses:
                if meses[x] in item.lower():
                    mes = meses[x]
                    print(mes)
                else:
                    pass
            
            mes = mes.capitalize()  
            
            trasp_email = os.path.split(item)
            print(trasp_email)
            trasp_email = trasp_email[-1]
            print(trasp_email)
            
            outlook = win32com.client.Dispatch('Outlook.Application')
            email = outlook.CreateItem(0)
            # email = outlook.CreateItemFromTemplate(os.getcwd() + '\\cte.msg')
            email.To= ''
            email.BodyFormat= 2
            email.Subject="Avalição dos Fornecedores - Base de Dados (%s) - %s"%(trasp_email,mes)
            # email.Subject= email.Subject.replace('[compName]','test')
            email.HTMLBody= (
            """%s 
                esperamos que todos estejam bem!<p>
                
                Nossa reunião de avaliação está próxima!<br>
                Estamos disponibilizando a base de dados referente ao mês de <b> %s  </b> para que vocês possam
                analisar e nos informar o que ocorreu.<p> 
                Para qualquer tipo de dúvida, estamos à disposição!<br>"""%(saudacao_email,mes)
                
                
                )
            # email.HTMLBody= email.HTMLBody.replace('fname', 'test')
            for x in email_path:
                email.Attachments.Add(Source= os.path.join(item,x))
            
           
            
            
            if qsend is True:
                email.Display(False)
            
            
            email.SaveAs(item+ '\\EMAIL- %s_%s.msg'%(trasp_email,mes), 3)
        
        self.label.setText("Finalizado!")    
        self.progressBar.setValue(100)    
        self.buttonBox.setEnabled(True)
            
##=================================================================================







    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Organizador de arquivos (Organizer v1.5)"))
        self.pushButton.setText(_translate("Dialog", "Selecionar Pasta"))
        self.label.setText(_translate("Dialog", "Processo"))
        self.label_2.setText(_translate("Dialog", "Arquivos Input"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("Dialog", "Arquivos Input"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("Dialog", "Arquivos Output"))
        self.label_4.setText(_translate("Dialog", "Arquivos Output"))
        self.pushButton_2.setText(_translate("Dialog", "Prosseguir"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
# sys.stdout = Logger()