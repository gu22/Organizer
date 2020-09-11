#import ordenador_arqs
import datetime
import easygui
import os, glob
import os.path
import sys
import glob
from os import *

from easygui import *

import openpyxl
from openpyxl import load_workbook
import pandas as pd

import shutil
from shutil import copyfile

from datetime import date

#EMAIL
import win32com.client

sys.setrecursionlimit(5000)

Data = datetime.datetime.now()
print(Data)
print ("Sistema de Organização de bases")
print("v 1.04")

shutil.move

#Preparação para tratamento
meses = {1:'janeiro',2:'fevereiro',3:'março',
4:'abril',5:'maio',6:'junho',
7:'julho',8:'agosto',9:'setembro',
10:'outubro',11:'novembro',12:'dezembro'}

tipos = {"Atraso":"OTD e Baixa de entregas","DOCK":"Dock scheduling","Evento":"SIPA",
         "Rejeitado":"Disponibilidade","Responsabilidade":"Complaint","STATUS":"Monitoramento",
         }

mes = 0



# Definindo como realizar a verificação de arquivos
msg = easygui.msgbox("""Os arquivos já estão na pasta?\nCertifique-se que:\n
                       - Criou um pasta (importante para nao comprometer
                        outros arquivos);\n
                       - Colocou os arquivos de Avaliação ""","Verificação de Arquivos")

while True:
    user_dir = easygui.diropenbox()
    
    choice = easygui.indexbox(("A pasta escolhida é:\n\n%s"%user_dir),
                              "Diretorio escolhido",
                               choices=("Sim","Não","Cancelar"),
                               cancel_choice="Cancelar")
    
    
    if choice == 0:
        break
    elif choice == 1:
        pass
    else:        
        sys.exit()
        print("falha")
        
files_show = 1        
while True:
# Loop de verificação de arquivos
    
    if files_show == 0:
        break
    elif files_show == 1 :    
        Path_name = user_dir
        print(Path_name) #Path_name = onde estao os arquivos
        print("-------------------------")
        print("")
        
        Path_file = os.listdir(Path_name)
        print (Path_file)
        
        Files=[]
        ext_verific = [".xlsx",".xls"]
        for i in Path_file: #Path_file = arquivos que estão na pasta
            dir_check = os.path.join(Path_name,i)
            if os.path.isdir(dir_check) is False:
                    Files.append(os.path.join(Path_name,i))
            else:
                pass
        
        print(Files)
        
        # extraindo extensao
        extensao=[]
        for i in Files:
            
            extensao.append(os.path.splitext(i)[1])
            print(extensao)
         
        print_file = []
        
        for i in Files:
            if ext_verific[0] in i:
                print_file.append(i)
            elif ext_verific[1] in i:
                print_file.append(i)
            else:
                pass
        
            
        print_file = "\n".join(print_file) 
        
        
        
        files_show = easygui.indexbox(("Os arquivos que serão renomeados, são:\n\n%s"%print_file),
                                      "Arquivos Localizados",
                                      choices=("Sim","Não","Cancelar"),
                                   cancel_choice="Cancelar")
        if files_show == 1:
            easygui.msgbox("""REMOVA OS ARQUIVOS QUE NÃO DEVEM SER RENOMEADOS\n
                           Precione OK - Após retirar""","ALERTA")
    else:
        sys.exit()
        print ('saindo')


nomediretorio = ".BASE_Original"

# criando diretorio com nome Base_Original
diretorio_baseoriginal = (os.path.join(Path_name,nomediretorio))
teste_base = os.path.exists(diretorio_baseoriginal)

if teste_base is False:
    os.mkdir(diretorio_baseoriginal)
    print(" Base Original - Criada com Sucesso")
    
#criando diretorio de armazenamento de acordo com data 
data_pasta = (str(Data.strftime("%y.%m.%d_%H.%M.%S")))
data_pasta = (os.path.join(Path_name,data_pasta))   
os.mkdir(data_pasta)
print("PASTA DATA CRIADA")
print("")


# Criando copy dos arquivos originais para a pasta com data >>> base_original
c=0
for g in Files:
   shutil.copy(g, data_pasta)
   c +=1

# movendo pasta
        
shutil.move(data_pasta, diretorio_baseoriginal)


# fazer função que excluir outros tipos de arquivo <<<<<
verificador = "nada"
c = 0


# Capturandos  os dados para classificar os arquivos
lista_pastas=[]
arq_dup=[]
for i in Files:
    
    ext_check = extensao[c]
    
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
            else:
                tipo = tipos[y]
          
        #definindo a transportadorra    
        transportadora = base.iloc[2][0]
        
        # Atribuindo nome original
        arquivo = Files[c]
        
        # Pegando extensão
        ext = extensao[c]
        
        #Renomeando arquivo
        os.rename(arquivo,(os.path.join(Path_name,((tipo+" - "+transportadora+"_").upper())+(mes)+ext)))
        print((Path_name+(tipo+" - "+transportadora+"_").upper()+(mes)+ext))
        print('<<<<<<<<<<<<<>>>>>>>>>>>>>')
        print(arquivo,(os.path.join(Path_name,((tipo+" - "+transportadora+"_").upper())+(mes)+ext)))

        
        #definido arquivo que sera movido
        arq_move = (os.path.join(Path_name,((tipo+" - "+transportadora+"_").upper()+(mes)+ext)))

        # Transferindo arquivo para a pasta correta (TESTE e transferencia)
        #centro_catch = str(base.iloc[2][1])
        centro = str(base.iloc[2][1])
        pasta_centro = os.path.join(Path_name,centro)
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
            arq_dup.append(i)
            print("Arquivo Duplicado\n",i,"\n",arq_move,"\n_______________________\n")
         
        if not pasta_tsp in lista_pastas:
            lista_pastas.append(pasta_tsp)
        
        
        c+=1

    else:
        c+=1
if len(arq_dup) > 0:
       
    easygui.msgbox(("Arquivos duplicados:\n %s "%arq_dup),"Arquivos Duplicados")        
        

#FUNÇAO DE EMAIL <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

qsend = easygui.ynbox(""" <<< ABRIR O OUTLOOK ANTES DE PROSSEGUIR OU OS EMAILS NÃO SERÃO CRIADOS >>> \nGostaria de enviar os emails agora?\n
                          Sim - para abrir todos os email e inserir os contatos\n
                          Não - para enviar mais tarde (todos emails ficarão salvos na pasta da transportadora)""","Enviar email",("Sim","Não"))

if qsend is False:
    easygui.msgbox("Os emails estão na pasta de cada transportadora","Lembrete")

for item in lista_pastas:
    
    orientacao_dia = int(Data.strftime("%H"))
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
#os.system("cd %s"%user_dir)