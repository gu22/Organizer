

#import ordenador_arqs
import datetime
import easygui
import os, glob
import os.path
import sys
from os import *

from easygui import *

import openpyxl
from openpyxl import load_workbook
import pandas as pd

from shutil import *
from shutil import copyfile

from datetime import date


#os.system('cmd /k Organizer.exe > log.txt')

Data = datetime.datetime.now()
print(Data)
print ("Sistema de Organização de bases")
print("v 1.01")


#File = easygui.fileopenbox()
#print(File)

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
msg = easygui.msgbox("""Os arquivos já estão na pasta?\nGaranta que:\n
                       - Criou um pasta (importante para nao comprometer
                        outros arquivos);\n
                       - Colocou os arquivos de base ""","Verificação de Arquivos")

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
            Files.append(os.path.join(Path_name,i))
        
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
    
    
        
    
    
# else:
#     Path_file = os.listdir()
#     print(Path_file)
#     Path = os.getcwd()

#     Files = []
#     for i in Path_file:
#         Files.append(os.path.join(Path, i))

    # print(Files)


nomediretorio = ".BASE_Original"

# criando diretorio com nome Base_Original
diretorio_baseoriginal = (os.path.join(Path_name,nomediretorio))
teste_base = os.path.exists(diretorio_baseoriginal)

if teste_base is False:
    os.mkdir(diretorio_baseoriginal)
    print(" Base Original - Criada com Sucesso")

# try:
#     os.mkdir(diretorio_baseoriginal)
# except OSError:
#     print("Pasta já existe")
# else:
#     print("Pasta: BASE_Original; criada com sucesso")
    
#criando diretorio de armazenamento de acordo com data 
data_pasta = (str(Data.strftime("%y.%m.%d_%H.%M.%S")))
data_pasta = (os.path.join(Path_name,data_pasta))   
os.mkdir(data_pasta)
print("PASTA DATA CRIADA")
print("")


# Criando copy dos arquivos originais para a pasta com data >>> base_original
c=0
for g in Files:
    try:
        copy(g, data_pasta)
        # copyfile(g,os.path.join(data_pasta,Path_file[c]))
        #copyfile(g,(os.path.join(diretorio,Path_file[c])))
        # print(copyfile(g,os.path.join(data_pasta,Path_file[c])))
        # print("")
        # print(copyfile(g,(os.path.join(diretorio,Path_file[c]))))
        c +=1
    except PermissionError:
        print ("Pasta de destino já contém arquivos com o mesmo nome, por favor mover ou excluir a pasta")

# movendo pasta
        
move(data_pasta, diretorio_baseoriginal)


# fazer função que excluir outros tipos de arquivo <<<<<
verificador = "nada"
c = 0

#verificando arquivos para criar pastas
#centro = []

# for i in Files:

#     base = pd.read_excel(i)
#     centro_catch = str(base.iloc[2][1])
    
#     if centro_catch in centro:
#         print(centro)
#         print("Pass centro <<<>>>>")
#         pass
#     else:
#         centro.append(centro_catch)
#         print(centro)
#         print("########")

# for i in centro:
#     os.mkdir(os.path.join(Path_name,str(i)))
    
##############################################################################
#----------------------------------------------------------------------------#

# Capturandos  os dados para classificar os arquivos

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
        
       
    
        #Preparando para transferir arquivo para a pasta correta
       # name_final = os.rename(arquivo,(os.path.join(Path_name,((tipo+" - "+transportadora+"_").upper())+(mes)+ext)))
        
        
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
        
        move(arq_move,pasta_tsp) 
        #----------------------------------------------------------
        # for i in centro:
        #     pasta_centro = os.path.join(Path_name,str(i))
        #     teste_pastacentro = os.path.exists(pasta_centro)
        #     if teste_pastacentro is True:
        #         print(pasta_centro,"<<<<<>>>>>/nTrue")
        #         print("")
        #         # if centro_catch == i:
        #         #     move(arq_move, pasta_centro)
        #     else:
        #         print("False")
        #         print(pasta_centro)
                
        # centro_catch = str(base.iloc[2][1])
        # if centro_catch == :
        #     move((arq_move), pasta_centro)        
                
        
        
        
        
        
        
         
        c+=1

    else:
        c+=1
        
#Movendo a pasta centro para o Mes centro
# move(pasta_centro,pasta_mes)
    
#----------------------------------------------------------------------------#
##############################################################################    
   
    
    
# c= 0
# for i in Files:
#
#     base = pd.read_excel(i)
#
#     head =base.columns[0]
#
#     # loop para capturar o mes
#
#         if meses[c] in head:
#             mes = meses[c]
#             print(mes)
#         else:
#             pass
#
#     for y in tipos:
#         if i in head:
#             tipo = tipos[y]
#             break
#         else:
#             tipo = tipos[y]
#
#     transportadora = base.iloc[2][0]
#
#     for z in Files:
#         os.rename(Files,(tipo+" - "+transportadora+"_").upper()+(mes))



#while True:
    #abrir arquivos e armazenar em lista

    #for vai realizar a leitura dos arquivos e determinar tipo mes transportadoras

    #for