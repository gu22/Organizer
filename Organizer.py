

#import ordenador_arqs
import datetime
import easygui
import os, glob
import os.path
from os import *

from easygui import *

import openpyxl
from openpyxl import load_workbook
import pandas as pd

from shutil import *

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
botao = easygui.msgbox("Os arquivos já estão na pasta","Escolha")

# Loop de verificação de arquivos
if botao == 'OK':
    Path_name = easygui.diropenbox()
    print(Path_name)

    Path_file = os.listdir(Path_name)
    print (Path_file)

    Files=[]
    for i in Path_file:
        Files.append(os.path.join(Path_name,i))

    print(Files)
# else:
#     Path_file = os.listdir()
#     print(Path_file)
#     Path = os.getcwd()

#     Files = []
#     for i in Path_file:
#         Files.append(os.path.join(Path, i))

    # print(Files)


nomediretorio = "BASE_Original"

diretorio = (os.path.join(Path_name,nomediretorio))

try:
    os.mkdir(diretorio)
except OSError:
    print("Pasta já existe")
else:
    print("Pasta: BASE_Original; criada com sucesso")

c=0
for g in Files:
    try:
        copyfile(g,(os.path.join(diretorio,Path_file[c])))
        print(copyfile(g,(os.path.join(diretorio,Path_file[c]))))
        c +=1
    except PermissionError:
        print ("Pasta de destino já contém arquivos com o mesmo nome, por favor mover ou excluir a pasta")

        



for i in Files:
    extensao=[]
    extensao.append(os.path.splitext(i)[1])
    print(extensao)
# fazer função que excluir outros tipos de arquivo <<<<<
verificador = "nada"
c = 0

#verificando arquivos para criar pastas
centro = []

for i in Files:

    base = pd.read_excel(i)
    centro_catch = str(base.iloc[2][1])
    
    if centro_catch in centro:
        print(centro)
        pass
    else:
        centro.append(centro_catch)
        print(centro)
        print("########")

for i in centro:
    os.mkdir(os.path.join(diretorio,str(i))

# Capturandos  os dados para classificar os arquivos
for i in Files:

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
    
    # Pegando extenção
    ext = extensao[0]
    
    #Renomeando arquivo
    os.rename(arquivo,(os.path.join(Path_name,((tipo+" - "+transportadora+"_").upper())+(mes)+ext)))
    print((Path_name+(tipo+" - "+transportadora+"_").upper()+(mes)+ext))
    print(arquivo,(os.path.join(Path_name,((tipo+" - "+transportadora+"_").upper())+(mes)+ext)))


    c+=1
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