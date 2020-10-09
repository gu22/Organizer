import os

import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox

class NewGuiApp:
    def __init__(self, master=None):
        # build ui
        self.frame_2 = ttk.Frame(master)
        self.notebook_2 = ttk.Notebook(self.frame_2)
        self.text_10 = tk.Text(self.notebook_2)
        self.text_10.config(font='TkDefaultFont', height='10', relief='flat', state='disabled')
        self.text_10.config(width='50')
        self.text_10.place(anchor='nw', bordermode='outside', height='10', width='10', x='0', y='0')
        self.notebook_2.add(self.text_10, padding='5', text='Arquivos Input')
        self.text_11 = tk.Text(self.notebook_2)
        self.text_11.config(blockcursor='false', height='10', relief='flat', width='50')
        self.text_11.pack(side='top')
        self.notebook_2.add(self.text_11, padding='5', state='disabled', text='Arquivos Output')
        self.notebook_2.config(height='200', width='300')
        self.notebook_2.place(anchor='nw', relx='0.02', rely='0.23', x='0', y='0')
        self.button_3 = ttk.Button(self.frame_2)
        self.button_3.config(text='OK')
        self.button_3.place(anchor='nw', relx='0.82', rely='0.88', x='0', y='0')
        self.button_4 = ttk.Button(self.frame_2)
        self.button_4.config(text='Selecionar Pasta')
        self.button_4.place(anchor='nw', relx='0.02', rely='0.02', x='0', y='0')
        self.button_5 = ttk.Button(self.frame_2)
        self.button_5.config(text='Prosseguir')
        self.button_5.place(anchor='nw', relx='0.65', rely='0.88', x='0', y='0')
        self.progressbar_2 = ttk.Progressbar(self.frame_2)
        self.progressbar_2.config(cursor='starting', maximum='100', orient='horizontal', takefocus=False)
        self.progressbar_2.config(value='0')
        self.progressbar_2.place(anchor='nw', relx='0.65', rely='0.8', width='160', x='0', y='0')
        self.message_2 = tk.Message(self.frame_2)
        self.message_2.config(justify='left', relief='flat', takefocus=False)
        self.message_2.place(anchor='nw', height='100', relx='0.65', rely='0.31', width='160', x='0', y='0')
        self.label_1 = ttk.Label(self.frame_2)
        self.label_1.config(text='Processo')
        self.label_1.place(anchor='nw', relx='0.65', rely='0.73', x='0', y='0')
        self.text_6 = tk.Text(self.frame_2)
        self.text_6.config(font='TkDefaultFont',height='1', relief='flat', state='disabled', takefocus=False)
        self.text_6.config(width='50')
        self.text_6.place(anchor='nw', relx='0.02', rely='0.13', width='300', x='0', y='0')
        self.text_7 = tk.Text(self.frame_2)
        self.text_7.config(blockcursor='false', font='TkDefaultFont', height='1', relief='flat')
        self.text_7.config(setgrid='false', state='disabled', width='7')
        self.text_7.place(anchor='nw', relx='0.71', rely='0.13', x='0', y='0')
        self.text_8 = tk.Text(self.frame_2)
        self.text_8.config(height='1', insertunfocussed='none', relief='flat', setgrid='true')
        self.text_8.config(state='disabled', width='5')
        self.text_8.place(anchor='nw', relx='0.87', rely='.13', x='0', y='0')
        self.label_2 = ttk.Label(self.frame_2)
        self.label_2.config(font='TkDefaultFont', justify='center', text='Arquivos\nInput')
        self.label_2.place(anchor='nw', relx='0.698', rely='0.01', x='0', y='0')
        self.label_3 = ttk.Label(self.frame_2)
        self.label_3.config(cursor='arrow', font='TkDefaultFont', justify='center', text='Arquivos\nOutput')
        self.label_3.place(anchor='nw', relx='0.86', rely='0.01', x='0', y='0')
        self.frame_2.config(height='300', relief='groove', width='500')
        self.frame_2.pack(side='top')

        # Main widget
        self.mainwindow = self.frame_2
        
        
        self.button_4.config(command = self.abrir_arquivo)
        
        

    def run(self):
        self.mainwindow.mainloop()

### --------------------------------------------------------------------
    def teste(self):
        print ("ok")
        path_name = tk.filedialog.askdirectory()
    
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
            path_name = tk.filedialog.askdirectory()
            # print(self.getfile)
            print(path_name)
            
            #self.textBrowser.setText(path_name)
            
            self.text_6.config(state="normal")
            self.text_6.insert('0.0',path_name)
            self.text_6.config(state="disabled")
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
            












if __name__ == '__main__':
    import tkinter as tk
    root = tk.Tk()
    app = NewGuiApp(root)
    app.run()

