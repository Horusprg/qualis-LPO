from turtle import width
from matplotlib.pyplot import text
import xlwt
import xlrd
import glob
from tkinter import *
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter import filedialog
from PIL import Image, ImageTk
import xml.etree.ElementTree as ET
import PyPDF2 as p2
import BaseDeCorrecoes
from xlutils.copy import copy
from tkinter import messagebox
import pandas as pd
from collections import Counter
import numpy as np


xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True

def center(win):
    # :param win: the main window or Toplevel window to center

    # Apparently a common hack to get the window size. Temporarily hide the
    # window to avoid update_idletasks() drawing the window in the wrong
    # position.
    win.update_idletasks()  # Update "requested size" from geometry manager

    # define window dimensions width and height
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width

    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width

    # Get the window position from the top dynamically as well as position from left or right as follows
    x = win.winfo_screenwidth() // 2 - win_width // 2
    y = win.winfo_screenheight() // 2 - win_height // 2

    # this is the line that will center your window
    win.geometry('+{}+{}'.format(int(x-width*1.5), int(y-height/2)))

    # This seems to draw the window frame immediately, so only call deiconify()
    # after setting correct window position
    win.deiconify()

curriculos = []

anos_validos = {'2021','2020','2019','2018','2017','2016'}


#Aplicação usando tkinter
class Application:

    def __init__(self):
        self.layout = Tk()
        self.layout.title("Gerador Qualis")
        self.layout.configure(bg="#c9e3d5")
        self.layout.resizable(False, False)
        center(self.layout)

        #Criar planilha de resultado
        self.col = 0
        notas = ['A1','A2','A3','A4','B1','B2','B3','B4','C']
        self.workbook = xlwt.Workbook()
        self.worksheet = self.workbook.add_sheet(u'Planilha_1')  #Cria aba Planilha_1
        self.worksheet.write(0, 0, u'Documento')
        self.worksheet.write(0, 1, u'Ano')
        self.worksheet.write(0, 2, u'Titulo')
        self.worksheet.write(0, 3, u'DOI')
        self.worksheet.write(0, 4, u'Sigla')
        self.worksheet.write(0, 5, u'Titulo Periodico ou Revista')
        self.worksheet.write(0, 6, u'Autores')
        self.worksheet.write(0, 7, u'Estratos')
        self.worksheet.write(0, 8, u'Notas')
        self.worksheet._cell_overwrite_ok = True

        #Cria aba Planilha_2
        self.worksheet3 = self.workbook.add_sheet(u'Planilha_2')
        self.worksheet3.write(0, self.col, u'Professor')
        self.col = self.col + 1
        self.worksheet3.write(0, self.col, u'2017')
        self.col = self.col + 1
        self.worksheet3.write(0, self.col, u'Conferência')
        self.col = self.col + 1

        for item in notas:
            self.worksheet3.write(0, self.col, item)
            self.col = self.col + 1

        self.worksheet3.write(0, self.col, u'Periódico')
        self.col = self.col + 1

        for item in notas:
            self.worksheet3.write(0, self.col, item)
            self.col = self.col + 1

        self.col = self.col + 1
        self.home()
        self.layout.mainloop()

    def home(self):
        #logo lpo
        logo1 = Image.open("assets/logo.png")
        logo1 = logo1.resize((145,80))
        logo1 = ImageTk.PhotoImage(logo1)
        self.imagem1 = Label(self.layout,
                    text = "adicionando",
                    image = logo1,
                    background="#c9e3d5")
        self.imagem1.image = logo1
        self.imagem1.grid(row=0, column=2, sticky=S, pady=10, padx=10)

        #logo ufpa
        logo2 = Image.open("assets/UFPA.png")
        logo2 = logo2.resize((78,100))
        logo2 = ImageTk.PhotoImage(logo2)
        self.imagem2 = Label(self.layout,
                    text = "adicionando",
                    image = logo2,
                    background="#c9e3d5")
        self.imagem2.image = logo2
        self.imagem2.grid(row=0, column=0, sticky=S, pady=10, padx=33)
        
        #mensagens
        self.msg1 = Label(self.layout,
                    text="GERADOR QUALIS",
                    font=("Calibri", "24", "bold"),
                    background="#c9e3d5")
        self.msg2 = Label(self.layout,
                    text="Selecione a pasta com os currículos correspondentes!",
                    font=("Calibri", "14"),
                    background="#c9e3d5")
        self.msg1.grid(row=0, column=1, sticky=S, pady=30)
        self.msg2.grid(row=1, column=1, sticky=S, pady=30, padx=30)
        
        #botão para selecionar a pasta com os currículos
        self.folder_search = Button(self.layout,
                    text="Buscar",
                    font=("Calibri", "14"),
                    width=10,
                    command=self.folder_select)
        self.folder_search.grid(row=2, column=1, sticky=S, pady=30, padx=30)
    
    #função para ler a pasta com os currículos
    def folder_select(self):
        folder_curriculos= filedialog.askdirectory()
        try:
            for f in glob.glob(folder_curriculos+"/*.xml"):
                curriculos.append(f)
        except:
            self.home()
                
        #limpar tela
        self.msg2.destroy()
        self.folder_search.destroy()
        
        #listar os curriculos
        self.listarCurriculos()

    def listarCurriculos(self):
        #mensagens
        self.msg3 = Label(self.layout,
                    text='CURRÍCULOS IMPORTADOS',
                    font=("Calibri", "16", "bold"),
                    background="#c9e3d5")
        self.msg3.grid(row=1, column=1, sticky=S, pady=30)

        #box dos curriculos
        self.curriculosImp = Listbox(self.layout,
                    background = "#a2aebc",
                    width=70,
                    height=15)
        self.curriculosImp.grid(row=2, column=1, sticky=S)

        #scroll vertical
        self.sbar_V = Scrollbar(self.layout,
                    orient = VERTICAL,
                    command=self.curriculosImp.yview)
        self.curriculosImp.configure(yscrollcommand=self.sbar_V.set,font=("Calibri", "12", "bold"))
        self.sbar_V.grid(row=2, column=1, stick=N+S+E)

        #exibe os curriculos importados
        for m in range(0, len(curriculos)):
            tree2 = ET.parse(curriculos[m])
            root2 = tree2.getroot()
            for t2 in root2.iter('DADOS-GERAIS'):                    #Imprimir nome do professor
                nomeProf2 = str(t2.attrib['NOME-COMPLETO']).upper()
                self.curriculosImp.insert(END, '{}) {}'.format(m+1, nomeProf2))
        
        
        #curriculo indesejado?
        self.msg4 = Label(self.layout,
                    background="#c9e3d5",
                    text="Se houver algum currículo indesejado, retire-o da pasta dos currículos manualmente.\nDeseja continuar?", 
                    font=("Calibri", "13", "bold"))
        self.msg4.grid(row=4, column=1, sticky=S, pady=5)
        self.msg4Y = Button(self.layout,
                    text="SIM",
                    font=("Calibri", "12"),
                    width=8,
                    command= self.lerCurriculos)
        self.msg4Y.grid(row=5, column=1, sticky=W, pady=15, padx=180)
        self.msg4N = Button(self.layout,
                    text="NÃO",
                    font=("Calibri", "12"),
                    width=8,
                    command= self.layout.destroy)
        self.msg4N.grid(row=5, column=1, sticky=E, pady=15, padx=180)
    
    def lerCurriculos(self):
        #limpar tela
        self.msg3.destroy()
        self.curriculosImp.destroy()
        self.sbar_V.destroy()
        self.msg4.destroy()
        self.msg4Y.destroy()
        self.msg4N.destroy()

        self.msg5 = Label(self.layout,
                    background="#c9e3d5",
                    text="Documento importado:\nQualisEventosComp.xls", 
                    font=("Calibri", "16", "bold"))
        self.msg5.grid(row=1, column=1, sticky=S, pady=40)

        xi = 1
        pdf = open("QUALIS_novo.pdf", "rb")                      #Script ler PDF inicio
        pdf_reader = p2.PdfFileReader(pdf)
        n = pdf_reader.numPages

        resultado_total = ['']
        for i in range(0, n):
            page = pdf_reader.getPage(i)
            pg_extraida = page.extractText().split("\n")
            resultado_total = (resultado_total + pg_extraida)     #Script ler PDF fim
		
        self.workbook2 = xlrd.open_workbook('QualisEventosComp.xls')    #Script ler xls
        self.worksheet2 = self.workbook2.sheet_by_index(1)
		
        x = 0
        somaNotas = 0

        for n in range(0, len(curriculos)):                        #Laço para ler currículos
            try:
                self.msg6.destroy()
            except:
                pass
            tree = ET.parse(curriculos[n])
            root = tree.getroot()
                
            cont = 0
            totalNota = 0
            trabalho_valido = False
            autores = ''
            conferencia = ''
            periodico = ''
            ####################################################################################
            #Contadores de Conferências por ano
            cont17c = 0
            cont18c = 0
            cont19c = 0
            cont20c = 0
            #Contadores de Periódicos por ano
            cont17p = 0
            cont18p = 0
            cont19p = 0
            cont20p = 0
            #Contadores de Nota por ano
            nota17 = 0
            nota18 = 0
            nota19 = 0
            nota20 = 0
            #Contadores de estratos por conferência em 2017
            c17A1 = 0
            c17A2 = 0
            c17A3 = 0
            c17A4 = 0
            c17B1 = 0
            c17B2 = 0
            c17B3 = 0
            c17B4 = 0
            c17C = 0
            #Contadores de estratos por periódico em 2017
            p17A1 = 0
            p17A2 = 0
            p17A3 = 0
            p17A4 = 0
            p17B1 = 0
            p17B2 = 0
            p17B3 = 0
            p17B4 = 0
            p17C = 0
            #Contadores de estratos por conferência em 2018
            c18A1 = 0
            c18A2 = 0
            c18A3 = 0
            c18A4 = 0
            c18B1 = 0
            c18B2 = 0
            c18B3 = 0
            c18B4 = 0
            c18C = 0
            #Contadores de estratos por periódico em 2018
            p18A1 = 0
            p18A2 = 0
            p18A3 = 0
            p18A4 = 0
            p18B1 = 0
            p18B2 = 0
            p18B3 = 0
            p18B4 = 0
            p18C = 0
            #Contadores de estratos por conferência em 2019
            c19A1 = 0
            c19A2 = 0
            c19A3 = 0
            c19A4 = 0
            c19B1 = 0
            c19B2 = 0
            c19B3 = 0
            c19B4 = 0
            c19C = 0
            #Contadores de estratos por periódico em 2019
            p19A1 = 0
            p19A2 = 0
            p19A3 = 0
            p19A4 = 0
            p19B1 = 0
            p19B2 = 0
            p19B3 = 0
            p19B4 = 0
            p19C = 0
            #Contadores de estratos por conferência em 2020
            c20A1 = 0
            c20A2 = 0
            c20A3 = 0
            c20A4 = 0
            c20B1 = 0
            c20B2 = 0
            c20B3 = 0
            c20B4 = 0
            c20C = 0
            #Contadores de estratos por periódico em 2020
            p20A1 = 0
            p20A2 = 0
            p20A3 = 0
            p20A4 = 0
            p20B1 = 0
            p20B2 = 0
            p20B3 = 0
            p20B4 = 0
            p20C = 0

            for t in root.iter('DADOS-GERAIS'):                    #Imprimir nome do professor
                nomeProf = str(t.attrib['NOME-COMPLETO']).upper()
                self.msg6 = Label(self.layout,
                            background="#c9e3d5",
                            text='Analisando publicações semelhantes de {}'.format(nomeProf), 
                            font=("Calibri", "16"))
                self.msg6.grid(row=2, column=1, sticky=S, pady=30)
                x = x + 2
                self.worksheet.write(x, 0, nomeProf)

            x = x + 1
            for trabalhos in root.iter('TRABALHO-EM-EVENTOS'):        #Varre currículo
                autores = ''
                trabalho_valido = False
                for trab in trabalhos.iter():                #Laço para identificar as conferências válidas
                    if trab.tag == 'DADOS-BASICOS-DO-TRABALHO' and trab.attrib['NATUREZA'] == 'COMPLETO' and trab.attrib['ANO-DO-TRABALHO'] in anos_validos:
                        conferencia = 'Conferencia;'
                        conferencia = conferencia + trab.attrib['ANO-DO-TRABALHO'] + ';' + trab.attrib['TITULO-DO-TRABALHO'] + ';' + trab.attrib['DOI'] + ';' + trab.attrib['NATUREZA']
                        trabalho_valido = True
                        cont = cont + 1
                        
                    if trabalho_valido and trab.tag == 'DETALHAMENTO-DO-TRABALHO':
                        conferencia = conferencia + ';'+ trab.attrib['NOME-DO-EVENTO'] + ';'+ trab.attrib['TITULO-DOS-ANAIS-OU-PROCEEDINGS']
                        
                    if trabalho_valido and trab.tag == 'AUTORES':
                        if autores:
                            autores = autores + '/ '+ trab.attrib['NOME-COMPLETO-DO-AUTOR']
                        else:
                            autores = trab.attrib['NOME-COMPLETO-DO-AUTOR']
                if trabalho_valido: 
                    resultado = (conferencia + ';' + autores)
                    resultado = resultado.split(";")
                    estratos = ''
                    condicao = ''
                    sigla = '-'
                    doi = str(resultado[3]).upper()
                    nomeEvento = resultado[5]
                    tituloAnais = resultado[6]
                    autor = resultado[7]
                    
                    ######################################################## Base de correção das Conferências
                    if (doi == str('10.1109/iV.2017.37').upper()):
                        estratos = 'A4'
                        condicao = '-'
                    elif(doi == str('10.1109/iV.2017.29').upper()):
                        estratos = 'A4'
                        condicao = '-'
                    elif(doi == str('10.1109/IV-2.2019.00019').upper()):
                        estratos = 'A4'
                        autor = autores #resultado[11]
                        condicao = '-'
                    elif(doi == str('10.1109/IV-2.2019.00020').upper()):
                        estratos = 'A4'
                        autor = autores #resultado[11]
                        condicao = '-'
                    elif(doi == str('10.1109/iccw.2018.8403776').upper()): #ICC Workshops
                        estratos = 'B3'
                        condicao = '-'
                    elif(doi == str('10.1145/3084226.3084278').upper()): #EASE
                        estratos = 'A3'
                        condicao = '-'
                    elif(doi == str('10.1145/3210459.3210462').upper()): #EASE
                        estratos = 'A3'
                        condicao = '-'
                    elif(doi == str('10.1109/IMOC.2017.8121084').upper()):
                        estratos = 'B4'
                        condicao = '-'
                    elif(doi == str('10.1109/icton.2017.8024977').upper()):
                        estratos = 'A4'
                        condicao = '-'
                    elif(doi == str('10.1109/IV-2.2019.00033').upper()):
                        estratos = 'A4'
                        autor = autores #resultado[7]
                        condicao = '-'
                    elif(doi == str('10.1145/3275245.3275290').upper()):
                        estratos = 'B1'
                        condicao = '-'
                    elif(str('Brazilian Symposium on Computer Networks and Distributed Systems').upper() in str(tituloAnais).upper()):
                        estratos = 'A4'
                        condicao = '-'
                    elif(str('Brazilian Symposium on Computer Networks and Distributed Systems').upper() in str(resultado[7]).upper()):
                        estratos = 'A4'
                        condicao = '-'
                    elif(str('Proceedings of the 18th Brazilian Symposium on Human Factors in Computing Systems').upper() in str(tituloAnais).upper()):
                        estratos = 'B1'
                        condicao = '-'
                    elif(str('Anais do I Workshop de Computação Urbana').upper() in str(tituloAnais).upper()):
                        estratos = 'B1'
                        condicao = '-'
                    elif(str('The 33rd ACM/SIGAPP Symposium On Applied Computing').upper() in str(tituloAnais).upper()):
                        estratos = 'A2'
                        condicao = '-'
                    elif(str('Proceedings of 2018 International Joint Conference on Neural Networks').upper() in str(tituloAnais).upper()):
                        estratos = 'A2'
                        condicao = '-'
                    
                    ########################################################
                    
                    if (condicao != '-'):
                        for row_num in range(self.worksheet2.nrows):                #Comparação por SIGLA no resultado[6]
                            if row_num == 0:
                                continue
                            row = self.worksheet2.row_values(row_num)
                            #Comparação pelo resultado[6]
                            if (' {} '.format(row[0]) in tituloAnais):
                                if (row[0] != 'SBRC'):
                                    sigla = row[0]
                                    estratos = row[8]
                                    break
                            elif ('({})'.format(row[0]) in tituloAnais):
                                sigla = row[0]
                                estratos = row[8]
                                break
                            elif ('({} '.format(row[0]) in tituloAnais):
                                sigla = row[0]
                                estratos = row[8]
                                break
                            elif ('{}&'.format(row[0]) in tituloAnais):
                                sigla = row[0]
                                estratos = row[8]
                                break
                            elif ('{}_'.format(row[0]) in tituloAnais):
                                sigla = row[0]
                                estratos = row[8]
                                break
                            elif (' {}2'.format(row[0]) in tituloAnais):
                                sigla = row[0]
                                estratos = row[8]
                                break
                                                                                #Comparação por SIGLA no resultado[5]
                            elif (' {} '.format(row[0]) in nomeEvento):
                                sigla = row[0]
                                estratos = row[8]
                                break
                            elif ('({})'.format(row[0]) in nomeEvento):
                                sigla = row[0]
                                estratos = row[8]
                                break
                            elif ('({} '.format(row[0]) in nomeEvento):
                                sigla = row[0]
                                estratos = row[8]
                                break
                            elif ('{}&'.format(row[0]) in nomeEvento):
                                sigla = row[0]
                                estratos = row[8]
                                break
                            elif ('{}_'.format(row[0]) in nomeEvento):
                                sigla = row[0]
                                estratos = row[8]
                                break
                            elif (' {}2'.format(row[0]) in nomeEvento):
                                sigla = row[0]
                                estratos = row[8]
                                break
                            elif ('XVII {}'.format(row[0]) in str(nomeEvento).upper()):
                                sigla = row[0]
                                estratos = row[8]
                                break
                            elif ('({})'.format(row[0]) in resultado[7]):
                                sigla = row[0]
                                estratos = row[8]
                                break
                            else:
                                sigla = '-'
                                estratos = '-'
                        
                        
                        for row_num in range(self.worksheet2.nrows):                       #Comparação por nome
                            if row_num == 0:
                                continue
                            row = self.worksheet2.row_values(row_num)
                            if (estratos == '-'):
                                if (str(row[1]).upper() in str(resultado[6]).upper()):
                                    sigla = row[0]
                                    estratos = row[8]
                                    break
                                elif (row[1] in resultado[5]):
                                    sigla = row[0]
                                    estratos = row[8]
                                    break
                                elif (row[1] in resultado[7]):
                                    sigla = row[0]
                                    estratos = row[8]
                                    break
                            
                        for row_num in range(self.worksheet2.nrows):                #Comparação por SIGLA casos especiais
                            if row_num == 0:
                                continue
                            row = self.worksheet2.row_values(row_num)
                            if (estratos == '-'):
                                if (" ({}'2019)".format(row[0]) in resultado[6]):
                                    sigla = row[0]
                                    estratos = row[8]
                                    break
                                elif ("{}'18 ".format(row[0]) in resultado[6] and row[0] != 'ER'):
                                    sigla = row[0]
                                    estratos = row[8]
                                    break
                    
                    self.worksheet.write(x, 0, resultado[0])
                    self.worksheet.write(x, 1, resultado[1])
                    self.worksheet.write(x, 4, sigla)
                    if ('COMPLETO' in tituloAnais):                          #Correção de tabela, elimina o "COMPLETO" do lugar errado
                        self.worksheet.write(x, 2, resultado[2] + resultado [3] + resultado[4])
                        self.worksheet.write(x, 3, resultado[5])
                        self.worksheet.write(x, 5, resultado[8] + ' / ' + autor)
                        self.worksheet.write(x, 6, resultado[9])
                    elif ('COMPLETO' in nomeEvento):
                        self.worksheet.write(x, 2, resultado[2] + resultado[3])
                        self.worksheet.write(x, 3, resultado[4])
                        self.worksheet.write(x, 5, autor + ' / ' + resultado[6])
                        self.worksheet.write(x, 6, resultado[8])
                    else:
                        self.worksheet.write(x, 2, resultado[2])
                        if (resultado[3] != ''):
                            self.worksheet.write(x, 3, resultado[3])
                        else:
                            self.worksheet.write(x, 3, '-')
                        self.worksheet.write(x, 5, tituloAnais + ' / ' + nomeEvento)
                        if (len(resultado) > 8):
                            if (nomeProf in str(autor).upper()):
                                self.worksheet.write(x, 6, autor)
                            elif (nomeProf in str(resultado[8]).upper()):
                                self.worksheet.write(x, 6, resultado[8])
                        else:
                            self.worksheet.write(x, 6, autor)
                    self.worksheet.write(x, 7, estratos)
                    
                    nota = 'SEM QUALIS'             #Calcula a nota do estrato
                    if (estratos == 'A1'):
                        nota = BaseDeCorrecoes.A1c
                    elif (estratos == 'A2'):
                        nota = BaseDeCorrecoes.A2c
                    elif (estratos == 'A3'):
                        nota = BaseDeCorrecoes.A3c
                    elif (estratos == 'A4'):
                        nota = BaseDeCorrecoes.A4c
                    elif (estratos == 'B1'):
                        nota = BaseDeCorrecoes.B1c
                    elif (estratos == 'B2'):
                        nota = BaseDeCorrecoes.B2c
                    elif (estratos == 'B3'):
                        nota = BaseDeCorrecoes.B3c
                    elif (estratos == 'B4'):
                        nota = BaseDeCorrecoes.B4c
                    elif (estratos == 'C'):
                        nota = BaseDeCorrecoes.Cc
                    self.worksheet.write(x, 8, nota)
                    
                    if (nota != 'SEM QUALIS'):                  #Contador de estratos das conferências
                        totalNota = totalNota + nota
                    if (estratos != '-'):
                        if (resultado[1] == '2017'):
                            cont17c = cont17c + 1
                            if (nota != 'SEM QUALIS'):          #somador de notas de 2017
                                nota17 = nota17 + nota
                            if (estratos == 'A1'):
                                c17A1 = c17A1 + 1
                            elif (estratos == 'A2'):
                                c17A2 = c17A2 + 1
                            elif (estratos == 'A3'):
                                c17A3 = c17A3 + 1
                            elif (estratos == 'A4'):
                                c17A4 = c17A4 + 1
                            elif (estratos == 'B1'):
                                c17B1 = c17B1 + 1
                            elif (estratos == 'B2'):
                                c17B2 = c17B2 + 1
                            elif (estratos == 'B3'):
                                c17B3 = c17B3 + 1
                            elif (estratos == 'B4'):
                                c17B4 = c17B4 + 1
                            elif (estratos == 'C'):
                                c17C = c17C + 1
                        elif (resultado[1] == '2018'):
                            cont18c = cont18c + 1
                            if (nota != 'SEM QUALIS'):          #somador de notas de 2018
                                nota18 = nota18 + nota
                            if (estratos == 'A1'):
                                c18A1 = c18A1 + 1
                            elif (estratos == 'A2'):
                                c18A2 = c18A2 + 1
                            elif (estratos == 'A3'):
                                c18A3 = c18A3 + 1
                            elif (estratos == 'A4'):
                                c18A4 = c18A4 + 1
                            elif (estratos == 'B1'):
                                c18B1 = c18B1 + 1
                            elif (estratos == 'B2'):
                                c18B2 = c18B2 + 1
                            elif (estratos == 'B3'):
                                c18B3 = c18B3 + 1
                            elif (estratos == 'B4'):
                                c18B4 = c18B4 + 1
                            elif (estratos == 'C'):
                                c18C = c18C + 1
                        elif (resultado[1] == '2019'):
                            cont19c = cont19c + 1
                            if (nota != 'SEM QUALIS'):          #somador de notas de 2019
                                nota19 = nota19 + nota
                            if (estratos == 'A1'):
                                c19A1 = c19A1 + 1
                            elif (estratos == 'A2'):
                                c19A2 = c19A2 + 1
                            elif (estratos == 'A3'):
                                c19A3 = c19A3 + 1
                            elif (estratos == 'A4'):
                                c19A4 = c19A4 + 1
                            elif (estratos == 'B1'):
                                c19B1 = c19B1 + 1
                            elif (estratos == 'B2'):
                                c19B2 = c19B2 + 1
                            elif (estratos == 'B3'):
                                c19B3 = c19B3 + 1
                            elif (estratos == 'B4'):
                                c19B4 = c19B4 + 1
                            elif (estratos == 'C'):
                                c19C = c19C + 1
                        elif (resultado[1] == '2020'):
                            cont20c = cont20c + 1
                            if (nota != 'SEM QUALIS'):          #somador de notas de 2020
                                nota20 = nota20 + nota
                            if (estratos == 'A1'):
                                c20A1 = c20A1 + 1
                            elif (estratos == 'A2'):
                                c20A2 = c20A2 + 1
                            elif (estratos == 'A3'):
                                c20A3 = c20A3 + 1
                            elif (estratos == 'A4'):
                                c20A4 = c20A4 + 1
                            elif (estratos == 'B1'):
                                c20B1 = c20B1 + 1
                            elif (estratos == 'B2'):
                                c20B2 = c20B2 + 1
                            elif (estratos == 'B3'):
                                c20B3 = c20B3 + 1
                            elif (estratos == 'B4'):
                                c20B4 = c20B4 + 1
                            elif (estratos == 'C'):
                                c20C = c20C + 1
                        
                    x = x + 1
        
            for trabalhos in root.iter('ARTIGO-PUBLICADO'):           #Varrer currículo
                autores = ''
                trabalho_valido = False
                for trab in trabalhos.iter():        #Laço para identificar os periódicos válidos
                    if trab.tag == 'DADOS-BASICOS-DO-ARTIGO' and trab.attrib['NATUREZA'] == 'COMPLETO' and trab.attrib['ANO-DO-ARTIGO'] in anos_validos:
                        periodico = 'Periodico;'
                        periodico = periodico + trab.attrib['ANO-DO-ARTIGO'] + ';'+ trab.attrib['TITULO-DO-ARTIGO'] +';' + trab.attrib['DOI'] +';' + trab.attrib['NATUREZA']
                        trabalho_valido = True
                        cont = cont + 1
                        
                    if trabalho_valido and trab.tag == 'DETALHAMENTO-DO-ARTIGO':
                        periodico = periodico + ';'+ trab.attrib['TITULO-DO-PERIODICO-OU-REVISTA']
                        
                    if trabalho_valido and trab.tag == 'AUTORES':
                        if autores: 
                            autores = autores + '/ '+ trab.attrib['NOME-COMPLETO-DO-AUTOR']
                        else:
                            autores = trab.attrib['NOME-COMPLETO-DO-AUTOR']
                if trabalho_valido:
                    resultado2 = (periodico + ';' + autores)
                    resultado2 = resultado2.split(";")
                    estratos2 = ''
                    doi = str(resultado2[3]).upper()
                    ##################################################### Base de correção dos Periódicos
                    if(doi == str('10.14209/jcis.2019.22').upper()):
                        estratos2 = 'A4'
                    elif(doi == str('10.1155/2017/2865482').upper()):
                        estratos2 = 'B1'
                    elif(doi == str('10.1177/1475921718799070').upper()):
                        estratos2 = 'A1'
                    elif(doi == str('10.1007/s00530-015-0501-6').upper()):
                        estratos2 = 'A2'
                    elif(doi == str('10.1016/j.compenvurbsys.2017.05.001').upper()):
                        estratos2 = 'A1'
                    elif(doi == str('10.1002/spe.2637').upper()):
                        estratos2 = 'A3'
                    elif(doi == str('10.1177/1475921718799070').upper()):
                        estratos2 = 'A1'
                    elif(doi == str('10.1590/0074-02760170111').upper()):
                        estratos2 = 'A2'
                    elif(doi == str('10.1002/nem.2055').upper()):
                        estratos2 = 'A4'
                    elif(str('REVISTA DA ABET').upper() in str(resultado2[5]).upper()):
                        estratos2 = 'A4'
                    elif(str('Journal of Communication and Information Systems').upper() in str(resultado2[5]).upper()):
                        estratos2 = 'A4' 
                    #######################################################
                        
                    if (estratos2 == ''):
                        for i in range(0,len(resultado_total)):                   #Comparação por nome
                            #nomePeriodico = str(resultado2[5]).upper()
                            if (str(resultado2[5]).upper() in resultado_total[i]):
                                if (' {} '.format(str(resultado2[5]).upper()) in resultado_total[i]):
                                    continue

                                if (len(resultado2[5]) == len(resultado_total[i])):
                                    estratos2 = resultado_total[i+1]
                                    break
                                elif (len(resultado2[5]) < len(resultado_total[i])):
                                    if ('{} (PRINT)'.format(str(resultado2[5]).upper()) == resultado_total[i]):
                                        estratos2 = resultado_total[i+1]
                                        break
                                    if ('{} (ONLINE)'.format(str(resultado2[5]).upper()) == resultado_total[i]):
                                        estratos2 = resultado_total[i+1]
                                        break
                                    elif ('ACS {}'.format(str(resultado2[5]).upper()) == resultado_total[i]):
                                        estratos2 = resultado_total[i+1]
                                        break
                                    elif ('THE {}'.format(str(resultado2[5]).upper()) == resultado_total[i]):
                                        estratos2 = resultado_total[i+1]
                                        break
                                    elif ('{} (19'.format(str(resultado2[5]).upper()) in resultado_total[i]):
                                        estratos2 = resultado_total[i+1]
                                        break
                                    elif (len(resultado_total[i]) - len(resultado2[5]) <= 12):
                                        if ('{} ('.format(str(resultado2[5]).upper()) in resultado_total[i]):
                                            estratos2 = resultado_total[i+1]
                                            break
                                        else:
                                            same = messagebox.askquestion("Verificação", resultado2[5] + ' é o mesmo que ' + resultado_total[i] + '?')
                                            resp = False
                                            while (resp == False):
                                                if (same == 'yes'):
                                                    estratos2 = resultado_total[i+1]
                                                    resp = True
                                                    break
                                                elif (same == 'no'):
                                                    estratos2 = '-'
                                                    resp = True

                                    elif (len(resultado_total[i]) - len(resultado2[5]) > 12):
                                        if ('{} ('.format(str(resultado2[5]).upper()) in resultado_total[i]):
                                            same = messagebox.askquestion("Verificação", resultado2[5] + ' é o mesmo que ' + resultado_total[i] + '?')
                                            resp = False
                                            while (resp == False):
                                                if (same == 'yes'):
                                                    estratos2 = resultado_total[i+1]
                                                    resp = True
                                                    break
                                                elif (same == 'no'):
                                                        estratos2 = '-'
                                                        resp = True
                            
                            elif (str(resultado2[6]).upper() in resultado_total[i]):
                                estratos2 = resultado_total[i+1]
                                break
                            else:
                                estratos2 = '-'
                    
                    
                    self.worksheet.write(x, 0, resultado2[0])
                    self.worksheet.write(x, 1, resultado2[1])
                    self.worksheet.write(x, 4, '-')
                    if ('COMPLETO' in resultado2[5]):                        #Correção de tabela, elimina o "COMPLETO" do lugar errado
                        self.worksheet.write(x, 2, resultado2[2] + resultado2[3])
                        self.worksheet.write(x, 3, resultado2[4])
                        self.worksheet.write(x, 5, resultado2[6])
                        self.worksheet.write(x, 6, resultado2[7])
                    else:
                        self.worksheet.write(x, 2, resultado2[2])
                        if (resultado2[3] != ''):
                            self.worksheet.write(x, 3, resultado2[3])
                        else:
                            self.worksheet.write(x, 3, '-')
                        self.worksheet.write(x, 5, resultado2[5])
                        self.worksheet.write(x, 6, resultado2[6])
                    self.worksheet.write(x, 7, estratos2)
                    
                    nota = 'SEM QUALIS'               #Calcula nota do estrato
                    if (estratos2 == 'A1'):
                        nota = BaseDeCorrecoes.A1p
                    elif (estratos2 == 'A2'):
                        nota = BaseDeCorrecoes.A2p
                    elif (estratos2 == 'A3'):
                        nota = BaseDeCorrecoes.A3p
                    elif (estratos2 == 'A4'):
                        nota = BaseDeCorrecoes.A4p
                    elif (estratos2 == 'B1'):
                        nota = BaseDeCorrecoes.B1p
                    elif (estratos2 == 'B2'):
                        nota = BaseDeCorrecoes.B2p
                    elif (estratos2 == 'B3'):
                        nota = BaseDeCorrecoes.B3p
                    elif (estratos2 == 'B4'):
                        nota = BaseDeCorrecoes.B4p
                    elif (estratos2 == 'C'):
                        nota = BaseDeCorrecoes.Cp
                    self.worksheet.write(x, 8, nota)
                    
                    if (nota != 'SEM QUALIS'):            #Contador de estratos dos periódicos
                        totalNota = totalNota + nota
                    if (estratos2 != '-'):
                        if (resultado2[1] == '2017'):
                            cont17p = cont17p + 1
                            if (nota != 'SEM QUALIS'):          #somador de notas de 2017
                                nota17 = nota17 + nota
                            if (estratos2 == 'A1'):
                                p17A1 = p17A1 + 1
                            elif (estratos2 == 'A2'):
                                p17A2 = p17A2 + 1
                            elif (estratos2 == 'A3'):
                                p17A3 = p17A3 + 1
                            elif (estratos2 == 'A4'):
                                p17A4 = p17A4 + 1
                            elif (estratos2 == 'B1'):
                                p17B1 = p17B1 + 1
                            elif (estratos2 == 'B2'):
                                p17B2 = p17B2 + 1
                            elif (estratos2 == 'B3'):
                                p17B3 = p17B3 + 1
                            elif (estratos2 == 'B4'):
                                p17B4 = p17B4 + 1
                            elif (estratos2 == 'C'):
                                p17C = p17C + 1
                        elif (resultado2[1] == '2018'):
                            cont18p = cont18p + 1
                            if (nota != 'SEM QUALIS'):          #somador de notas de 2018
                                nota18 = nota18 + nota
                            if (estratos2 == 'A1'):
                                p18A1 = p18A1 + 1
                            elif (estratos2 == 'A2'):
                                p18A2 = p18A2 + 1
                            elif (estratos2 == 'A3'):
                                p18A3 = p18A3 + 1
                            elif (estratos2 == 'A4'):
                                p18A4 = p18A4 + 1
                            elif (estratos2 == 'B1'):
                                p18B1 = p18B1 + 1
                            elif (estratos2 == 'B2'):
                                p18B2 = p18B2 + 1
                            elif (estratos2 == 'B3'):
                                p18B3 = p18B3 + 1
                            elif (estratos2 == 'B4'):
                                p18B4 = p18B4 + 1
                            elif (estratos2 == 'C'):
                                p18C = p18C + 1
                        elif (resultado2[1] == '2019'):
                            cont19p = cont19p + 1
                            if (nota != 'SEM QUALIS'):          #somador de notas de 2019
                                nota19 = nota19 + nota
                            if (estratos2 == 'A1'):
                                p19A1 = p19A1 + 1
                            elif (estratos2 == 'A2'):
                                p19A2 = p19A2 + 1
                            elif (estratos2 == 'A3'):
                                p19A3 = p19A3 + 1
                            elif (estratos2 == 'A4'):
                                p19A4 = p19A4 + 1
                            elif (estratos2 == 'B1'):
                                p19B1 = p19B1 + 1
                            elif (estratos2 == 'B2'):
                                p19B2 = p19B2 + 1
                            elif (estratos2 == 'B3'):
                                p19B3 = p19B3 + 1
                            elif (estratos2 == 'B4'):
                                p19B4 = p19B4 + 1
                            elif (estratos2 == 'C'):
                                p19C = p19C + 1
                        elif (resultado2[1] == '2020'):
                            cont20p = cont20p + 1
                            if (nota != 'SEM QUALIS'):          #somador de notas de 2020
                                nota20 = nota20 + nota
                            if (estratos2 == 'A1'):
                                p20A1 = p20A1 + 1
                            elif (estratos2 == 'A2'):
                                p20A2 = p20A2 + 1
                            elif (estratos2 == 'A3'):
                                p20A3 = p20A3 + 1
                            elif (estratos2 == 'A4'):
                                p20A4 = p20A4 + 1
                            elif (estratos2 == 'B1'):
                                p20B1 = p20B1 + 1
                            elif (estratos2 == 'B2'):
                                p20B2 = p20B2 + 1
                            elif (estratos2 == 'B3'):
                                p20B3 = p20B3 + 1
                            elif (estratos2 == 'B4'):
                                p20B4 = p20B4 + 1
                            elif (estratos2 == 'C'):
                                p20C = p20C + 1
                        
                    x = x + 1
            self.worksheet.write(x, 7, 'Nota Total')
            self.worksheet.write(x, 8, totalNota)
            contTotalc = cont17c
            contTotalp = cont17p
            totalNota = nota17
                
            #Planilha_2
            yi = 0
            if (xi <= len(curriculos)):
                self.worksheet3.write(xi, yi, nomeProf)
                yi = yi + 2
            
                self.worksheet3.write(xi, yi, cont17c)
                yi = yi + 1
                self.worksheet3.write(xi, yi, c17A1)
                yi = yi + 1
                self.worksheet3.write(xi, yi, c17A2)
                yi = yi + 1
                self.worksheet3.write(xi, yi, c17A3)
                yi = yi + 1
                self.worksheet3.write(xi, yi, c17A4)
                yi = yi + 1
                self.worksheet3.write(xi, yi, c17B1)
                yi = yi + 1
                self.worksheet3.write(xi, yi, c17B2)
                yi = yi + 1
                self.worksheet3.write(xi, yi, c17B3)
                yi = yi + 1
                self.worksheet3.write(xi, yi, c17B4)
                yi = yi + 1
                self.worksheet3.write(xi, yi, c17C)
                yi = yi + 1
                self.worksheet3.write(xi, yi, cont17p)
                yi = yi + 1
                self.worksheet3.write(xi, yi, p17A1)
                yi = yi + 1
                self.worksheet3.write(xi, yi, p17A2)
                yi = yi + 1
                self.worksheet3.write(xi, yi, p17A3)
                yi = yi + 1
                self.worksheet3.write(xi, yi, p17A4)
                yi = yi + 1
                self.worksheet3.write(xi, yi, p17B1)
                yi = yi + 1
                self.worksheet3.write(xi, yi, p17B2)
                yi = yi + 1
                self.worksheet3.write(xi, yi, p17B3)
                yi = yi + 1
                self.worksheet3.write(xi, yi, p17B4)
                yi = yi + 1
                self.worksheet3.write(xi, yi, p17C)
                yi = yi + 3
                
                yi = yi - 1
                self.worksheet3.write(xi, yi, contTotalc)
                yi = yi + 1
                self.worksheet3.write(xi, yi, contTotalp)
                yi = yi + 1
                self.worksheet3.write(xi, yi, totalNota)
                    
                somaNotas = somaNotas + totalNota
                xi = xi + 1

        try:
            self.msg6.destroy()
        except:
            pass
        self.worksheet3.write(0, self.col, u'Total Conferências')
        self.col = self.col + 1
        self.worksheet3.write(0, self.col, u'Total Periódicos')
        self.col = self.col + 1
        self.worksheet3.write(0, self.col, u'Pontuação Total')

        mediaNotas = (somaNotas/len(curriculos))
        self.worksheet3.write(xi+1, yi-1, 'SOMA')
        #self.worksheet3.write(xi+1, 70, somaNotas)
        #self.worksheet3.write(xi+1, 71, '=SOMA(BS2:BS23)')#As fórmulas ficam apenas como texto, precisa clicar na célula e apertar "enter"
        self.worksheet3.write(xi+2, yi-1, 'MÉDIA')
        #self.worksheet3.write(xi+2, 70, mediaNotas)
        #self.worksheet3.write(xi+2, 71, '=SOMA(BS2:BS23)/len(curriculos)')

        self.file = filedialog.askdirectory(title="Selecione um local para salvar a planilha produzida!")
        self.file = self.file+'\\Resultado.xls'
        self.workbook.save(self.file)#salva em arquivo xls
        ###############################################################
        #
        #############################INTEGRAÇÃO DO PROGRAMA DE CORREÇÃO
        #
        ###############################################################
        def decSim():
            self.msg5.destroy()
            self.decisao.destroy()
            self.decisaoY.destroy()
            self.decisaoN.destroy()

            rb = xlrd.open_workbook(self.file)        #Ler arquivo para fazer cópia
            wb = copy(rb)
            
            lista = []
            self.workbook = xlrd.open_workbook(self.file)  #Carrega arquivo para leitura
            self.worksheet = self.workbook.sheet_by_index(0)
            for row_num in range(self.worksheet.nrows):
                if row_num == 0:
                    continue
                row = self.worksheet.row_values(row_num)
                lista = lista + row
            
            #Base de correção para os que não foram reconhecidos
            lista.append(BaseDeCorrecoes.listaBase)
            totalNotas2 = 0
            somaNotas2 = 0
            nota172 = 0
            nota182 = 0
            nota192 = 0
            nota202 = 0
            nt = 1
            for row_num in range(self.worksheet.nrows):     #Varre linha por linha do NotasExtraídas
                w_sheet = wb.get_sheet(0)
                if row_num == 0:
                    continue
                row = self.worksheet.row_values(row_num)

                if (row[8] != 'SEM QUALIS'  and row[1] != ''):
                    novaNota2 = row[8]
                    cont = (str(lista).upper()).count(str(row[2]).upper())
                    if (cont > 1):
                        novaNota2 = row[8]/cont
                        w_sheet.write(row_num, 8, novaNota2)
                        #print (novaNota2)
                        w_sheet.write(row_num, 9, cont)
                        #print (cont)
                    
                    totalNotas2 = totalNotas2 + novaNota2
                    if (row[1] == '2017'):
                        nota172 = nota172 + novaNota2
                    elif (row[1] == '2018'):
                        nota182 = nota182 + novaNota2
                    elif (row[1] == '2019'):
                        nota192 = nota192 + novaNota2
                    elif (row[1] == '2020'):
                        nota202 = nota202 + novaNota2
                    
                if (row[7] == 'Nota Total'):
                    w_sheet.write(row_num, 8, totalNotas2)
                    
                    w_sheet = wb.get_sheet(1)
                    totalNotas2 = nota172

                    somaNotas2 = somaNotas2 + totalNotas2
                    w_sheet.write(nt, yi, totalNotas2)
                    totalNotas2 = 0
                    nota172 = 0
                    nota182 = 0
                    nota192 = 0
                    nota202 = 0
                    nt = nt + 1

            w_sheet = wb.get_sheet(1)
            mediaNotas2 = somaNotas2/len(curriculos)
            w_sheet.write(nt+1, yi, somaNotas2)
            w_sheet.write(nt+2, yi, mediaNotas2)
                    
            wb.save(self.file)
            self.msgf = Label(self.layout,
                        background="#c9e3d5",
                        text='NOTAS CORRIGIDAS! \nPara conferir a planilha com os resultados, consulte o arquivo Resultados.xls.', 
                        font=("Calibri", "13", "bold"))
            self.msgf.grid(row=3, column=1, sticky=S, pady=5)
            self.map = Label(self.layout,
                    background="#c9e3d5",
                    text='DESEJA MAPEAR O EXTRATO PRODUZIDO?', 
                    font=("Calibri", "13", "bold"))
            self.map.grid(row=4, column=1, sticky=S, pady=5)
            self.mapY = Button(self.layout,
                        text="SIM",
                        font=("Calibri", "12"),
                        width=8,
                        command= lambda: [self.mapear(),
                                self.map.destroy(),
                                self.mapY.destroy(),
                                self.mapN.destroy(),
                                self.msgf.destroy()])

            self.mapY.grid(row=5, column=1, sticky=W, pady=15, padx=180)
            self.mapN = Button(self.layout,
                        text="NÃO",
                        font=("Calibri", "12"),
                        width=8,
                        command= self.layout.destroy)
            self.mapN.grid(row=5, column=1, sticky=E, pady=15, padx=180)
            
        def decNao():
            self.msg5.destroy()
            self.decisao.destroy()
            self.decisaoY.destroy()
            self.decisaoN.destroy()

            self.map = Label(self.layout,
                    background="#c9e3d5",
                    text='DESEJA MAPEAR O EXTRATO PRODUZIDO?', 
                    font=("Calibri", "13", "bold"))
            self.map.grid(row=4, column=1, sticky=S, pady=5)
            self.map.grid(row=4, column=1, sticky=S, pady=5)
            self.mapY = Button(self.layout,
                        text="SIM",
                        font=("Calibri", "12"),
                        width=8,
                        command= lambda: [self.mapear(),
                                self.map.destroy(),
                                self.mapY.destroy(),
                                self.mapN.destroy(),
                                self.msgf.destroy()])

            self.mapY.grid(row=5, column=1, sticky=W, pady=15, padx=180)
            self.mapN = Button(self.layout,
                        text="NÃO",
                        font=("Calibri", "12"),
                        width=8,
                        command= self.layout.destroy)
            self.mapN.grid(row=5, column=1, sticky=E, pady=15, padx=180)

        self.decisao = Label(self.layout,
                    background="#c9e3d5",
                    text='DESEJA APLICAR A CORREÇÃO DE NOTAS?\nA correção é a divisão de notas para uma publicação que está no currículo de mais de um professor.', 
                    font=("Calibri", "13", "bold"))
        self.decisao.grid(row=3, column=1, sticky=S, pady=5)
        self.decisaoY = Button(self.layout,
                    text="SIM",
                    font=("Calibri", "12"),
                    width=8,
                    command= decSim)
        self.decisaoY.grid(row=4, column=1, sticky=W, pady=15, padx=180)
        self.decisaoN = Button(self.layout,
                    text="NÃO",
                    font=("Calibri", "12"),
                    width=8,
                    command= decNao)
        self.decisaoN.grid(row=4, column=1, sticky=E, pady=15, padx=180)
    
    def mapear(self):
        self.final = Label(self.layout,
                    background="#c9e3d5",
                    text='MAPEAMENTO REALIZADO!', 
                    font=("Calibri", "13", "bold"))
        self.final.grid(row=2, column=1, sticky=S, pady=5)

        df = pd.read_excel(self.file)
        df2 = pd.read_csv('qualis.csv')
        R = list(Counter(df['Titulo Periodico ou Revista']))[1:]

        dic = {}

        for revista in R:
            idxrevista = df2[df2['Título'] == revista.upper()].index
            
            if len(idxrevista) > 0:
                estratoantigo = df2.loc[idxrevista]['Estrato'][idxrevista[0]]
            else:
                estratoantigo = np.nan
                
            dic[revista] = estratoantigo

        for revista in dic:
            idxs = df[df["Titulo Periodico ou Revista"] == revista].index
            df.loc[idxs, "Estrato Antigo"] = dic[revista]

        df.to_excel(self.file)

Application()