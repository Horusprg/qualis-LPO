import xlwt
import xlrd
import xml.etree.ElementTree as ET
import PyPDF2 as p2
from xlutils.copy import copy
import glob
import BaseDeCorrecoes
xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True

curriculos = []
for f in glob.glob('curriculos/*.xml'):    #Lista os arquivos xml do mesmo diretório que o programa
	curriculos.append(f)        #Adiciona o nome de cada arquivo xml ao final da lista curriculos[]

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet(u'Planilha_1')  #Cria aba Planilha_1
worksheet.write(0, 0, u'Documento')
worksheet.write(0, 1, u'Ano')
worksheet.write(0, 2, u'Titulo')
worksheet.write(0, 3, u'DOI')
worksheet.write(0, 4, u'Sigla')
worksheet.write(0, 5, u'Titulo Periodico ou Revista')
worksheet.write(0, 6, u'Autores')
worksheet.write(0, 7, u'Estratos')
worksheet.write(0, 8, u'Notas')



#anos_validos = { '2017', '2018', '2019', '2020'}
anos_validos = {'2021','2020','2019','2018','2017','2016'}

#Define intervalo de tempo desejado
"""
	respAno = input('''Qual ano você gostaria de analisar? \n
1) TODOS
2) 2017
3) 2018
4) 2019
5) 2020

Digite o número correspondente à sua escolha. \n''')
"""
respAno = '2'
col = 0
todos = False
r15 = False
r16 = False
r17 = False
r18 = False
r19 = False
r20 = False
r = False

while (r == False):
	if (respAno == '2'):
		r17 = True
		r = True
	elif (respAno == '3'):
		r18 = True
		r = True
	elif (respAno == '4'):
		r19 = True
		r = True
	elif (respAno == '5'):
		r20 = True
		r = True
	elif (respAno == '1'):
		r15 = True
		r16 = True
		r17 = True
		r18 = True
		r19 = True
		r20 = True
		r = True
	else:
		input('Número inválido. Digite o número correspondente à alternativa desejada.')

worksheet3 = workbook.add_sheet(u'Planilha_2')  #Cria aba Planilha_2
worksheet3.write(0, col, u'Professor')
col = col + 1

if (r17 == True):                                   #Para definir qual cabeçalho vai aparecer
	worksheet3.write(0, col, u'2017')
	col = col + 1
	worksheet3.write(0, col, u'Conferência')
	col = col + 1
	worksheet3.write(0, col, u'A1')
	col = col + 1
	worksheet3.write(0, col, u'A2')
	col = col + 1
	worksheet3.write(0, col, u'A3')
	col = col + 1
	worksheet3.write(0, col, u'A4')
	col = col + 1
	worksheet3.write(0, col, u'B1')
	col = col + 1
	worksheet3.write(0, col, u'B2')
	col = col + 1
	worksheet3.write(0, col, u'B3')
	col = col + 1
	worksheet3.write(0, col, u'B4')
	col = col + 1
	worksheet3.write(0, col, u'C')
	col = col + 1
	worksheet3.write(0, col, u'Periódico')
	col = col + 1
	worksheet3.write(0, col, u'A1')
	col = col + 1
	worksheet3.write(0, col, u'A2')
	col = col + 1
	worksheet3.write(0, col, u'A3')
	col = col + 1
	worksheet3.write(0, col, u'A4')
	col = col + 1
	worksheet3.write(0, col, u'B1')
	col = col + 1
	worksheet3.write(0, col, u'B2')
	col = col + 1
	worksheet3.write(0, col, u'B3')
	col = col + 1
	worksheet3.write(0, col, u'B4')
	col = col + 1
	worksheet3.write(0, col, u'C')
	col = col + 2

print('\nCurrículos importados:')                          #Imprime os currículos que serão importados
for m in range(0, len(curriculos)):
	tree2 = ET.parse(curriculos[m])
	root2 = tree2.getroot()
	for t2 in root2.iter('DADOS-GERAIS'):                    #Imprimir nome do professor
		nomeProf2 = str(t2.attrib['NOME-COMPLETO']).upper()
		print('{}) {}'.format(m+1, nomeProf2))


'''
#
F
I
N
A
L
I
Z
A
D
O
#
'''


desejo = input('\nSe houver algum currículo indesejado, retire-o da pasta do programa. Deseja continuar? (S/N) \n')
desejo2 = False
while (desejo2 == False):
	if (desejo == 's' or desejo == 'S'):
		
		xi = 1
		print('\nImportanto documento: QUALIS_novo.pdf...')
		pdf = open("QUALIS_novo.pdf", "rb")                      #Script ler PDF inicio
		pdf_reader = p2.PdfFileReader(pdf)
		n = pdf_reader.numPages
		
		resultado_total = ['']
		for i in range(0, n):
			page = pdf_reader.getPage(i)
			pg_extraida = page.extractText().split("\n")
			resultado_total = (resultado_total + pg_extraida)     #Script ler PDF fim
		
		print('Importanto documento: QualisEventosComp.xls...')
		workbook2 = xlrd.open_workbook('QualisEventosComp.xls')    #Script ler xls
		worksheet2 = workbook2.sheet_by_index(1)
		
		x = 0
		somaNotas = 0
		
		print('Lendo currículo(s)... \n')
		for n in range(0, len(curriculos)):                        #Laço para ler currículos
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
			##################################################################################
			
			for t in root.iter('DADOS-GERAIS'):                    #Imprimir nome do professor
				nomeProf = str(t.attrib['NOME-COMPLETO']).upper()
				print('Analisando publicações de {}'.format(nomeProf))
				x = x + 2
				worksheet.write(x, 0, nomeProf)
		
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
						for row_num in range(worksheet2.nrows):                #Comparação por SIGLA no resultado[6]
							if row_num == 0:
								continue
							row = worksheet2.row_values(row_num)
							#Comparação pelo resultado[6]
							if (' {} '.format(row[0]) in tituloAnais):
								if (row[0] != 'SBRC'):
									sigla = row[0]
		#							print(row[8] + ' --> ' + sigla)# + ' # 6 #')
									estratos = row[8]
									break
							elif ('({})'.format(row[0]) in tituloAnais):
								sigla = row[0]
		#						print(row[8] + ' --> ' + sigla)# + ' #(6)#')
								estratos = row[8]
								break
							elif ('({} '.format(row[0]) in tituloAnais):
								sigla = row[0]
		#						print(row[8] + ' --> ' + sigla)# + ' #(6 #')
								estratos = row[8]
								break
							elif ('{}&'.format(row[0]) in tituloAnais):
								sigla = row[0]
		#						print(row[8] + ' --> ' + sigla)# + ' #6&#')
								estratos = row[8]
								break
							elif ('{}_'.format(row[0]) in tituloAnais):
								sigla = row[0]
		#						print(row[8] + ' --> ' + sigla)# + ' #6_#')
								estratos = row[8]
								break
							elif (' {}2'.format(row[0]) in tituloAnais):
								sigla = row[0]
		#						print(row[8] + ' --> ' + sigla)# + ' # 62#')
								estratos = row[8]
								break
																				#Comparação por SIGLA no resultado[5]
							elif (' {} '.format(row[0]) in nomeEvento):
								sigla = row[0]
		#						print(row[8] + ' --> ' + sigla)# + ' # 5 #')
								estratos = row[8]
								break
							elif ('({})'.format(row[0]) in nomeEvento):
								sigla = row[0]
		#						print(row[8] + ' --> ' + sigla)# + ' #(5)')
								estratos = row[8]
								break
							elif ('({} '.format(row[0]) in nomeEvento):
								sigla = row[0]
		#						print(row[8] + ' --> ' + sigla)# + ' #(5 #')
								estratos = row[8]
								break
							elif ('{}&'.format(row[0]) in nomeEvento):
								sigla = row[0]
		#						print(row[8] + ' --> ' + sigla)# + ' #5&#')
								estratos = row[8]
								break
							elif ('{}_'.format(row[0]) in nomeEvento):
								sigla = row[0]
		#						print(row[8] + ' --> ' + sigla)# + ' #5_#')
								estratos = row[8]
								break
							elif (' {}2'.format(row[0]) in nomeEvento):
								sigla = row[0]
		#						print(row[8] + ' --> ' + sigla)# + ' # 52#')
								estratos = row[8]
								break
							elif ('XVII {}'.format(row[0]) in str(nomeEvento).upper()):
								sigla = row[0]
		#						print(row[8] + ' --> ' + sigla)# + ' #XVII 5up#')
								estratos = row[8]
								break
							elif ('({})'.format(row[0]) in resultado[7]):
								sigla = row[0]
		#						print(row[8] + ' --> ' + sigla)# + '#(5)')
								estratos = row[8]
								break
							else:
								sigla = '-'
								estratos = '-'
								
							
						#print(resultado)                #imprime resultado completo
		#				print(tituloAnais)               #imprime apenas o titulo-do-anais-ou-periodico, resultado[6]
						#print(nomeEvento)                #imprime apenas o nome-do-evento, resultado[5]
						#print(estratos)
						
						for row_num in range(worksheet2.nrows):                       #Comparação por nome
							if row_num == 0:
								continue
							row = worksheet2.row_values(row_num)
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
							
						for row_num in range(worksheet2.nrows):                #Comparação por SIGLA casos especiais
							if row_num == 0:
								continue
							row = worksheet2.row_values(row_num)
							if (estratos == '-'):
								if (" ({}'2019)".format(row[0]) in resultado[6]):
									sigla = row[0]
									estratos = row[8]
									break
								elif ("{}'18 ".format(row[0]) in resultado[6] and row[0] != 'ER'):
									sigla = row[0]
									estratos = row[8]
									break
					
					worksheet.write(x, 0, resultado[0])
					worksheet.write(x, 1, resultado[1])
					worksheet.write(x, 4, sigla)
					if ('COMPLETO' in tituloAnais):                          #Correção de tabela, elimina o "COMPLETO" do lugar errado
						worksheet.write(x, 2, resultado[2] + resultado [3] + resultado[4])
						worksheet.write(x, 3, resultado[5])
						worksheet.write(x, 5, resultado[8] + ' / ' + autor)
						worksheet.write(x, 6, resultado[9])
					elif ('COMPLETO' in nomeEvento):
						worksheet.write(x, 2, resultado[2] + resultado[3])
						worksheet.write(x, 3, resultado[4])
						worksheet.write(x, 5, autor + ' / ' + resultado[6])
						worksheet.write(x, 6, resultado[8])
					else:
						worksheet.write(x, 2, resultado[2])
						if (resultado[3] != ''):
							worksheet.write(x, 3, resultado[3])
						else:
							worksheet.write(x, 3, '-')
						worksheet.write(x, 5, tituloAnais + ' / ' + nomeEvento)
						if (len(resultado) > 8):
							if (nomeProf in str(autor).upper()):
								worksheet.write(x, 6, autor)
							elif (nomeProf in str(resultado[8]).upper()):
								worksheet.write(x, 6, resultado[8])
						else:
							worksheet.write(x, 6, autor)
					worksheet.write(x, 7, estratos)
					
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
					worksheet.write(x, 8, nota)
					
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
								#print(resultado_total[i+1])
									#estratos2 = '-'
									continue
								if (len(resultado2[5]) == len(resultado_total[i])):
		#							print(resultado_total[i+1])
									estratos2 = resultado_total[i+1]
									break
								elif (len(resultado2[5]) < len(resultado_total[i])):
									if ('{} (PRINT)'.format(str(resultado2[5]).upper()) == resultado_total[i]):
		#								print(resultado_total[i+1])
										estratos2 = resultado_total[i+1]
		#								print(estratos2)
										break
									if ('{} (ONLINE)'.format(str(resultado2[5]).upper()) == resultado_total[i]):
		#								print(resultado_total[i+1])
										estratos2 = resultado_total[i+1]
		#								print(estratos2)
										break
									elif ('ACS {}'.format(str(resultado2[5]).upper()) == resultado_total[i]):
		#								print(resultado_total[i+1])
										estratos2 = resultado_total[i+1]
		#								print(estratos2)
										break
									elif ('THE {}'.format(str(resultado2[5]).upper()) == resultado_total[i]):
		#								print(resultado_total[i+1])
										estratos2 = resultado_total[i+1]
		#								print(estratos2)
										break
									elif ('{} (19'.format(str(resultado2[5]).upper()) in resultado_total[i]):
		#								print(resultado_total[i+1])
										estratos2 = resultado_total[i+1]
		#								print(estratos2)
										break
									elif (len(resultado_total[i]) - len(resultado2[5]) <= 12):
										if ('{} ('.format(str(resultado2[5]).upper()) in resultado_total[i]):
		#									print(resultado_total[i+1])
											estratos2 = resultado_total[i+1]
		#									print(estratos2)
											#print(resultado_total[i])
											break
										else:
											same = input(resultado2[5] + ' é o mesmo que ' + resultado_total[i] + '? (S/N) \n')
											resp = False
											while (resp == False):
												if (same == 's' or same == 'S'):
		#											print(resultado_total[i+1])
													estratos2 = resultado_total[i+1]
													resp = True
													break
												elif (same == 'n' or same == 'N'):
													estratos2 = '-'
													resp = True
												else:
													same = input('Letra inválida. Digite "S" para sim ou "N" para não \n.')
									elif (len(resultado_total[i]) - len(resultado2[5]) > 12):
										if ('{} ('.format(str(resultado2[5]).upper()) in resultado_total[i]):
											same = input(resultado2[5] + ' é o mesmo que ' + resultado_total[i] + '? (S/N) \n')
											resp = False
											while (resp == False):
												if (same == 's' or same == 'S'):
		#											print(resultado_total[i+1])
													estratos2 = resultado_total[i+1]
													resp = True
													break
												elif (same == 'n' or same == 'N'):
														estratos2 = '-'
														resp = True
												else:
													resp = input('Letra inválida. Digite "S" para sim ou "N" para não. \n')
							elif (str(resultado2[6]).upper() in resultado_total[i]):
		#						print(resultado_total[i+1])
								estratos2 = resultado_total[i+1]
								break
							else:
								estratos2 = '-'
					
		#			print(resultado2[5])   #imprime apenas o nome do periódico
					#print(estratos2)      #imprime só o estrato
					#print(resultado2)     #imprime o resultado completo
					
					worksheet.write(x, 0, resultado2[0])
					worksheet.write(x, 1, resultado2[1])
					worksheet.write(x, 4, '-')
					if ('COMPLETO' in resultado2[5]):                        #Correção de tabela, elimina o "COMPLETO" do lugar errado
						worksheet.write(x, 2, resultado2[2] + resultado2[3])
						worksheet.write(x, 3, resultado2[4])
						worksheet.write(x, 5, resultado2[6])
						worksheet.write(x, 6, resultado2[7])
					else:
						worksheet.write(x, 2, resultado2[2])
						if (resultado2[3] != ''):
							worksheet.write(x, 3, resultado2[3])
						else:
							worksheet.write(x, 3, '-')
						worksheet.write(x, 5, resultado2[5])
						worksheet.write(x, 6, resultado2[6])
					worksheet.write(x, 7, estratos2)
					
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
					worksheet.write(x, 8, nota)
					
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
					
			worksheet.write(x, 7, 'Nota Total')
			worksheet.write(x, 8, totalNota)
		#	print('######')
			print('Total de publicações = {}'.format(cont))            #Quantidade de documentos válidos de cada professor
			print('Pontuação total = {}'.format(totalNota))            #Nota do professor
			print('------------------------------------------------------------')
			
			if (respAno == '1'):
				contTotalc = cont17c + cont18c + cont19c + cont20c
				contTotalp = cont17p + cont18p + cont19p + cont20p
			elif (respAno == '2'):
				contTotalc = cont17c
				contTotalp = cont17p
				totalNota = nota17
			elif (respAno == '3'):
				contTotalc = cont18c
				contTotalp = cont18p
				totalNota = nota18
			elif (respAno == '4'):
				contTotalc = cont19c
				contTotalp = cont19p
				totalNota = nota19
			elif (respAno == '5'):
				contTotalc = cont20c
				contTotalp = cont20p
				totalNota = nota20
				
			#Planilha_2
			yi = 0
			if (xi <= len(curriculos)):
				worksheet3.write(xi, yi, nomeProf)
				yi = yi + 2
				if (r17 == True):                         #Contador de estratos por ano e tipo de publicação
					worksheet3.write(xi, yi, cont17c)
					yi = yi + 1
					worksheet3.write(xi, yi, c17A1)
					yi = yi + 1
					worksheet3.write(xi, yi, c17A2)
					yi = yi + 1
					worksheet3.write(xi, yi, c17A3)
					yi = yi + 1
					worksheet3.write(xi, yi, c17A4)
					yi = yi + 1
					worksheet3.write(xi, yi, c17B1)
					yi = yi + 1
					worksheet3.write(xi, yi, c17B2)
					yi = yi + 1
					worksheet3.write(xi, yi, c17B3)
					yi = yi + 1
					worksheet3.write(xi, yi, c17B4)
					yi = yi + 1
					worksheet3.write(xi, yi, c17C)
					yi = yi + 1
					worksheet3.write(xi, yi, cont17p)
					yi = yi + 1
					worksheet3.write(xi, yi, p17A1)
					yi = yi + 1
					worksheet3.write(xi, yi, p17A2)
					yi = yi + 1
					worksheet3.write(xi, yi, p17A3)
					yi = yi + 1
					worksheet3.write(xi, yi, p17A4)
					yi = yi + 1
					worksheet3.write(xi, yi, p17B1)
					yi = yi + 1
					worksheet3.write(xi, yi, p17B2)
					yi = yi + 1
					worksheet3.write(xi, yi, p17B3)
					yi = yi + 1
					worksheet3.write(xi, yi, p17B4)
					yi = yi + 1
					worksheet3.write(xi, yi, p17C)
					yi = yi + 3
				
				if (r18 == True):
					worksheet3.write(xi, yi, cont18c)
					yi = yi + 1
					worksheet3.write(xi, yi, c18A1)
					yi = yi + 1
					worksheet3.write(xi, yi, c18A2)
					yi = yi + 1
					worksheet3.write(xi, yi, c18A3)
					yi = yi + 1
					worksheet3.write(xi, yi, c18A4)
					yi = yi + 1
					worksheet3.write(xi, yi, c18B1)
					yi = yi + 1
					worksheet3.write(xi, yi, c18B2)
					yi = yi + 1
					worksheet3.write(xi, yi, c18B3)
					yi = yi + 1
					worksheet3.write(xi, yi, c18B4)
					yi = yi + 1
					worksheet3.write(xi, yi, c18C)
					yi = yi + 1
					worksheet3.write(xi, yi, cont18p)
					yi = yi + 1
					worksheet3.write(xi, yi, p18A1)
					yi = yi + 1
					worksheet3.write(xi, yi, p18A2)
					yi = yi + 1
					worksheet3.write(xi, yi, p18A3)
					yi = yi + 1
					worksheet3.write(xi, yi, p18A4)
					yi = yi + 1
					worksheet3.write(xi, yi, p18B1)
					yi = yi + 1
					worksheet3.write(xi, yi, p18B2)
					yi = yi + 1
					worksheet3.write(xi, yi, p18B3)
					yi = yi + 1
					worksheet3.write(xi, yi, p18B4)
					yi = yi + 1
					worksheet3.write(xi, yi, p18C)
					yi = yi + 3
				
				if (r19 == True):
					worksheet3.write(xi, yi, cont19c)
					yi = yi + 1
					worksheet3.write(xi, yi, c19A1)
					yi = yi + 1
					worksheet3.write(xi, yi, c19A2)
					yi = yi + 1
					worksheet3.write(xi, yi, c19A3)
					yi = yi + 1
					worksheet3.write(xi, yi, c19A4)
					yi = yi + 1
					worksheet3.write(xi, yi, c19B1)
					yi = yi + 1
					worksheet3.write(xi, yi, c19B2)
					yi = yi + 1
					worksheet3.write(xi, yi, c19B3)
					yi = yi + 1
					worksheet3.write(xi, yi, c19B4)
					yi = yi + 1
					worksheet3.write(xi, yi, c19C)
					yi = yi + 1
					worksheet3.write(xi, yi, cont19p)
					yi = yi + 1
					worksheet3.write(xi, yi, p19A1)
					yi = yi + 1
					worksheet3.write(xi, yi, p19A2)
					yi = yi + 1
					worksheet3.write(xi, yi, p19A3)
					yi = yi + 1
					worksheet3.write(xi, yi, p19A4)
					yi = yi + 1
					worksheet3.write(xi, yi, p19B1)
					yi = yi + 1
					worksheet3.write(xi, yi, p19B2)
					yi = yi + 1
					worksheet3.write(xi, yi, p19B3)
					yi = yi + 1
					worksheet3.write(xi, yi, p19B4)
					yi = yi + 1
					worksheet3.write(xi, yi, p19C)
					yi = yi + 3
				'''
				if (cont20c > 0 or cont20p > 0): # tem erro aqui 
					if (r20 == True):
						worksheet3.write(xi, yi, cont20c)
						yi = yi + 1
						worksheet3.write(xi, yi, c20A1)
						yi = yi + 1
						worksheet3.write(xi, yi, c20A2)
						yi = yi + 1
						worksheet3.write(xi, yi, c20A3)
						yi = yi + 1
						worksheet3.write(xi, yi, c20A4)
						yi = yi + 1
						worksheet3.write(xi, yi, c20B1)
						yi = yi + 1
						worksheet3.write(xi, yi, c20B2)
						yi = yi + 1
						worksheet3.write(xi, yi, c20B3)
						yi = yi + 1
						worksheet3.write(xi, yi, c20B4)
						yi = yi + 1
						worksheet3.write(xi, yi, c20C)
						yi = yi + 1
						worksheet3.write(xi, yi, cont20p)
						yi = yi + 1
						worksheet3.write(xi, yi, p20A1)
						yi = yi + 1
						worksheet3.write(xi, yi, p20A2)
						yi = yi + 1
						worksheet3.write(xi, yi, p20A3)
						yi = yi + 1
						worksheet3.write(xi, yi, p20A4)
						yi = yi + 1
						worksheet3.write(xi, yi, p20B1)
						yi = yi + 1
						worksheet3.write(xi, yi, p20B2)
						yi = yi + 1
						worksheet3.write(xi, yi, p20B3)
						yi = yi + 1
						worksheet3.write(xi, yi, p20B4)
						yi = yi + 1
						worksheet3.write(xi, yi, p20C)
						yi = yi + 3
						
						worksheet3.write(0, col, u'2020')
						col = col + 1
						worksheet3.write(0, col, u'Conferência')
						col = col + 1
						worksheet3.write(0, col, u'A1')
						col = col + 1
						worksheet3.write(0, col, u'A2')
						col = col + 1
						worksheet3.write(0, col, u'A3')
						col = col + 1
						worksheet3.write(0, col, u'A4')
						col = col + 1
						worksheet3.write(0, col, u'B1')
						col = col + 1
						worksheet3.write(0, col, u'B2')
						col = col + 1
						worksheet3.write(0, col, u'B3')
						col = col + 1
						worksheet3.write(0, col, u'B4')
						col = col + 1
						worksheet3.write(0, col, u'C')
						col = col + 1
						worksheet3.write(0, col, u'Periódico')
						col = col + 1
						worksheet3.write(0, col, u'A1')
						col = col + 1
						worksheet3.write(0, col, u'A2')
						col = col + 1
						worksheet3.write(0, col, u'A3')
						col = col + 1
						worksheet3.write(0, col, u'A4')
						col = col + 1
						worksheet3.write(0, col, u'B1')
						col = col + 1
						worksheet3.write(0, col, u'B2')
						col = col + 1
						worksheet3.write(0, col, u'B3')
						col = col + 1
						worksheet3.write(0, col, u'B4')
						col = col + 1
						worksheet3.write(0, col, u'C')
						col = col + 2
				'''	
				yi = yi - 1
				worksheet3.write(xi, yi, contTotalc)
				yi = yi + 1
				worksheet3.write(xi, yi, contTotalp)
				yi = yi + 1
				worksheet3.write(xi, yi, totalNota)
					
				somaNotas = somaNotas + totalNota
				xi = xi + 1
				
				
		worksheet3.write(0, col, u'Total Conferências')
		col = col + 1
		worksheet3.write(0, col, u'Total Periódicos')
		col = col + 1
		worksheet3.write(0, col, u'Pontuação Total')
		
		mediaNotas = (somaNotas/len(curriculos))
		worksheet3.write(xi+1, yi-1, 'SOMA')
		#worksheet3.write(xi+1, 70, somaNotas)
		#worksheet3.write(xi+1, 71, '=SOMA(BS2:BS23)')#As fórmulas ficam apenas como texto, precisa clicar na célula e apertar "enter"
		worksheet3.write(xi+2, yi-1, 'MÉDIA')
		#worksheet3.write(xi+2, 70, mediaNotas)
		#worksheet3.write(xi+2, 71, '=SOMA(BS2:BS23)/len(curriculos)')
		
		workbook.save('Resultado.xls')#salva em arquivo xls
		
		###############################################################
		#
		#############################INTEGRAÇÃO DO PROGRAMA DE CORREÇÃO
		#
		###############################################################
		resp = False
		decisao = input('''DESEJA APLICAR A CORREÇÃO DE NOTAS? (S/N) \nA correção é a divisão de notas para uma publicação que está no currículo de mais de um professor. \n''')
		while (resp == False):
			if (decisao == 's' or decisao == 'S'):
				print('\n')
				print('CORRIGINDO NOTAS, POR FAVOR AGUARDE. \n')
				
				rb = xlrd.open_workbook('Resultado.xls')        #Ler arquivo para fazer cópia
				wb = copy(rb)
				
				lista = []
				workbook = xlrd.open_workbook('Resultado.xls')  #Carrega arquivo para leitura
				worksheet = workbook.sheet_by_index(0)
				for row_num in range(worksheet.nrows):
					if row_num == 0:
						continue
					row = worksheet.row_values(row_num)
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
				for row_num in range(worksheet.nrows):     #Varre linha por linha do NotasExtraídas
					w_sheet = wb.get_sheet(0)
					if row_num == 0:
						continue
					row = worksheet.row_values(row_num)
					if (row[0] != '' and row[1] == ''):
						print('Corrigindo notas de {}'.format(row[0]))
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
						if (respAno == '2'):
							totalNotas2 = nota172
							print('Pontuação total = {}'.format(totalNotas2))
							print('------------------------------------------------------------')
						elif (respAno == '3'):
							totalNotas2 = nota182
							print('Pontuação total = {}'.format(totalNotas2))
							print('------------------------------------------------------------')
						elif (respAno == '4'):
							totalNotas2 = nota192
							print('Pontuação total = {}'.format(totalNotas2))
							print('------------------------------------------------------------')
						elif (respAno == '5'):
							totalNotas2 = nota202
							print('Pontuação total = {}'.format(totalNotas2))
							print('------------------------------------------------------------')
						else:
							print('Pontuação total = {}'.format(totalNotas2))
							print('------------------------------------------------------------')
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
						
				wb.save('Resultado.xls')
				print('\n')
				print('NOTAS CORRIGIDAS! \nPara conferir a planilha com os resultados, consulte o arquivo Resultados.xls.')
				resp = True
			elif (decisao == 'n' or decisao == 'N'):
				print('PROGRAMA ENCERRADO.')
				resp = True
			else:
				decisao = input('Letra inválida. Digite "S" para sim ou "N" para não. \n')
		desejo2 = True
	elif (desejo == 'n' or desejo == 'N'):
		print('PROGRAMA ENCERRADO.')
		desejo2 = True
	else:
		desejo = input('Letra inválida. Digite "S" para sim ou "N" para não. \n')