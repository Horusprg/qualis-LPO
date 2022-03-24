PARA INICIAR A MÁQUINA VIRTUAL:
conda activate py39env

Criar package:
pyinstaller GeradorQualis.py --windowed --clean --icon= lpo.ico

<h1>Tutorial<h1/>
<h2>Introdução<h2/>
<h3>O objetivo dessa aplicação é facilitar o processo de consulta da lista sucupira e cálculo das pontuações dos currículos selecionados.<h3/>
<h2>Estrutura<h2/>
<h3>
A aplicação foi desenvolvida utilizando a biblioteca tkinter, a fim de produzir um executável para distribuição. Devido a isso, todo o código está contido em um único arquivo (link), tal que está dividido em:

1. Definição da base de correções
2. Aplicação usando o tkinter
2.1 Tela inicial
2.2 Tela de currículos importados pelo usuário
2.3 Análise dos currículos baseado no QUALIS novo e o arquivo de eventos importados
2.4 Aplicação da correção das notas
2.5 Comparação do QUALIS antigo e o QUALIS novo das publicações dos currículos analisados
