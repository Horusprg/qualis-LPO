PARA INICIAR A MÁQUINA VIRTUAL:
conda activate py39env

Criar package:
pyinstaller GeradorQualis.py --windowed --clean --icon= lpo.ico

<h1 align="center">Consulta Qualis - LPO<BR><img src="https://www.lpo.ufpa.br/logo.png" width="140" height="80"/> <img src="https://iconape.com/wp-content/files/wn/195324/svg/195324.svg" width="120" height="120"/> <img src="https://www.ppgee.propesp.ufpa.br/IMAGENS/ppgee_site.jpg" width="180" height="70"/></>
  
---
  Introdução
  
<h3>O objetivo dessa aplicação é facilitar o processo de consulta da lista sucupira e cálculo das pontuações dos currículos selecionados.</>
<h1 align="center"></>

---
  Estrutura
  
<h3>A aplicação foi desenvolvida utilizando a biblioteca tkinter, a fim de produzir um executável para distribuição. Devido a isso, todo o código está contido em um único arquivo (link), tal que está dividido em:<BR/>
<BR/>
  
1. Definição da lista de referência para as correções
  <img src="" width="140" height="80"/>
2. Função para centralizar o ecrã
3. Aplicação usando o tkinter
4. Tela inicial
5. Tela de currículos importados pelo usuário
6. Análise dos currículos baseado no QUALIS novo e o arquivo de eventos importados
7. Aplicação da correção das notas
8. Comparação do QUALIS antigo e o QUALIS novo das publicações dos currículos analisados
