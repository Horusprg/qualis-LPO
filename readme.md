<h1 align="center">Consulta Qualis - LPO<BR><img src="https://www.lpo.ufpa.br/logo.png" width="140" height="80"/> <img src="https://iconape.com/wp-content/files/wn/195324/svg/195324.svg" width="140" height="120"/> <img src="https://www.ppgee.propesp.ufpa.br/IMAGENS/ppgee_site.jpg" width="180" height="70"/></>

---
  Introdução
  
<h3>O objetivo dessa aplicação é facilitar o processo de consulta da lista sucupira e cálculo das pontuações dos currículos selecionados.
  </br></br>Este documento tem como objetivo esclarecer como está disposta a estrutura de desenvolvimento da aplicação, para acessar o tutorial de como utilizar a ferramenta acesse o arquivo abaixo:</>
</br></br>
<a href="https://github.com/Horusprg/qualis-LPO/blob/main/Tutorial%20Buscador%20Qualis.pdf">Tutorial Buscador Qualis</a>

</br></br>
Para fazer o download da ferramenta, acesse o link:
</br></br>
<img src="https://upload.wikimedia.org/wikipedia/commons/4/48/Windows_logo_-_2012_%28dark_blue%29.svg" width="15" height="15"/> <a href="https://drive.google.com/file/d/1WBQQsPLNO-_1p6LZZaRhtjNR8l_1E-PY/view?usp=sharing"> Gerador Qualis - Windows</a>
<h1 align="center"></>

  Estrutura
  
<h3>A aplicação foi desenvolvida utilizando a biblioteca tkinter, a fim de produzir um executável para distribuição. Devido a isso, todo o código está contido no arquivo <a href="https://github.com/Horusprg/qualis-LPO/blob/main/GeradorQualis.py">GeradorQualis.py</a>, tal que está dividido em:<BR/>
<BR/>
  
1. Definição da lista de referência para as correções
  <img src="https://github.com/Horusprg/qualis-LPO/blob/main/assets/referencias.png"/>
  <BR/><BR/>
  
2. Função para centralizar o ecrã
  <img src="https://github.com/Horusprg/qualis-LPO/blob/main/assets/centralizar.png"/>
  <BR/><BR/>
  
3. Aplicação usando o tkinter
  <img src="https://github.com/Horusprg/qualis-LPO/blob/main/assets/aplic.png"/>
  <BR/><BR/>
  
4. Tela inicial
  <img src="https://github.com/Horusprg/qualis-LPO/blob/main/assets/home.png"/>
  <BR/><BR/>
  
5. Tela de currículos importados pelo usuário
  <img src="https://github.com/Horusprg/qualis-LPO/blob/main/assets/curriculos.png"/>
  <BR/><BR/>
  
6. Análise dos currículos baseado no QUALIS novo e o arquivo de eventos importados
  <img src="https://github.com/Horusprg/qualis-LPO/blob/main/assets/analise_curriculos.png"/>
  <BR/><BR/>
  
7. Aplicação da correção das notas
  <img src="https://github.com/Horusprg/qualis-LPO/blob/main/assets/correcao.png"/>
  <BR/><BR/>
  
8. Comparação do QUALIS antigo e o QUALIS novo das publicações dos currículos analisados
  <img src="https://github.com/Horusprg/qualis-LPO/blob/main/assets/mapeamento.png"/>
