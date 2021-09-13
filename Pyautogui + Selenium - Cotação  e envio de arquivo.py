#!/usr/bin/env python
# coding: utf-8

# ### Roteiro:  
# - <font color='red'>abrir navegador, pegar cotação dólar e euro      -SL  -</font> 
# - <font color='orange'>abrir o navegador, entrar no Discord Web, clicar no Server, enviar a mensagem de cotação e faturamento          -SL</font>   
# - <font color='green'>baixar os arquivos de faturamento pelo Google Drive   -PA</font> 
# - <font color='purple'>recortar o arquivo de downloads para a pasta do projeto   -PA-</font> 
# - <font color='pink'>Anexar o arquivo e enviar separado no Server do Discord       - PA       </font> 
# 
#  
# 

#  ### Texto no Discord  
# Olá, amigs!  
# 
# Pra se manterem informados:  
# 
# A cotação do dólar hoje está em: {cotação_dólar}  
# A cotação do euro hoje está em: {cotação_euro}  
# 
# Faturamento da empresa:  
# 
# Ontem  nossa empresa obteve o faturamento de: {faturamento} e a nossa meta de {meta} foi atingida com êxito! (ou não foi atingida com êxito.)  
# 
# Segue em anexo a planilha que é atualizada diariamente e enviada para vocês automaticamente.  
# 
# Atenciosamente,  
# 
# Evilly >emote  
# 

# In[23]:


#1°  baixar os arquivos de faturamento pelo Google Drive -PA
#2°  recortar o arquivo de downloads para a pasta do projeto -PA

import pyautogui as py 
import pandas as pd 
import pyperclip

#ENTRANDO NO LINK
py.PAUSE = 2 #tempo de espera para executar cada ação
py.press('win') #pressiona tecla windows
py.write('chrome') #digita chrome
py.press('enter') #aperta enter para abrir o navegador
py.hotkey('win','up') #maximiza a tela do google
link = (r'https://drive.google.com/drive/u/0/folders/10m5QYxXsmk0S7zWRxS5Gx-RsHf12tjSA') #link do arquivo
pyperclip.copy(link) #copia link do arquivo
py.hotkey('ctrl' , 'v') #cola link do arquivo na barra de pesquisa
py.press('enter') #aperta enter para pesquisar
py.sleep(5) #tempo de espera para baixar o arquivo

#BAIXANDO O ARQUIVO
py.rightClick(x=384, y=378) #clica com o botão direito no arquivo
py.click(x=436, y=773) #clica em fazer download
py.sleep(5) #tempo de espera para baixar o arquivo

#MOSTRANDO O ARQUIVO NA PASTA
py.click(x=208, y=829) #clica na setinha do arquivo baixado
py.click(x=246, y=765) #clica em mostrar na pasta

#MOVENDO O ARQUIVO BAIXADO
py.hotkey('win','up') #maximiza a tela do explorador de arquivos  
py.click(x=79, y=567) #clica em Windows C
py.doubleClick(x=234, y=277) #clica em Usuários
py.doubleClick(x=218, y=128) #clica em Evilly
py.doubleClick(x=240, y=359) #duplo clique para selecionar a pasta Jupyter
py.click(x=1318, y=68) #clica na barra de pesquisa
py.press('Caps Lock')
pasta = ('PROJETOS') #digita o nome no campo de pesquisa
pyperclip.copy(pasta) #copia nome da pasta
py.hotkey('ctrl' , 'v') #cola nome da pasta
py.press('enter') #aperta enter
py.doubleClick(x=323, y=117) #abre a pasta de projetos
py.click(x=1318, y=68) #clica no campo de pesquisa
pasta2 = ('Projeto Pyautogui + Selenium')
pyperclip.copy(pasta2)
py.hotkey('ctrl','v')
py.press('enter')
py.doubleClick(x=373, y=121)



#ABRINDO OUTRO EXPLORADOR E COLANDO O ARQUIVO
py.rightClick(x=171, y=882) #clica com o botão direito em cima da tarefa aberta na barra inferior
py.click(x=141, y=780) #abre um novo explorador
py.hotkey('win' , 'up') #maximiza a janela
py.click(x=84, y=165) #clicar em downloads
py.click(x=228, y=159) #clica no arquivo
py.hotkey('ctrl' , 'x')#recorta o arquivo
py.click(x=177, y=885) #clica no explorador na barra de tarefas
py.click(x=135, y=793) #seleciona o outro
py.hotkey('ctrl' , 'v') #cola o arquivo na pasta
py.press('enter')


# In[24]:


#3° abrir o navegador, entrar no Discord Web, clicar no Server e anexar o arquivo-PY e SL
import pyautogui as py
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

#PS: O Selenium aguarda a página carregar automaticamente, não precisa definir.
navegador = webdriver.Chrome() #abre o navegador
navegador.get('https://discord.com/login') #insere o link

#clicar = navegador.find_element_by_xpath(
#    '//*[@id="app-mount"]/div/div/div[1]/div[2]/div/div[2]/button').click() #clica em um botão na página

navegador.find_element_by_xpath(
    '//*[@id="app-mount"]/div[2]/div/div/div/div/form/div/div/div[1]/div[2]/div[1]/div/div[2]/input').send_keys('evillyf05@gmail.com') #escreve o meu e-mail na caixa de login
navegador.find_element_by_xpath('//*[@id="app-mount"]/div[2]/div/div/div/div/form/div/div/div[1]/div[2]/div[1]/div/div[2]/input').send_keys(Keys.TAB) #aperta tab pra mudar de campo
navegador.find_element_by_xpath('//*[@id="app-mount"]/div[2]/div/div/div/div/form/div/div/div[1]/div[2]/div[2]/div/input').send_keys('sandydias') #digita a senha no campo
navegador.find_element_by_xpath('//*[@id="app-mount"]/div[2]/div/div/div/div/form/div/div/div[1]/div[2]/div[2]/div/input').send_keys(Keys.ENTER) #pressiona enter para logar

#enviar a mensagem e anexar o arquivo com o Py
py.PAUSE = 2 #tempo de espera para cada ação
py.click(x=551, y=35) #clica em uma área neutra da página
py.hotkey('win','up') #maximiza a janela
py.click(x=38, y=220) #clica no meu Server
py.click(x=388, y=814) #clica em na área de chat para sair o pop-out
py.click(x=388, y=814) #clica no campo da conversa
py.hotkey('ctrl','shift','u')
py.click(x=126, y=61) #clica na barra de pesquisa de diretório
diretorio = (r'C:\Users\Evilly\Jupyter\PROJETOS\Projeto Pyautogui + Selenium')
pyperclip.copy(diretorio)
py.hotkey('ctrl' , 'v')
py.press('enter')
py.click(x=194, y=422) #digitar no campo de inserir o nome do arquivo
py.write('VendasAmazon.xlsx') #escrever o nome do arquivo
py.press('enter') #arquivo em anexo
py.press('enter') #para enviar o arquivo ###conferir


# In[1]:


#4° atualizar os dados, abrir navegador, pegar cotação dólar e euro -SL e PY


# ATUALIZAR A BASE DE DADOS

import pandas as pd 
import pyautogui as py 
import pyperclip

py.PAUSE = 2 
tabela = pd.read_excel('VendasAmazon.xlsx')
#tabela = tabela.drop('SalesPeriod', axis=1)
tabela = tabela.drop(columns=['Unnamed: 5','Marketplace']) #Remove mais de uma coluna
tabela["Price"] = pd.to_numeric(tabela["Price"], errors="coerce")
#tabela['Amount'] = pd.to_numeric(tabela['Amount'], errors='coerce') #converter tabela pra int
preço = tabela['Price'].sum() #soma a coluna
qntd = tabela['Amount'].sum() #soma esta coluna
faturamento = (preço * qntd)
meta = 70000


#if faturamento >= meta:

#    print('A nossa meta é de R$ {:,.0f}\nNós alcançamos R$ {:,.0f}\nEsse mês nós atingimos a meta!'.format(meta,faturamento)) #.replace(',','.')  
#else:  
#    print('oi')
#print(tabela.info()) #Para ver as especificações da tabela, se em tal lugar é int, str e ect 
print(tabela)
print(preço)
print(qntd)




#PEGANDO COTAÇÃO
from selenium import webdriver
from selenium.webdriver.common.keys import Keys  #Passar essa linha para comentário qndo quiser rodar em segundo plano, tentei mas não foi em 2° plano

#RODAR A APLICAÇÃO EM SEGUNDO PLANO
#from selenium.webdriver.chrome.options import Options
#chrome_options = Options()
#chrome_options.headless = True 
#navegador = webdriver.Chrome(options=chrome_options)



#abrir o navegador
navegador = webdriver.Chrome()

#ENTRANDO NO NAVEGADOR E BUSCANDO PELAS INFORMAÇÕES
#DEVE-SE ENTRAR NO MODO ANÔNIMO PARA PODER PEGAR AS LOCALIZAÇÕES
navegador.get('https://www.google.com/') #Digita direto na barra de pesquisa pq já tá selecionado automaticamente


#como pegar loc do elemento? Assim: botão direito no elemento ou na página > inspecionar > setinha superior esquerda
#p/ selecionar o seu elemento > seleciona > em azul clica com botão direito > copiar > copiar xpath
navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação dólar') #Localização da barra de pesquisa e texto que será buscado
navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)


#pega valor da cotação dólar  ### IMPORTANTE: No atributo deve-se usar ASPAS DUPLAS ###
cotacao_dolar = navegador.find_element_by_xpath(
    '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value") #ASPAS DUPLAS. pega o atributo do valor
print(cotacao_dolar)



#pega valor da cotação euro
navegador.get('https://www.google.com/')

navegador.find_element_by_xpath(
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação euro')
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_euro = navegador.find_element_by_xpath(
    '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value") #ASPAS DUPLAS

#FORMANTANDO COTAÇÕES
cotacao_dolar = ('{:.4}'.format(cotacao_dolar).replace('.',','))
cotacao_euro = ('{:.4}'.format(cotacao_euro).replace('.',','))
print(cotacao_dolar,cotacao_euro)

# ENVIAR A MENSAGEM DE COTAÇÃO E FATURAMENTO

#PS: O Selenium aguarda a página carregar automaticamente, não precisa definir.
navegador = webdriver.Chrome() #abre o navegador
navegador.get('https://discord.com/login') #insere o link


#FAZENDO LOGIN
navegador.find_element_by_xpath(
    '//*[@id="app-mount"]/div[2]/div/div/div/div/form/div/div/div[1]/div[2]/div[1]/div/div[2]/input').send_keys('evillyf05@gmail.com') #escreve o meu e-mail na caixa de login
navegador.find_element_by_xpath('//*[@id="app-mount"]/div[2]/div/div/div/div/form/div/div/div[1]/div[2]/div[1]/div/div[2]/input').send_keys(Keys.TAB) #aperta tab pra mudar de campo
navegador.find_element_by_xpath('//*[@id="app-mount"]/div[2]/div/div/div/div/form/div/div/div[1]/div[2]/div[2]/div/input').send_keys('sandydias') #digita a senha no campo
navegador.find_element_by_xpath('//*[@id="app-mount"]/div[2]/div/div/div/div/form/div/div/div[1]/div[2]/div[2]/div/input').send_keys(Keys.ENTER) #pressiona enter para logar

#ENVIAR A MENSAGEM 
py.PAUSE = 2 #tempo de espera para cada ação
py.click(x=551, y=35) #clica em uma área neutra da página
py.hotkey('win','up') #maximiza a janela
py.click(x=38, y=220) #clica no meu Server
py.click(x=388, y=814) #clica em na área de chat para sair o pop-out
py.click(x=388, y=814) #clica no campo da conversa

if faturamento >= meta:
    mensagem = f'''Olá, amigs!  
Para se manterem informados:  
A cotação do dólar hoje está em: {cotacao_dolar}  
A cotação do euro hoje está em: {cotacao_euro}  
(Atualiza diretamente do google)
Faturamento da empresa:  
Ontem  nossa empresa obteve o faturamento de: {faturamento} e a nossa meta de {meta} foi atingida com êxito!   
Segue em anexo a planilha que é atualizada diariamente e enviada para vocês automaticamente.  
(O cálculo é feito automático)
Atenciosamente,  
Evilly'''

    


else:
    faturamento < meta
    mensagem = f'''Olá, amigs!  
Para se manterem informados:  
A cotação do dólar hoje está em: {cotacao_dolar}  
A cotação do euro hoje está em: {cotacao_euro}  
(Atualiza diretamente do google)
Faturamento da empresa:  
Ontem  nossa empresa obteve o faturamento de: {faturamento} e a nossa meta de {meta} foi atingida com êxito!   
Segue em anexo a planilha que é atualizada diariamente e enviada para vocês automaticamente.  
(O cálculo é feito automático)
Atenciosamente,  
Evilly'''
    
pyperclip.copy(mensagem)
py.hotkey('ctrl','v')
py.press('enter')

  
 


# In[ ]:





#    ###POSIÇÃO DO CLIQUE  
# import time  
# import pyautogui  
# time.sleep(5)  
# pyautogui.position()  
