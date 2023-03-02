#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# A intenção deste código é de enviar um e-mail assim que inicia a jornada de trabalho, sendo demonstrado na mensagem do email o faturamento e a quantidade de produtos vendidos no dia anterior. Tudo isso de forma automatizada, bastando apenas rodar o código uma única vez. Após isso, o próprio código fará o download do arquivo xlsx localizado no Google Drive, e através de comandos básicos da biblioteca Pandas, será automatizado a análise e o envio dos dados de faturamento e quantidade de produtos vendidos no dia anterior para a diretoria.
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com.br
# 
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado e a biblioteca Pandas para análise de dados.

# In[35]:


#Importando bibliotecas

import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

pyautogui.hotkey("ctrl", "t") #abre nova guia
pyperclip.copy("https://drive.google.com/drive") #Copia o texto indicado entre parenteses
pyautogui.hotkey("ctrl", "v") #Cola o texto indicado entre parenteses
pyautogui.press("enter")
time.sleep(1.5)
pyautogui.click(x=468, y=277, clicks=2) #Duplo clique com o mouse no local indicado
time.sleep(2.5)
pyautogui.click(x=468, y=277, button="right") #Seleciona para acessar a pasta correspondente no Google Drive
time.sleep(1.5)
pyautogui.click(x=682, y=756, button="left") #Seleciona o arquivo para download
time.sleep(5) #5 segundos 


# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[56]:


#Análise básica de planilha com Pandas

import pandas as pd

tabela = pd.read_excel("Vendas - Dez.xlsx")
display(tabela)
faturamento = tabela["Valor Final"].sum() #Soma da coluna "Valor Final"
quantidade = tabela["Quantidade"].sum( ) #Soma da coluna "Quantidade"


# ### Vamos agora enviar um e-mail pelo gmail

# In[57]:


#Envio do email utilizando automação pela biblioteca pyautogui

pyautogui.hotkey("ctrl", "t") #abre nova guia
pyperclip.copy("https://mail.google.com/mail/u/1/#inbox") #copia o texto indicado entre os parenteses
pyautogui.hotkey("ctrl", "v") #cola o texto indicado entre os parenteses
pyautogui.press("enter") # "pressiona" o botão enter do teclado
time.sleep(1.5)
pyautogui.click(x=191, y=192, clicks=2) #duplo click no local indicado
time.sleep(2.5)
pyautogui.write("seugmail+diretoria@gmail.com.br") #escreve no campo de email do Gmail o email desejo
time.sleep(1.5)
pyautogui.press("tab") #botão para pular para o título do email
time.sleep(1)
pyautogui.press("tab") #botão para pular para o título do email
pyperclip.copy("Relatório de Vendas") #copia o texto indicado entre os parenteses
pyautogui.hotkey("ctrl", "v") #cola o texto indicado pelos parenteses
time.sleep(1)
pyautogui.press("tab") #botão para pular para o corpo do email

texto = f"""
Prezados,

Segue relatório de vendas.
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Qualquer dúvida estou à disposição.
Att.,
Lucas
"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter") #envio do email para a diretoria


# #### Código para descobrir qual a posição de um item que queira clicar

# In[55]:


time.sleep(5)
pyautogui.position()

