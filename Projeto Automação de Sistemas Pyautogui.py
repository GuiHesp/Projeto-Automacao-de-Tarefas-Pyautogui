#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Para controle de custos, todos os dias, seu chefe pede um relatório com todas as compras de mercadorias da empresa.
# O seu trabalho, como analista, é enviar um e-mail para ele, assim que começar a trabalhar, com o total gasto, a quantidade de produtos compradas e o preço médio dos produtos.
# 
# E-mail do seu chefe: para o nosso exercício, coloque um e-mail seu como sendo o e-mail do seu chefe<br>
# Link de acesso ao sistema da empresa: https://pages.hashtagtreinamentos.com/aula1-intensivao-sistema
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# In[2]:


import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

# escrever o passo a passo do q fazer para resolver o desafio:
# Passo 1: entrar no sistema
pyautogui.hotkey("ctrl", "t")

# link imcompleto/url errada/caractere especial
#pyperclip.copy("link zuado")
#pyautogui.hotkey("ctrl", "v")

pyautogui.write("https://pages.hashtagtreinamentos.com/aula1-intensivao-sistema")
pyautogui.press("enter")

time.sleep(3)

# Passo 2: navegar até o local do relatório

pyautogui.click(x=1341, y=384)
pyautogui.write("login")
pyautogui.click(x=1335, y=464)
pyautogui.write("senha")
pyautogui.click(x=1453, y=530)  #clicar no login 

time.sleep(2)                                # parametro no click para duplo click (clicks=2)

# Passo 3: fazer download do relatório
pyautogui.click(x=1528, y=316)
pyautogui.click(x=1612, y=809)

time.sleep(5)

# Passo 4: Calcular os indicadores
tabela = pd.read_csv(r"C:\Users\Gui\Downloads\Compras.csv", sep=';')
display(tabela)

faturamento = tabela["ValorFinal"].sum()
qtdade = tabela["Quantidade"].sum()

display(faturamento)
display(qtdade)

# Passo 5: Entrar e-mail
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# Passo 6: Enviar por e-mail o resultado

pyautogui.click(x=1059, y=200)
pyautogui.write("guihesp+diretoria@gmail.com")
pyautogui.press("tab") #seleciona o email
pyautogui.press("tab") #pula para o campo assunto

pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab") #pular corpo do email
texto = f"""Prezados,

Segue o Relatório de Vendas.

O faturamento foi de R${faturamento:,.2f} reais.
A quantidade de produtos foi de {qtdade:,} unidades.

Abs, Gui."""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# clicar no botao enviar
pyautogui.hotkey("ctrl", "enter")


# In[5]:


#time.sleep(3)
#pyautogui.position()

