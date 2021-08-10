import pyautogui
import pyperclip
import time
import pandas as pd
#from IPython.display import display

pyautogui.PAUSE = 1

#Abrir o sistema
pyautogui.press("win")
pyautogui.write("Google Chrome")
pyautogui.press("enter")
link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(3)

#Encontrar o arquivo
#pyautogui.click(x=390,y=375, clicks=2)
pyautogui.doubleClick(x=390,y=375)
time.sleep(2)

#Baixar o arquivo
pyautogui.rightClick(x=390,y=375)
pyautogui.click(x=485,y=855)
time.sleep(5)

#Calcular o faturamento
tabela = pd.read_excel(r"C:\Users\vinic\Downloads\Vendas - Dez.xlsx")
##display(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

#Abrir email
pyautogui.hotkey("ctrl", "t")
link = "https://mail.google.com"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(3)

#Escrever email
pyautogui.click(x=80,y=205)
pyautogui.write("vini_sk892@hotmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
assuntoEmail = "Relatório de Vendas"
pyperclip.copy(assuntoEmail)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
corpoRelatorio = f"""Prezados, 

O faturamento foi de R${faturamento:,.2f}
A quantidade de produtos vendidos foi de {quantidade:,}

Atenciosamente, 

VC"""
pyautogui.write(corpoRelatorio)
pyautogui.hotkey("ctrl", "enter")
