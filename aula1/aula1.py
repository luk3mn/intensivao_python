import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1 # permite uma pausa de 1seg em cada linha do pyoutogui

# Passo 1: Entrar no sistema da empresa (link do drive)
time.sleep(3)
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")   
pyautogui.press("enter")

# Passo 2: Navegar até o local do relatório (entrar na pasta exportar)
time.sleep(10)
pyautogui.click(x=347, y=342, clicks=2)
time.sleep(3)

# Passo 3: Exportar o relatório (fazer o download)
pyautogui.click(x=393, y=412)
pyautogui.click(x=1709, y=245)
pyautogui.click(x=1460, y=609)

# Passo 4: Calcular os indicadores (faturamento e quantidade de produtos)
import pandas as pd
time.sleep(3)
df = pd.read_excel('/home/lukemn/Downloads/Vendas - Dez.xlsx')
quantidade = df['Quantidade'].sum()
faturamento = df['Valor Final'].sum()

# Passo 5: Enviar um e-mail para a diretoria
time.sleep(3)
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/1/#inbox")
pyautogui.hotkey("ctrl","v")
pyautogui.press('enter')

time.sleep(5)
pyautogui.click(x=83, y=243)
time.sleep(5)
pyautogui.write('renann.luke@gmail.com')
pyautogui.hotkey("tab")

pyautogui.hotkey("tab")
pyperclip.copy('Relatório de Vendas')
pyautogui.hotkey("ctrl","v")
pyautogui.hotkey("tab")

texto = f"""Segue relatório de vendas. 
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Qualquer dúvida estou à disposição.

Att.,Lukemn"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")
pyautogui.hotkey("ctrl","enter")

# time.sleep(3)
# print(pyautogui.position())