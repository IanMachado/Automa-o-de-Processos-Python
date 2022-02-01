import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 2

# é extremamente importante colocar tempo no pyautogui, pro computador processar!!!

pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')
time.sleep(3)
pyautogui.click(x=212, y=48)

pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(3)
pyautogui.click(x=316, y=233, clicks=2)
pyautogui.click(x=398, y=309)
pyautogui.click(x=1201, y=148)
pyautogui.click(x=1045, y=490)
time.sleep(4)

import pandas as pd
import numpy
import openpyxl

# 'as pd' apelido de pandas, pra ficar melhor!!
# No pandas não precisa se preocupar com tempo!!!!

tabela = pd.read_excel(r'C:\Users\ianqm\Downloads\Vendas - Dez.xlsx')
# Se quisesse pegar outra aba do excel era só colocar ('...Dez.xlsx', sheets = x) x = a aba que vc quer pegar

# o 'r' é pro python não ler caracteres especiais(já que '\' é especial)
print(tabela)
faturamento = tabela['Valor Final'].sum()
qtde_produtos = tabela['Quantidade'].sum()


pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(2)
pyautogui.click(x=34, y=176)
pyautogui.write('brunomachado3818@gmail.com')
pyautogui.press('enter')
pyautogui.press('tab')

pyperclip.copy('TESTE DE AUTOMAÇÃO')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
pyautogui.write(f'Prezados analistas, bom dia.\n'
                f'O faturamento de ontem foi de: R${faturamento:,.2f}\n'
                f'A quantidade de produtos vendidos foi de: {qtde_produtos:,}\n'
                f' \n'
                f'Abs\n'
                f'Ian Machado ')

pyautogui.click(x=838, y=636)










