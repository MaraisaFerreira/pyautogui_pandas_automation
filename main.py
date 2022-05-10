import pyautogui
import pyperclip
import pandas as pd
import time

pyautogui.PAUSE = 1

pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')


time.sleep(8)
pyautogui.hotkey('ctrl', 't')
pyperclip.copy(
    'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

time.sleep(8)
pyautogui.click(x=374, y=319, clicks=2)

time.sleep(8)
pyautogui.click(x=430, y=428, button='right')

time.sleep(8)
pyautogui.click(x=602, y=641)

time.sleep(8)
pyautogui.click(x=511, y=448)

time.sleep(8)
table = pd.read_excel(r'C:\Users\MF\Downloads\Vendas - Dez.xlsx')

amount = table['Quantidade'].sum()
total = table['Valor Final'].sum()

pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

pyautogui.sleep(8)
pyautogui.click(x=118, y=209)

pyautogui.sleep(8)
pyperclip.copy('EMAIL de destino aqui')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
pyautogui.press('tab')

pyperclip.copy('Automação')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

body = f'''
Email com dados da tabela excell:
quantidade: {amount:,}
total: {total:,.2f}

'''
pyperclip.copy(body)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'enter')
