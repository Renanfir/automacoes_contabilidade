import re

import pyautogui as bot
import os
from os import listdir
import pyperclip
emailaureo = ''
emailvalmor = ''


pasta = r'C:\Users\Usuario\Pictures\Screenshots'

bot.click(1800, 21)
bot.sleep(1)
bot.click(1753, 306)
bot.press('end')
bot.press('end')
bot.press('end')
bot.sleep(0.4)
bot.click(765, 562)
bot.write(emailaureo)
bot.press('tab')
bot.write('')
bot.press('tab')
bot.write('ADVOGADO')
bot.press('tab')
bot.write('')
bot.press('tab')
bot.write(emailvalmor)
bot.press('tab')
bot.press('tab')
bot.press('enter')
bot.sleep(1.5)
bot.press('enter')
bot.sleep(1.5)
bot.click(320, 662)
bot.sleep(1.5)
bot.click(1701, 285)
bot.sleep(1)
bot.click(138, 349)
bot.write('')
bot.press('tab')
bot.press('tab')
bot.sleep(0.2)
bot.write(emailvalmor)
bot.press('tab')
bot.write('ADVOGADO')
bot.press('tab')
bot.press('up')
bot.press('tab')
bot.write('42853')
bot.press('tab')
bot.write('s')
bot.press('tab')
bot.press('tab')
bot.press('tab')
bot.press('space')
bot.press('tab')
bot.sleep(0.2)
bot.press('enter')
bot.sleep(2)
bot.press('enter')
bot.sleep(2)
bot.click(1703, 292)
bot.sleep(1.2)
bot.click(150, 351)
bot.write("")
bot.press('tab')
bot.press('tab')
bot.sleep(0.2)
bot.write(emailaureo)
bot.press('tab')
bot.write('CONTADOR')
bot.press('tab')
bot.press('tab')
bot.press('tab')
bot.press('tab')
bot.press('space')
bot.press('tab')
bot.press('enter')
bot.sleep(2.2)
bot.press('enter')
bot.sleep(2)
bot.hotkey('winleft','shift','s')
start_x, start_y = 47, 524
end_x, end_y = 1872, 683
bot.sleep(1.3)
bot.moveTo(start_x, start_y)
bot.dragTo(end_x, end_y, duration=1)
bot.sleep(1.5)
bot.click(1183, 607)
bot.click(1183, 607)
bot.click(1183, 607)
bot.hotkey('ctrl', 'c')
vlrctrlc = pyperclip.paste()
vlrctrlc = vlrctrlc.split('/')[0]
vlrctrlc = re.sub(r'[\\/*?:"<>|]', '', vlrctrlc)
vlrctrlc = vlrctrlc.strip()


bot.sleep(1.5)

pastas = os.listdir(pasta)

for i in pastas:
    if i[:7] == 'Captura':
        nome_arquivo = str(i)
        caminho_antigo = os.path.join(pasta, i)
        novo_nome = i.replace(nome_arquivo, vlrctrlc+'.png')
        caminho_novo = os.path.join(pasta, novo_nome)
        os.rename(caminho_antigo, caminho_novo)

bot.sleep(0.2)
bot.click(1815, 110)
bot.sleep(0.2)
bot.click(1693, 445)
bot.click(1693, 445)
bot.sleep(0.2)
bot.hotkey('winleft','r')
bot.sleep(0.3)
bot.press('enter')



# bot.click(938, 1061)
# bot.sleep(0.5)
# bot.press('f5')
# bot.write('2GO')
# bot.press('f2')
# bot.write(pyperclip.paste())
# bot.press('enter')










