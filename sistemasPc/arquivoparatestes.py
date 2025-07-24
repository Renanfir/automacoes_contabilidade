import pyautogui as bot
import os
from os import listdir
import pyperclip


pasta = r'C:\Users\Usuario\Pictures\Screenshots'
vlrctrlc = pyperclip.paste()
bot.sleep(1.5)

pastas = os.listdir(pasta)

for i in pastas:
    if i[:7] == 'Captura':
        caminho_antigo = os.path.join(pasta, i)
        novo_nome = i.replace('Captura', vlrctrlc)
        caminho_novo = os.path.join(pasta, novo_nome)
        os.rename(caminho_antigo, caminho_novo)



