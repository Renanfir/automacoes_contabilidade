import time
import pyautogui as bot
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
import xlrd

bot.click(1804, 17)
bot.sleep(1.3)
bot.hotkey('ctrl','f')
bot.write('cnpj')
caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\Genesys_FAT.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('informacoes')
for row in range(sheet.nrows):
    nome = sheet.cell_value(row, 0)
    cnpj = sheet.cell_value(row, 1)

    localiza_CNPJ = bot.locateOnScreen(r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\cnpj.jpeg' ,confidence=0.8)
    if localiza_CNPJ:
        bot.click(402, 384)
        bot.click(402, 384)
        bot.write(cnpj)
        bot.moveTo(415, 433)
        bot.sleep(4)
        bot.press('tab')
        bot.press('tab')
        bot.press('tab')
        bot.press('enter')
        bot.sleep(2.5)
        bot.write(nome)
        bot.press('enter')
        bot.sleep(0.3)
        bot.hotkey('shift','tab')
        bot.press('enter')
        bot.sleep(5)
        
    
    
    
