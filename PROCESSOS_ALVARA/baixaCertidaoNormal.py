import sys
import time
import pyautogui as bot
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import xlrd

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

driver.get("https://google.com")
driver.maximize_window()

wdw = WebDriverWait(driver, 30)
visibilidade = EC.visibility_of_element_located

#Escreve o link no url
url = f"https://www.blumenau.sc.gov.br/cidadao/Pages/Siatu/CND/EmissaoCND.aspx?Especial=N"
driver.get(url)
bot.sleep(5)

caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\cnpjPrefeitura.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('Plan1')


for row in range(sheet.nrows):
    cnpj = sheet.cell_value(row, 0)
    nome = sheet.cell_value(row, 1)
    bot.sleep(1)
    Campocnpj = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_ContentBody_cbkEmissaoCND_txtCpfCnpj_I"]'))).send_keys(cnpj)
    bot.press('tab')
    bot.press('down')
    bot.press('tab')
    bot.sleep(3.5)
    BotaoEnvia = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_ContentBody_cbkEmissaoCND_btPesquisar_CD"]'))).click()
    bot.sleep(7)
    BotaoImprime = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_ContentBody_btnImprimir"]'))).click()
    bot.sleep(2.5)
    bot.write(nome)
    bot.press('enter')
    bot.sleep(1)
    CertidaoNormal = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_divMenuPortalDinamico"]/dd[5]/a'))).click()

            



