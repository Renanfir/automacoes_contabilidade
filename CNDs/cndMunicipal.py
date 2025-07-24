#FORMATAR EXCLE COM COLUNA 1 DE CNPJ E COLUNA 2 PARA NOME

#NO TEMPO DE ESPERA APÓS ABRIR O GOOGLE
#habilite perguntar onde fazer o download
#desabilite mostrar download

#MODIFIQUE O NOME DO GRUPO MEDICO S/S E TIRE A BARRA

#MODIFIQUE OS NOMES IGUAIS, COMO CHARCOT E NATURAPRO, ADICIONANDO UM 2 NA FILIAL

#LEMBRE DE DIGITAR O CAPTCHA ENQUANTO O PROGRAMA É EXECUTADO

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

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

driver.get("https://google.com")
driver.maximize_window()

wdw = WebDriverWait(driver, 15)
visibilidade = EC.visibility_of_element_located



bot.sleep(6)


#Escreve o link no url
url = f"https://www.blumenau.sc.gov.br/cidadao/Pages/Siatu/CND/EmissaoCND.aspx?Especial=N"
driver.get(url)


caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\bots_contabilidade\CNDs\Genesys_FAT.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('Plan1')

for row in range(sheet.nrows):

    nome = sheet.cell_value(row, 0)
    cnpj = sheet.cell_value(row, 1)

    caixaTexto = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_ContentBody_cbkEmissaoCND_txtCpfCnpj"]/tbody/tr/td'))).click()
    bot.hotkey('ctrl','a')
    bot.press('delete')
    bot.sleep(2)
    bot.write(cnpj)     
    bot.press('tab')
    bot.press('pageup')
    bot.press('pageup')
    bot.press('down')
    bot.press('tab')
    bot.sleep(3.5)
    emitir = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_ContentBody_cbkEmissaoCND_btPesquisar"]'))).click()
    try:
        gravar = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_ContentBody_btnImprimir"]'))).click()
        bot.sleep(1.5)
        bot.write(nome + ' - MUNICIPAL')
        bot.press('enter')
        volta = wdw.until(visibilidade((By.XPATH, '//*[@id="ctl00_divMenuPortalDinamico"]/dd[4]/a'))).click()
    except:
        print(nome)
        driver.refresh()
        continue

    