import sys
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
from twocaptcha import TwoCaptcha
solver = TwoCaptcha('')

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

driver.get("https://google.com")
driver.maximize_window()

wdw = WebDriverWait(driver, 15)
visibilidade = EC.visibility_of_element_located


bot.sleep(10)


#Escreve o link no url
url = f"https://sat.sef.sc.gov.br/tax.NET/Sat.CtaCte.Web/SolicitacaoCnd.aspx"
driver.get(url)

caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\bots_contabilidade\CNDs\Genesys_FAT.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('Plan2')
abas = driver.window_handles

for row in range(sheet.nrows):
    driver.refresh()
    cnpj = sheet.cell_value(row, 0)
    nome = sheet.cell_value(row, 1)

    try:
        result = solver.recaptcha(sitekey='6Lc84kQiAAAAAJR5eZ-COz0IM5GNx-g_XbdLoPVa', url='https://sat.sef.sc.gov.br/tax.NET/Sat.CtaCte.Web/SolicitacaoCnd.aspx')
    except Exception as e:
        sys.exit(e)
    else:
        buscar = wdw.until(visibilidade((By.XPATH, '//*[@id="Body_Main_Main_sepBusca_idnCnd_MaskedField"]'))).send_keys(cnpj)
        token = result['code']

        #Feito com ajuda do gpt
        # Injetar o token no campo 'g-recaptcha-response'
        driver.execute_script("""
            document.getElementById('g-recaptcha-response').style.display = 'block';  
            document.getElementById('g-recaptcha-response').value = arguments[0];
        """, token)
        bot.click(463, 550)
        try:
            solicitar = wdw.until(visibilidade((By.XPATH, '//*[@id="Body_Main_Main_ctnResultado_btnGerarCnd"]'))).click()

            bot.sleep(3)
            bot.write(nome + ' - ESTADUAL')
            bot.press('enter')
            voltar = wdw.until(visibilidade((By.XPATH, '//*[@id="Body_Main_Main_ctnResultado_btnVoltar"]'))).click()
        except:
            sys.exit(e)
        else:    
            bot.sleep(3.5)
            bot.screenshot(nome+'.png')
            #driver.close
            bot.middleClick(390, 16)
            driver.switch_to.window(abas[0])



