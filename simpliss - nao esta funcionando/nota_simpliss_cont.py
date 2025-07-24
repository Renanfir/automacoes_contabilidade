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


#Entra no google
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
driver.get("https://google.com")
driver.maximize_window()

wdw = WebDriverWait(driver, 15)
visibilidade = EC.visibility_of_element_located

#Escreve o link no url
url = f"file:///C:/Users/Usuario/Downloads/30.647.741_0001-20%20_%20-%20SIMPLISS%20WEB.html"
driver.get(url)
selec = wdw.until(visibilidade((By.XPATH, '//*[@id="wrap"]/div[2]/div/div[2]/ul/li[1]/div/div/ul[1]/li[2]/a'))).click()
bot.sleep(2)
link_url = f"https://nfse.blumenau.sc.gov.br/contrib/08955625000119/AspAccess/Index?id=nfse/previa_tipo_guia.aspx"
driver.execute_script("window.open(arguments[0], '_blank');", link_url)


#Preenche o CPF
campo_cpf = driver.find_element(By.ID, 'Login_UserName')
campo_cpf.send_keys('')

#Preenche a senha EXCLUIR A SENHA SE POSTAR NO GITHUB!!!!!!!
campo_senha = driver.find_element(By.ID, "Login_Password")
campo_senha.send_keys('')

#Clica no botão enter
acessar = driver.find_element(By.XPATH, '//*[@id="conteudo-principal"]/form/input[6]').click()

#Seleciona contador
caixa_selecao = driver.find_element(By.XPATH, '//*[@id="TipoAcesso_Regra"]').click()
bot.press('down')
bot.press('enter')

#Clica no botão enter
entrar = driver.find_element(By.XPATH, '//*[@id="conteudo-principal"]/form/div[2]/div/button').click()

time.sleep(1.5)


#Muda numero de empresas na tela de 10 para 100
caixa_selecao = wdw.until(visibilidade((By.XPATH, '//*[@id="DataTable_length"]/label/select')))
selecao = Select(caixa_selecao)
selecao.select_by_value("100")


#Seleciona ordem de nome
razao_social = driver.find_element(By.XPATH, '//*[@id="DataTable"]/thead/tr/th[5]')
razao_social.click()
razao_social.click()

cnpj = '30647741000120'
xpath = f'//img[@title="Selecionar" and contains(@onclick, "{cnpj}")]'
button = wdw.until(visibilidade((By.XPATH, xpath))).click()



bot.sleep(20)
