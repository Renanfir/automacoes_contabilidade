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
from bs4 import BeautifulSoup






#Entra no google
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
driver.get("https://google.com")
driver.maximize_window()
bot.sleep(7)
wdw = WebDriverWait(driver, 15)
visibilidade = EC.visibility_of_element_located

#Escreve o link no url
url = f"https://nfse.blumenau.sc.gov.br/contrib/Account/Login"
driver.get(url)

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

#Seleciona ISSQN no topo da página
nfse_element = driver.find_element(By.CLASS_NAME, "sistemaAtual")
ActionChains(driver).move_to_element(nfse_element).perform()

laranjinha = driver.find_element(By.CLASS_NAME, "col-md-2")
laranjinha.click()

# boneco = driver.find_element(By.XPATH, '//*[@id="wrap"]/div[2]/div/div[2]/div[1]/a[1]').click()
boneco = wdw.until(visibilidade((By.XPATH, '//*[@id="wrap"]/div[2]/div/div[2]/div[1]/a[1]'))).click()


#Muda numero de empresas na tela de 10 para 100
caixa_selecao = wdw.until(visibilidade((By.XPATH, '//*[@id="DataTable_length"]/label/select')))
selecao = Select(caixa_selecao)
selecao.select_by_value("100")


#Seleciona ordem de nome
razao_social = driver.find_element(By.XPATH, '//*[@id="DataTable"]/thead/tr/th[5]')
razao_social.click()
razao_social.click()



caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\cnpjPrefeitura.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('Plan1')

for row in range(sheet.nrows):
    cnpj = sheet.cell_value(row, 0)
    xpath = f'//img[@title="Selecionar" and contains(@onclick, "{cnpj}")]'
    button = wdw.until(visibilidade((By.XPATH, xpath))).click()
    bot.sleep(5)
    bot.hotkey('ctrl','f')
    bot.write('Relatorios')

    localiza_relatorio = bot.locateOnScreen(r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\simpliss\relatorios.jpeg' ,confidence=0.8)
    
    if localiza_relatorio:
        x,y = bot.center(localiza_relatorio)
        bot.moveTo(x,y)
        bot.sleep(2)

        bot.hotkey('ctrl','a')
        bot.write('Movimentacao Mensal')
        localiza_movimentacao = bot.locateOnScreen(r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\simpliss\movimentacao.jpeg' ,confidence=0.8)
        if localiza_movimentacao:
            x,y = bot.center(localiza_movimentacao)
            bot.click(x,y)
            bot.sleep(3)


#ACESSANDO A PRIMEIRA NOTA
        bot.click(1787, 93)
        bot.click(786, 585)
        bot.press('down')
        bot.press('enter')
        bot.press('tab')
        bot.press('tab')
        bot.write('01112024')
        bot.sleep(0.2)
        bot.write('30112024')
        bot.press('tab')
        bot.press('tab')
        bot.press('tab')
        bot.press('tab')
        bot.press('enter')
        bot.sleep(7.5)

#VERIFICANDO SE A NOTA ESTÁ VAZIA
        try:
            localiza_nota_vazia_prestado = bot.locateOnScreen(r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\simpliss\notaVaziaPrestado.jpeg', confidence=0.8)

            if localiza_nota_vazia_prestado:
                bot.click(710, 15)

#NOTA NÃO VAZIA
            else:
                bot.hotkey('ctrl', 'p')
                bot.sleep(1)
                bot.press('enter')
                bot.sleep(2)
                bot.click(710, 15)

#ACESSANDO PROXIMA NOTA
                bot.click(786, 585)
                bot.press('down')
                bot.press('down')
                bot.press('enter')
                bot.press('tab')
                bot.press('tab')
                bot.press('tab')
                bot.press('tab')
                bot.press('tab')
                bot.press('tab')
                bot.press('tab')
                bot.press('tab')
                bot.press('enter')
                bot.sleep(7.5)
                try:
                    localiza_nota_vazia_aquisicao = bot.locateOnScreen(r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\simpliss\notaVaziaAquisicao.jpeg', confidence=0.8)

                    if localiza_nota_vazia_aquisicao:
                        bot.click(710, 15)
                    else:
                        bot.hotkey('ctrl', 'p')
                        bot.sleep(1)
                        bot.press('enter')
                        bot.sleep(2)
                        bot.click(710, 15)

                except bot.ImageNotFoundException:
                    bot.hotkey('ctrl', 'p')
                    bot.sleep(1)
                    bot.press('enter')
                    bot.sleep(2)
                    bot.click(710, 15)
#NOTA NÃO VAZIA
        except bot.ImageNotFoundException:
            bot.hotkey('ctrl', 'p')
            bot.sleep(1)
            bot.press('enter')
            bot.sleep(2)
            bot.click(710, 15)

#ACESSANDO PROXIMA NOTA
            bot.click(786, 585)
            bot.press('down')
            bot.press('down')
            bot.press('enter')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('enter')
            bot.sleep(7.5)
            
            try:
                localiza_nota_vazia_aquisicao = bot.locateOnScreen(r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\simpliss\notaVaziaAquisicao.jpeg', confidence=0.8)

                if localiza_nota_vazia_aquisicao:
                    bot.click(710, 15)
                else:
                    bot.hotkey('ctrl', 'p')
                    bot.sleep(1)
                    bot.press('enter')
                    bot.sleep(2)
                    bot.click(710, 15)

            except bot.ImageNotFoundException:
                bot.hotkey('ctrl', 'p')
                bot.sleep(1)
                bot.press('enter')
                bot.sleep(2)
                bot.click(710, 15)


    bot.sleep(1000000)

    

        


    
