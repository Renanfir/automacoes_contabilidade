#Fazer uma rodada de forma manual para configuração dos campos.
#Formatar cnpj apenas os numeros.
#Utilizar google chrome

import xlrd
import pyautogui as bot

try:
    while True:
        caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\Genesys_FAT.xls'
        workbook = xlrd.open_workbook(caminho_arquivo)
        sheet = workbook.sheet_by_name('informacoes')
        bot.click(1800, 18)             #sai da tela do vscode
        for row in range(sheet.nrows):

            bot.sleep(1)
            cnpj = sheet.cell_value(row, 1)
            cpf = sheet.cell_value(row, 2)
            bot.click(880, 663)         #Seleciona o campo do cnpj
            bot.sleep(1)
            bot.write(cnpj)
            bot.sleep(1)
            bot.click(1054, 664)
            bot.click(1054, 664)        # Aperta no botão de entrar
            bot.sleep(3.2)
            bot.moveTo(249, 306)        #Retenções previdenciárias
            bot.sleep(1)
            bot.moveTo(224, 532)
            bot.click(224, 532)         #Fechamento do movimento
            bot.sleep(1)
            bot.click(169, 662)         #Informar fechamento sem movimento
            bot.sleep(1)
            bot.click(128, 418)         #Período de apuração
            bot.write('122024')
            bot.moveTo(938, 335)        
            bot.click(938, 335)         #Clica fora dos campos
            bot.sleep(0.5)
            bot.click(151, 522)         #CPF
            bot.write('')
            bot.sleep(0.5)
            bot.click(492, 515)         #Nome
            bot.write('')
            bot.sleep(0.5)
            bot.click(115, 1000)        #concluir e enviar
            bot.sleep(10.6)
            bot.moveTo(525, 329)        #Rendimentos pagos
            bot.sleep(0.5)
            bot.moveTo(489, 350)
            bot.click(489, 350)         #Incluir pagamento
            bot.sleep(0.5)
            
            bot.click(301, 467)        #Benef pessoa fisica
            bot.sleep(0.5)
            bot.press('tab')
            bot.write('122024')
            bot.press('tab')
            bot.write(cnpj)
            bot.press('tab')
            bot.write(cpf)
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('enter')
            bot.sleep(0.5)
            bot.press('tab')
            bot.press('tab')
            bot.press('enter')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('down')
            bot.press('down')
            bot.press('down')
            bot.press('tab')
            bot.press('down')
            bot.press('tab')
            bot.press('tab')
            bot.press('enter')
            bot.sleep(0.5)
            bot.click(383, 657)
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.write('31122024')
            bot.press('tab')
            bot.write('10000')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('enter')
            bot.press('end')
            bot.press('end')
            bot.press('end')
            bot.sleep(0.5)
            
            bot.hotkey('ctrl','f')
            bot.write('concluir e enviar')

            localiza_concluir = bot.locateOnScreen(r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\ecac\concluir.jpeg' ,confidence=0.8)
    
            if localiza_concluir:
                x,y = bot.center(localiza_concluir)
                bot.click(x,y)

            bot.sleep(6)
            bot.press('home')
            bot.press('home')
            bot.press('home')
            bot.moveTo(511, 329)       #Rendimentos pagos
            bot.sleep(0.5)
            bot.click(482, 392)        #Fechamento
            bot.sleep(0.8)
            bot.click(1388, 513)       #Fechar
            bot.sleep(0.5)
            bot.click(127, 512)        #CPF
            bot.write('')
            bot.sleep(0.5)
            bot.click(694, 512)        #Nome 
            bot.write('')
            bot.click(113, 634)        #Concluir e enviar
            bot.sleep(4.5)
            bot.click(1607, 222)       #Selecionar empresa
            bot.sleep(1)
            
except KeyboardInterrupt:
    print("Programa interrompido pelo usuário.")




    
