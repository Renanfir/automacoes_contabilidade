#Fazer uma rodada de forma manual para configuração dos campos.
#Formatar cnpj apenas os numeros.
#Utilizar google chrome

import xlrd
import pyautogui as bot

try:
    while True:
        caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\bots_contabilidade\Genesys_FAT.xls'
        workbook = xlrd.open_workbook(caminho_arquivo)
        sheet = workbook.sheet_by_name('dctfweb')
        bot.click(1800, 18)             #sai da tela do vscode
        for row in range(sheet.nrows):      

            bot.sleep(1)
            value0 = sheet.cell_value(row, 0)
            bot.click(880, 663)         #Seleciona o campo do cnpj
            bot.sleep(1)
            bot.write(value0)
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
            bot.write('072025')
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
            bot.click(1155, 445)        #Benef não ident
            bot.sleep(0.5)
            bot.click(131, 443)         #Período
            bot.sleep(0.5)
            bot.write('072025')
            bot.sleep(0.5)
            bot.click(336, 445)         #CNPJ
            bot.write(value0)
            bot.sleep(0.5)
            bot.click(97, 500)
            bot.sleep(1)
            bot.click(275, 492)         #Natureza do rendimento
            bot.sleep(0.5)
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('down')
            bot.press('down')
            bot.press('down')
            bot.press('tab')
            bot.press('enter')


            bot.sleep(0.5)
            bot.click(319, 523)         #Detalhamento dos pagamentos
            bot.sleep(0.5)
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.write('11072025')
            bot.press('tab')
            bot.write('1')
            bot.press('tab')
            bot.press('tab')
            bot.write('NAO IDENTIFICADO')
            bot.press('tab')
            
            bot.press('enter')

            # bot.click(272, 608)         #Data
            # bot.write('08082024')
            # bot.sleep(0.5)
            # bot.click(756, 608)         #Valor
            # bot.write('1')
            # bot.sleep(0.5)
            # bot.click(285, 705)        #Descrição
            # bot.write('NAO IDENTIFICADO')
            # bot.sleep(0.5)
            # bot.click(237, 778)        #Enter

            bot.sleep(0.5)
            bot.click(221, 721)        #Concluir a pagina
            bot.sleep(5)
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
