import xlrd
import pyautogui as bot

try:
    while True:
        caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\Genesys_FAT.xls'
        workbook = xlrd.open_workbook(caminho_arquivo)
        sheet = workbook.sheet_by_name('dctfweb')
        bot.click(1800, 18)             
        for row in range(sheet.nrows):
            value0 = sheet.cell_value(row, 0)
            value1 = sheet.cell_value(row, 1)
            bot.sleep(3)
            bot.click(880, 663)         #Seleciona o campo do cnpj
            bot.sleep(2)
            bot.write(value0)
            bot.sleep(2)
            bot.click(1054, 664)
            bot.click(1054, 664)        # Aperta no botão de entrar
            bot.sleep(3)
            bot.click(1625, 362)
            bot.sleep(0.3)
            bot.click(1429, 683)
            bot.sleep(0.155881196000162)
            bot.press('end')
            bot.press('end')
            bot.press('end')
            bot.press('end')
            bot.sleep(1)
            bot.click(618, 612)
            bot.sleep(1)
            bot.click(233, 615)
            bot.hotkey('ctrl','a')
            bot.press('delete')
            bot.write(value1)
            # bot.press('backspace')
            # bot.press('backspace')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.press('tab')
            bot.write('')
            bot.sleep(2)
            bot.click(1774, 923)
            bot.sleep(2)
            bot.press('home')
            bot.press('home')
            bot.press('home')
            bot.press('home')
            bot.sleep(3.5)
            bot.click(1611, 224)
            




except KeyboardInterrupt:
    print("Programa interrompido pelo usuário.")