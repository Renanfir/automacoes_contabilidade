import xlrd
import pyautogui as bot

caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\bots_contabilidade\Genesys_FAT.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('dctfweb')
bot.click(1800, 18)             #sai da tela do vscode
for row in range(sheet.nrows):
    value0 = sheet.cell_value(row, 0)
    bot.sleep(2)
    bot.click(880, 663)         #Seleciona o campo do cnpj
    bot.sleep(1)
    bot.write(value0)
    bot.sleep(1)
    bot.click(1054, 664)
    bot.click(1054, 664)
    bot.sleep(5)
    bot.hotkey('ctrl','f')
    bot.press('delete')
    bot.write('ir para a caixa postal')
    bot.sleep(0.5)
    try:
        localiza_caixa = bot.locateOnScreen(r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\bots_contabilidade\Captura de tela 2025-06-27 094441.png' ,confidence=0.8)
    except bot.ImageNotFoundException:
        localiza_caixa = None
    
    if localiza_caixa:
        bot.sleep(1)
        bot.click(1076, 629)
        bot.sleep(3)
    

       
    bot.hotkey('ctrl','f')
    bot.press('delete')
    bot.write('TERMO DE EXCLUSAO DO SIMPLES NACIONAL n')
    bot.sleep(0.5)

    try:
        localiza_msg = bot.locateOnScreen(r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\bots_contabilidade\Captura de tela 2025-06-26 192509.png' ,confidence=0.8)
    except bot.ImageNotFoundException:
        localiza_msg = None  

    if localiza_msg:
        x,y = bot.center(localiza_msg)
        bot.click(x,y)
        bot.sleep(2)
        bot.press('printscreen')
        bot.sleep(2)
        bot.moveTo(42, 436)
        bot.mouseDown()
        bot.moveTo(1856, 754)
        bot.mouseUp()
        bot.sleep(2.5)
    bot.sleep(10)
    bot.click(1632, 224)
    