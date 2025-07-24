import pyautogui as bot
import xlrd

bot.click(1805, 16)
bot.sleep(1.2)

caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\Genesys_FAT.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('dctfweb')

for row in range(sheet.nrows):
    bot.click(880, 663)
    bot.sleep(1)
    bot.write(cnpj)
    bot.sleep(1)
    bot.click(1054, 664)
    bot.click(1054, 664)
    bot.sleep(2)
    bot.click(79, 440)
    bot.click(100, 502)
    bot.sleep(1.5)
    bot.click(94, 366)
    bot.sleep(0.3)
    bot.click(82, 398)
    bot.sleep(1.5)
    bot.write("012025")
    bot.press('tab')
    bot.press('enter')
    bot.sleep(1.5)
    bot.click(1623, 223)
    bot.sleep(8)

    