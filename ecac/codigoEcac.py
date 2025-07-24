import pyautogui as bot
import xlrd

bot.click(1807, 16)
bot.sleep(1.2)


caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\bots_contabilidade\Genesys_FAT.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('dctfweb')

for row in range(sheet.nrows):
    bot.click(311, 656)
    bot.sleep(0.3)
    bot.click(311, 656)
    cnpj = sheet.cell_value(row, 0)
    bot.write(cnpj)
    bot.sleep(1.5)
    bot.click(575, 659)
    bot.click(575, 659)
    bot.sleep(7)
    bot.moveTo(76, 297)
    bot.sleep(1)
    bot.click(95, 320)
    bot.sleep(3)
    bot.click(313, 459)
    bot.sleep(6)
    bot.click(731, 223)
    bot.sleep(4)