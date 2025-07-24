import pyautogui as bot
import xlrd

bot.click(1806, 17)
bot.sleep(1)
caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\bots_contabilidade\CNDs\Genesys_FAT.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('Plan1')

for row in range(sheet.nrows):

    nome = sheet.cell_value(row, 0)
    cnpj = sheet.cell_value(row, 1)
    bot.click(424, 493)
    bot.hotkey('ctrl', 'a')
    bot.press('delete')
    bot.write(cnpj)
    bot.sleep(0.3)
    bot.click(1454, 685)
    bot.sleep(3)
    bot.click(1098, 649)
    bot.sleep(7)
    bot.write(nome + ' - FEDERAL')
    bot.press('enter')
    bot.sleep(1.2)
    bot.click(1457, 704)
    bot.sleep(1)

