import pyautogui as bot
import xlrd

caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\bots_Contabilidade\dadosefd.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('Plan1')

bot.click(1803, 17)
bot.sleep(1)


for row in range(sheet.nrows):
    cnpj = sheet.cell_value(row, 0)
    bot.click(56, 65)
    bot.sleep(0.2)
    bot.write(cnpj)
    bot.press('enter')
    bot.sleep(1.7)
    bot.press('enter')
    bot.sleep(1.7)
    bot.press('enter')
    bot.sleep(1.7)

    
    
    