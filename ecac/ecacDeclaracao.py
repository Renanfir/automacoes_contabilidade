import pyautogui as bot
import xlrd

bot.click(1804, 15)
bot.sleep(1.2)

caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\bots_Contabilidade\ecac\nomescpffat.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('Plan1')
for row in range(sheet.nrows):
        nome = sheet.cell_value(row, 0)
        cnpj = sheet.cell_value(row, 1)
        
        
        bot.click(1640, 235)
        bot.sleep(0.5)
        bot.click(887, 666)
        bot.write(cnpj)
        bot.click(1064, 665)
        bot.click(1064, 665)
        bot.sleep(3)
        bot.click(79, 453)
        bot.sleep(1)
        bot.click(100, 515)
        bot.sleep(2)
        bot.click(67, 464)
        bot.sleep(0.15)
        bot.click(71, 500)
        bot.sleep(1.5)
        bot.click(68, 418)
        bot.sleep(1)
        bot.click(78, 513)
        bot.sleep(3.5)
        bot.click(1027, 432)
        bot.sleep(0.1)
        bot.click(932, 475)
        bot.sleep(0.5)
        bot.click(949, 347)
        bot.sleep(3.8)
        bot.write(nome)
        bot.press('enter')
        bot.sleep(4)