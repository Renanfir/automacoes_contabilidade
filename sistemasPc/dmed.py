import pyautogui as bot
import xlrd
import os

bot.click(1802, 17)
bot.sleep(2)

caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\bots_contabilidade\ISADORA.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('Plan3')
for row in range(sheet.nrows):
    cpf = str(sheet.cell_value(row, 0))
    nome = sheet.cell_value(row, 1)
    valor = sheet.cell_value(row, 2)
    # valor = valor[:-2]
    valor = f"{valor:.2f}".replace('.', ',')


    valor = str(valor)
    bot.write(cpf)
    bot.press('tab')
    bot.write(nome)
    bot.press('tab')
    bot.write(valor)
    bot.click(1832, 141)
    bot.sleep(0.2)

os._exit(0)