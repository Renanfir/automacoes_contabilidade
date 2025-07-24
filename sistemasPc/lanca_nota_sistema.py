#Abra o pyautogui em tela cheia, sobre a tela cheia do lançamento e preencha o nome da planilha conforme a linha 7


import xlrd
import pyautogui as bot


caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\bots_Contabilidade\TESTENOTAS.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('Plan1')

bot.click(1804, 17)
for z in range(1):
    for row in range(sheet.nrows):
        value0 = sheet.cell_value(row, 0)

        value1 = sheet.cell_value(row, 1)

        value1 = int(value1)
        value1 = str(value1)

        value2 = sheet.cell_value(row, 2)
        value2 = str(value2)

        print(value0, value1, value2)

        bot.click(337, 497)
        bot.hotkey('ctrl', 'a')
        bot.press("Backspace")
        bot.write(value0)
        bot.click(336, 603)
        bot.write(value1)
        bot.press("enter")
        bot.press("enter")
        bot.press("enter")
        bot.press("enter")
        bot.press("enter")
        bot.write(value2)

        for i in range(6):
            bot.press("enter")

        bot.write("0")
        bot.press("enter")
        bot.press("enter")
        bot.write("0")
        bot.press("enter")
        bot.write("0")
        bot.press("enter")
        bot.write("0")
        bot.press("enter")
        bot.press("enter")
        bot.write("0")
        bot.press("enter")
        bot.write("0")
        bot.press("enter")
        bot.write("0")
        bot.press("enter")
        bot.write("0")
        bot.press("enter")
        bot.press("enter")
        bot.press("enter")
        bot.write("0")
        bot.press("enter")
        bot.write("0")
        bot.press("enter")
        bot.press("enter")
