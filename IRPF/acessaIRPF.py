import pyautogui as bot
import xlrd

region = (550, 150, 900, 450)
bot.click(1803, 20)

# while True:
caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\bots_Contabilidade\IRPF\cpfsEdatas.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('Plan1')
for row in range(sheet.nrows):
        nome = sheet.cell_value(row, 0)
        cpf = sheet.cell_value(row, 1)
        ddn = sheet.cell_value(row, 2)

        if(cpf == 'Não encontrado' or nome == 'Não encontrado'):
                print(nome)
        else:   
            bot.sleep(1)
            bot.click(604, 252)
            bot.hotkey('ctrl', 'a')
            bot.press('Backspace')
            bot.write(cpf)
            bot.sleep(1)
            bot.press('tab')
            for a in range(10):
                   bot.press('delete')
            bot.write(ddn)
            bot.sleep(7)
            bot.click(969, 622)
            bot.sleep(2.5)
            screenshot = bot.screenshot(region=region)
            screenshot.save(f'C:\\Users\\Usuario\\Pictures\\Screenshots\\IRPF\\{nome}.png')
            bot.sleep(2)
            bot.click(585, 116)