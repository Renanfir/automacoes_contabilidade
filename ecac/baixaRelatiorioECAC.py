import pyautogui as bot
import xlrd
codigo_js = '''
const link = [...document.querySelectorAll('a')].find(a => a.textContent.includes('Gerar Relatório'));
if (link) {
    link.click();
} else {
    console.log("Link não encontrado.");
}
'''

try:
    while True:
        caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\Genesys_FAT.xls'
        workbook = xlrd.open_workbook(caminho_arquivo)
        sheet = workbook.sheet_by_name('SIMPL')
        bot.click(1800, 18)             #sai da tela do vscode
        for row in range(sheet.nrows):

            bot.sleep(1)
            cnpj = sheet.cell_value(row, 0)
            nome = sheet.cell_value(row, 1)
            bot.click(880, 663)         #Seleciona o campo do cnpj
            bot.sleep(1)
            bot.write(cnpj)
            bot.sleep(1)
            bot.click(1054, 664)
            bot.click(1054, 664)        # Aperta no botão de entrar
            bot.sleep(14)


            bot.hotkey('ctrl','f')
            bot.write('Gerar Relatorio')
            localizacao_imagem = bot.locateOnScreen('gerarRelatorio.jpeg',confidence=0.8)
            try:
                if localizacao_imagem:
                    x,y = bot.center(localizacao_imagem)
                    bot.click(x,y)
                    bot.sleep(5)
            except bot.ImageNotFoundException:
                print("Não achou a imagem")



            bot.click(1253, 384)
            bot.sleep(4.5)
            bot.write(nome)

            bot.press('enter')
            bot.sleep(1.2)
            bot.click(1642, 222)
except KeyboardInterrupt:
    print("Programa interrompido pelo usuário.")