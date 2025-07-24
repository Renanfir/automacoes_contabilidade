#BAIXAR PDF DO BALANÇO, ARRUMAR AS COLUNAS HORIZONTALMENTE, PARA OS DADOS FICAREM
#EM APENAS UMA PAGINA, SALVAR O ARQUIVO NA PASTA "TESTE BALANCETE - BALANCO AUTOMATICO"
#ALTERAR O NOME ao final "path", fazer o primeiro lançamento:
# D - 3335 / 3334 etc
# C - 7545
# H - 326

# Após esse lançamento, D - 7545, e pare o cursor no C que o programa continuará.
# Ao final fazer o lançamento dos lucros acumulados(valor ao final da folha):
# C - 3244
# H - 814 - TRANSFERENCIA PARA A CONTA DE LUCROS ACUMULADOS 
# Conferir os zeramentos balançete

import pdfplumber
import pyautogui as bot

# bot.click(1806, 16)
#bot.sleep(1.5)

path = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\bots_contabilidade\TESTE BALANCETE - BALANCO AUTOMATICO\balancete.pdf'

listaContas = [
    "6303", "6305", "1179", "3212", "6306", "6304", "2113", "6511", "6497", "6474", "6468", 
    "6563", "6451", "6481", "6570", "6942", "7485", "1542", "6513", "6586", "2983", "6882", 
    "6706", "1234", "1235", "50055", "7008", "7114", "1000", "65123","58271","1101","1518",
    "886", "6831", "6669", "66111", "57997", "7048", "7108","6528", "6401", "1680","66401",
    "6818","3215","3062","57981","710808","57171","67932","67151","68452","7115","7121",
    "6534","2116","2654","7077","66275","66418","2951","66097","1258","6500","6729",
    "1254","66281"

]


with pdfplumber.open(path) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            for line in text.split("\n"):
                if not line[0] == "*" and line[0].isdigit():
                    if line.split()[0] in listaContas:
                     if line[-1] == "D":
                        codigo = line.split()[0]
                        valor = line.split()[-1].rstrip("D").replace(".","")
                        print(codigo, valor)
                        # bot.write(codigo)
                        # bot.press('enter')
                        # bot.sleep(0.15)
                        # bot.write(valor)
                        # bot.press('enter')
                        # bot.sleep(0.15)
                        # bot.press('enter')
                        # bot.sleep(0.15)
                        # bot.press('enter')
                        # bot.sleep(0.15)
                        # bot.press('enter')
                        # bot.sleep(0.15)
                        # bot.press('enter')
                        # bot.sleep(0.15)
                        # bot.press('enter')
                




