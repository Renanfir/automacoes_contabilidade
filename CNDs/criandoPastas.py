import xlrd
import os

caminho_arquivo = r'C:\Users\Usuario\PROGRAMACAO_CENTRAL\python\bots_Contabilidade\bots_contabilidade\CNDs\Genesys_FAT.xls'
workbook = xlrd.open_workbook(caminho_arquivo)
sheet = workbook.sheet_by_name('Plan1')          
for row in range(sheet.nrows):
    nome = sheet.cell_value(row, 0)
    cnpj = sheet.cell_value(row, 1)
    caminho = f"G:\RENAN\CERIDÕES NEGATIVAS\{nome}"

    os.makedirs(caminho, exist_ok=True)

