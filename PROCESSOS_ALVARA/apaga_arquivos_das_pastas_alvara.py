import os

#CAMINHOS
pasta_alvara = 'C:\ALVARA 2024'
pasta_vanessa = 'G:\\EMAIL VANESSA\\1 - IMPOSTOS\\2024\\03 - 2024'
pasta_isabela = 'G:\\EMAIL ISABELA\\2024\\03 - 2024'
pasta_aureo = 'G:\EMAIL AUREO'



def apagar_arquivos_recursivamente(caminho_pasta):
    # Verifica se o caminho dado é uma pasta
    if not os.path.isdir(caminho_pasta):
        print(f'O caminho "{caminho_pasta}" não é uma pasta válida.')
        return
    
    # Percorre recursivamente todas as pastas e arquivos
    for pasta_atual, _, arquivos in os.walk(caminho_pasta):
        for arquivo in arquivos:
            caminho_arquivo = os.path.join(pasta_atual, arquivo)
            try:
                os.remove(caminho_arquivo)
                print(f'Arquivo "{caminho_arquivo}" removido com sucesso.')
            except Exception as e:
                print(f'Ocorreu um erro ao excluir o arquivo "{caminho_arquivo}": {e}')

# Exemplo de uso
caminho_da_pasta_principal = 'G:\\EMAIL ISABELA\\2024\\03 - 2024'
apagar_arquivos_recursivamente(caminho_da_pasta_principal)