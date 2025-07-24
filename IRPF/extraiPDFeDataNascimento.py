import fitz
import re
import os

# Definir o caminho da pasta contendo os arquivos PDF
pasta = 'C:\\Users\\Usuario\\PROGRAMACAO_CENTRAL\\bots_Contabilidade\\testespdf'

# Definir padrões regex para CPF e data de nascimento
cpf_pattern = r'\b\d{3}\.\d{3}\.\d{3}-\d{2}\b'
data_nascimento_pattern = r'\b\d{2}/\d{2}/\d{4}\b'

# Inicializar listas para armazenar os resultados
resultados = []

# Iterar sobre todos os arquivos na pasta
for arquivo in os.listdir(pasta):
    if arquivo.endswith('.pdf'):
        caminho_arquivo = os.path.join(pasta, arquivo)
        
        # Abrir o documento PDF
        documento = fitz.open(caminho_arquivo)
        
        # Inicializar uma variável para armazenar o texto
        texto = ''
        
        # Iterar sobre todas as páginas do PDF
        for pagina in documento:
            texto += pagina.get_text()
        
        # Usar regex para encontrar o CPF e a data de nascimento
        cpf = re.findall(cpf_pattern, texto)
        data_nascimento = re.findall(data_nascimento_pattern, texto)
        
        # Adicionar os resultados à lista
        resultados.append({
            'arquivo': arquivo,
            'cpf': cpf[0] if cpf else 'Não encontrado',
            'data_nascimento': data_nascimento[0] if data_nascimento else 'Não encontrado'
        })

# Exibir os resultados
for resultado in resultados:
    print(f"Arquivo: {resultado['arquivo']}")
    print(f"CPF: {resultado['cpf']}")
    print(f"Data de Nascimento: {resultado['data_nascimento']}")
    print('-' * 40)
