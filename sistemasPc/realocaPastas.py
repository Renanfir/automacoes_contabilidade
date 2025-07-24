import os
import shutil
from pathlib import Path

path = 'G:\\EMAIL VANESSA\\1 - IMPOSTOS\\2024\\10.2024'

# Caminho da pasta Downloads, que varia dependendo do sistema operacional
downloads_folder = str(Path.home() / 'Downloads')

# Verifica se o diretório de destino (Downloads) existe
if not os.path.exists(downloads_folder):
    os.makedirs(downloads_folder)

# Percorre os arquivos na pasta de origem
for i in os.listdir(path):
    item_path = os.path.join(path, i)
    
    # Verifica se o item é uma pasta (subdiretório)
    if os.path.isdir(item_path):
        for j in os.listdir(item_path):
            palavraInverso = j[::-1]
            if palavraInverso[0:4] == 'slx.':
                arquivo_origem = os.path.join(item_path, j)
                arquivo_destino = os.path.join(downloads_folder, j)
                
                # Copia o arquivo para o diretório de Downloads
                shutil.copy(arquivo_origem, arquivo_destino)
                print(f"Arquivo copiado para Downloads: {j}")
