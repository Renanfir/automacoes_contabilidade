from pathlib import Path
import shutil

pastaArquivos = Path(r"C:\Users\Usuario\Downloads")
pastaPastas = Path(r"C:\AUREO\fazendo")

print()

for arquivo in pastaArquivos.iterdir():
    for pasta in pastaPastas.iterdir():
        if arquivo.name[:-14] == pasta.name:
            shutil.move(pastaArquivos.joinpath(arquivo.name), pastaPastas.joinpath(pasta.name))
