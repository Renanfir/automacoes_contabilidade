import os
import PyPDF2
pdf_dir = 'C:\\Users\\Usuario\\Documents\\NovaPasta'
texto_completo = ''

for nome_arquivo in os.listdir(pdf_dir):
    caminho_pdf = os.path.join(pdf_dir, nome_arquivo)
    with open(caminho_pdf, 'rb') as f:
            pdf = PyPDF2.PdfReader(f)
            texto_pdf = ''
            for page in pdf.pages:
                texto_pdf += page.extract_text()
            
            texto_completo += texto_pdf
            if 'POSITIVA' in texto_pdf:
                print(nome_arquivo)
            
