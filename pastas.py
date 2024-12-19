import os
import shutil
import pdfplumber
import re


# Função para identificar o modelo do PDF
def identificar_modelo_pdf(pdf_path):
    # Função para ler o conteúdo do PDF com pdfplumber
    texto = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            texto += page.extract_text()
          
    # Exemplo: Identificando o modelo com base em uma palavra-chave
    if re.search(r'ACESSORIAS', texto):
        return 'modelo_1'
    elif re.search(r'CATHO', texto):
        return 'modelo_2'
    elif re.search(r'CLICKSIGN', texto):
        return 'modelo_3'
    elif re.search(r'NEW VALE', texto):
        return 'modelo_5'
    elif re.search(r'SIEG', texto):
        return 'modelo_6'
    elif re.search(r'THOMSON', texto):
        return 'modelo_7'
    elif re.search(r'UNIMED', texto):
        return 'modelo_8'
    elif re.search(r'FLASH', texto):
        return 'modelo_9'
    else:
        return 'modelo_desconhecido'

# Função para mover o arquivo para a pasta correta
def mover_arquivo(pdf_path, destino):
    try:
        if not os.path.exists(destino):
            os.makedirs(destino)
        shutil.move(pdf_path, destino)
        print(f"Arquivo movido para: {destino}")
    except Exception as e:
        print(f"Erro ao mover o arquivo {pdf_path}: {e}")

# Função principal para processar a pasta
def processar_pasta_pdfs(pasta_origem, pasta_destino_modelo1, pasta_destino_modelo2 , pasta_destino_modelo3, pasta_destino_modelo5, pasta_destino_modelo6, pasta_destino_modelo7 , pasta_destino_modelo8 , pasta_destino_modelo9):
    for arquivo in os.listdir(pasta_origem):
        if arquivo.endswith(".pdf"):
            caminho_pdf = os.path.join(pasta_origem, arquivo)
            modelo = identificar_modelo_pdf(caminho_pdf)

            if modelo == 'modelo_1':
                mover_arquivo(caminho_pdf, pasta_destino_modelo1)
            elif modelo == 'modelo_2':
                mover_arquivo(caminho_pdf, pasta_destino_modelo2)
            elif modelo == 'modelo_3':
                mover_arquivo(caminho_pdf, pasta_destino_modelo3)
            elif modelo == 'modelo_5':
                mover_arquivo(caminho_pdf, pasta_destino_modelo5)
            elif modelo == 'modelo_6':
                mover_arquivo(caminho_pdf, pasta_destino_modelo6)
            elif modelo == 'modelo_7':
                mover_arquivo(caminho_pdf, pasta_destino_modelo7)
            elif modelo == 'modelo_8':
                mover_arquivo(caminho_pdf, pasta_destino_modelo8)
            elif modelo == 'modelo_9':
                mover_arquivo(caminho_pdf, pasta_destino_modelo9)
            else:
            
                print(f"Modelo desconhecido para o arquivo: {arquivo}")

# Definir os caminhos das pastas
pasta_origem = "Z:\\INFORMATICA\\Testes Notas"
pasta_destino_modelo1 = "Z:\\INFORMATICA\\Testes Notas\\ACESSORIAS"
pasta_destino_modelo2 = "Z:\\INFORMATICA\\Testes Notas\\CATHO"
pasta_destino_modelo3 = "Z:\\INFORMATICA\\Testes Notas\\CLICKSIGN"
pasta_destino_modelo5 = "Z:\\INFORMATICA\\Testes Notas\\NEWVALE"
pasta_destino_modelo6 = "Z:\\INFORMATICA\\Testes Notas\\SIEG"
pasta_destino_modelo7 = "Z:\\INFORMATICA\\Testes Notas\\THOMSON"
pasta_destino_modelo8 = "Z:\\INFORMATICA\\Testes Notas\\UNIMED"
pasta_destino_modelo9 = "Z:\\INFORMATICA\\Testes Notas\\FLASH"

# Chamar a função para processar a pasta de PDFs
processar_pasta_pdfs(pasta_origem, pasta_destino_modelo1, pasta_destino_modelo2 , pasta_destino_modelo3 , pasta_destino_modelo5, pasta_destino_modelo6 , pasta_destino_modelo7 , pasta_destino_modelo8 , pasta_destino_modelo9)



