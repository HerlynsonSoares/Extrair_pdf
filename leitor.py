import os
import fitz  
import pytesseract
from pdf2image import convert_from_path
from PyPDF2 import PdfReader, PdfWriter
import pandas as pd

# Configuração do Tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\XXXXXX\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Users\XXXXXXXXX\AppData\Local\Programs\Tesseract-OCR\tessdata"

# Caminho do Poppler
poppler_path = r"C:\Users\XXXXXXXXXX\Documents\poppler-25.07.0\Library\bin"  

# Lista de colaboradores do Excel
df = pd.read_excel("colaboradores.xlsx")
colaboradores = df.iloc[:, 0].dropna().tolist()  

pasta_entrada = r"C:\Users\XXXXXXXXXXXXXXXX\Documents\Leitor_Pdf\Entrada_arquivos"
pasta_saida = r"C:\Users\XXXXXXXXXXXXXXXXX\Documents\Leitor_Pdf\Saida_arquivos"

# Criar pasta saída se não existir
os.makedirs(pasta_saida, exist_ok=True)

# Controle de colaboradores encontrados
colaboradores_encontrados = set()

def extrair_paginas_otimizado(arquivo_pdf, colaboradores, pasta_saida, poppler_path=None):
    nome_pdf = os.path.splitext(os.path.basename(arquivo_pdf))[0]

    pdf_reader = PdfReader(arquivo_pdf)
    total_paginas = len(pdf_reader.pages)
    doc = fitz.open(arquivo_pdf)

    paginas_sem_texto = []

    for i in range(total_paginas):
        texto = doc[i].get_text().strip()
        if not texto:
            paginas_sem_texto.append(i)
        else:
            for nome in colaboradores:
                if nome.lower() in texto.lower():
                    colaboradores_encontrados.add(nome)
                    writer = PdfWriter()
                    writer.add_page(pdf_reader.pages[i])
                    nome_arquivo = os.path.join(
                        pasta_saida, f"{nome.replace(' ', '_')}_{nome_pdf}_pag{i+1}.pdf"
                    )
                    with open(nome_arquivo, "wb") as f:
                        writer.write(f)
                    print(f"Página {i+1} (texto nativo) salva para {nome} -> {nome_arquivo}")

    if paginas_sem_texto:
        imagens = convert_from_path(arquivo_pdf, dpi=200, poppler_path=poppler_path)
        for i in paginas_sem_texto:
            imagem_rgb = imagens[i].convert("RGB")
            texto = pytesseract.image_to_string(imagem_rgb, lang="por")
            for nome in colaboradores:
                if nome.lower() in texto.lower():
                    colaboradores_encontrados.add(nome)
                    writer = PdfWriter()
                    writer.add_page(pdf_reader.pages[i])
                    nome_arquivo = os.path.join(
                        pasta_saida, f"{nome.replace(' ', '_')}_{nome_pdf}_pag{i+1}.pdf"
                    )
                    with open(nome_arquivo, "wb") as f:
                        writer.write(f)
                    print(f"Página {i+1} (OCR) salva para {nome} -> {nome_arquivo}")


# ===== EXECUTAR PARA TODOS OS PDFs =====
for arquivo in os.listdir(pasta_entrada):
    if arquivo.lower().endswith(".pdf"):
        caminho_pdf = os.path.join(pasta_entrada, arquivo)
        print(f"\n📂 Processando: {caminho_pdf}")
        extrair_paginas_otimizado(caminho_pdf, colaboradores, pasta_saida, poppler_path=poppler_path)

# ===== RELATÓRIO FINAL =====
colaboradores_nao_encontrados = set(colaboradores) - colaboradores_encontrados
relatorio_path = os.path.join(pasta_saida, "relatorio.txt")

with open(relatorio_path, "w", encoding="utf-8") as f:
    f.write("=== RELATÓRIO DE PROCESSAMENTO ===\n\n")
    f.write(f"Total de colaboradores na lista: {len(colaboradores)}\n")
    f.write(f"Total encontrados: {len(colaboradores_encontrados)}\n")
    f.write(f"Total não encontrados: {len(colaboradores_nao_encontrados)}\n\n")

    f.write("⚠️ Colaboradores não encontrados:\n")
    for nome in sorted(colaboradores_nao_encontrados):
        f.write(f"- {nome}\n")

print(f"\n📑 Relatório gerado em: {relatorio_path}")