import os
import fitz
import pytesseract
from pdf2image import convert_from_path
import re
from openpyxl import Workbook
from tkinter import Tk, filedialog
from tqdm import tqdm 

# Adicione esse parâmetro indicando o caminho do poppler
POPPLER_PATH = r'C:\poppler\poppler-24.08.0\Library\bin'

# Abre o Dialog para selecionar a pasta
root = Tk()
root.withdraw()
folder_path = filedialog.askdirectory(title="Selecione a pasta com os PDFs")

if not folder_path:
    print("Nenhuma pasta selecionada.")
    exit()

# Caminho para o Poppler no Windows
POPPLER_PATH = r"C:/poppler/poppler-24.08.0/Library/bin"

pytesseract.pytesseract.tesseract_cmd = os.path.join(os.path.dirname(__file__), "Tesseract","Tesseract-OCR", "tesseract.exe")

PASTA_PDFS = folder_path
RESULTADOS = []

def extrair_texto_ocr(pdf_path):
    try:
        imagens = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)
        texto_total = ""
        for img in imagens:
            try:
                img_rotacionada = img.rotate(180, expand=True)
                texto_total += pytesseract.image_to_string(img_rotacionada, lang='por') + "\n"
            except pytesseract.TesseractError as e:
                print(f"Erro no Tesseract ao processar imagem: {e}")
        return texto_total
    except Exception as e:
        print(f"Erro ao converter PDF para imagem: {e}")
        return ""

def extrair_campo(texto, campo):
    for linha in texto.split("\n"):
        if f"{campo}:" in linha.lower():
            match = re.search(rf"{campo}:\s*(.+)$", linha, re.IGNORECASE)
            if match:
                return match.group(1).strip()
    return ""

def extrair_dados(pdf_path):
    if not os.path.exists(pdf_path):
        return ["", "", "", ""]

    doc = fitz.open(pdf_path)
    texto = ""
    for page in doc:
        texto += page.get_text()
    doc.close()

    if not texto.strip():
        texto = extrair_texto_ocr(pdf_path)

    # Extrair campos que funcionam bem com a função atual
    convenio = extrair_campo(texto, "convênio")
    data     = extrair_campo(texto, "data")
    
    # Extrair paciente - pegar o que está antes de "data"
    paciente = ""
    for linha in texto.split("\n"):
        if "paciente:" in linha.lower():
            # Encontrar a posição de "data:" na mesma linha
            pos_data = linha.lower().find("data:")
            if pos_data != -1:
                # Pegar o conteúdo entre "paciente:" e "data:"
                match = re.search(r"paciente:\s*(.+?)\s*data:", linha, re.IGNORECASE)
                if match:
                    paciente = match.group(1).strip()
            else:
                # Se não encontrar "data:" na mesma linha, usar o método original
                match = re.search(r"paciente:\s*(.+)$", linha, re.IGNORECASE)
                if match:
                    paciente = match.group(1).strip()
            break
    
    # Extrair exame - pegar o que está antes de "anestesista"
    exame = ""
    for linha in texto.split("\n"):
        if "exame:" in linha.lower():
            # Encontrar a posição de "anestesista:" na mesma linha
            pos_anestesista = linha.lower().find("anestesista:")
            if pos_anestesista != -1:
                # Pegar o conteúdo entre "exame:" e "anestesista:"
                match = re.search(r"exame:\s*(.+?)\s*anestesista:", linha, re.IGNORECASE)
                if match:
                    exame = match.group(1).strip()
            else:
                # Se não encontrar "anestesista:" na mesma linha, usar o método original
                match = re.search(r"exame:\s*(.+)$", linha, re.IGNORECASE)
                if match:
                    exame = match.group(1).strip()
            break

    return [paciente, convenio, exame, data]

for arquivo in tqdm(os.listdir(PASTA_PDFS), desc="Processando PDFs"):
    if arquivo.lower().endswith('.pdf'):
        caminho_pdf = os.path.join(PASTA_PDFS, arquivo)
        dados = extrair_dados(caminho_pdf)
        RESULTADOS.append(dados)

# Salvar em Excel
wb = Workbook()
ws = wb.active
ws.title = "protocolos" #se o codigo quebrar, mude pra "resultados"

# Adicionar cabeçalhos
ws.append(['Paciente', 'Convênio', 'Exame', 'Data'])

# Adicionar dados
for dados in RESULTADOS:
    ws.append(dados)

# Salvar arquivo
base_folder_name = os.path.basename(PASTA_PDFS.rstrip("\\/"))
output_filename = f"protocolos_{base_folder_name}.xlsx"
wb.save(output_filename)

print(f"Extração concluída. Dados salvos em {output_filename}.")
