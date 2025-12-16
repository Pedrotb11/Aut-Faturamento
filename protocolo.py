import os
import fitz
import pytesseract
from pdf2image import convert_from_path
import re
from openpyxl import Workbook
from tqdm import tqdm

# Caminho para o Poppler no Windows
POPPLER_PATH = r"C:/poppler/poppler-25.12.0/Library/bin"

# Puxa o Tesseract do sistema
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Pasta base onde estão as subpastas já separadas por anestesista
BASE_SEPARADOS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "SEPARADOS")

def extrair_texto_ocr(pdf_path):
    try:
        imagens = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)
        texto_total = ""
        for img in imagens:
            try:
                img_rotacionada = img
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


def processar_pasta(pasta_pdfs):
    """Processa todos os PDFs de uma pasta e gera um Excel com os protocolos."""
    resultados = []

    for arquivo in tqdm(os.listdir(pasta_pdfs), desc=f"Processando PDFs em {os.path.basename(pasta_pdfs)}"):
        if arquivo.lower().endswith('.pdf'):
            caminho_pdf = os.path.join(pasta_pdfs, arquivo)
            dados = extrair_dados(caminho_pdf)
            resultados.append(dados)

    if not resultados:
        print(f"Nenhum PDF encontrado em {pasta_pdfs}. Pulando.")
        return

    # Salvar em Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "protocolos"  # se o codigo quebrar, mude pra "resultados"

    # Adicionar cabeçalhos
    ws.append(['Paciente', 'Convênio', 'Exame', 'Data'])

    # Adicionar dados
    for dados in resultados:
        ws.append(dados)

    # Salvar arquivo dentro da própria pasta do anestesista
    base_folder_name = os.path.basename(pasta_pdfs.rstrip("\\/"))
    output_filename = f"protocolos_{base_folder_name}.xlsx"
    caminho_saida = os.path.join(pasta_pdfs, output_filename)
    wb.save(caminho_saida)

    print(f"Extração concluída. Dados salvos em {caminho_saida}.")


if __name__ == "__main__":
    if not os.path.isdir(BASE_SEPARADOS):
        print(f"Pasta 'SEPARADOS' não encontrada em: {BASE_SEPARADOS}")
        raise SystemExit(1)

    # Percorre todas as subpastas dentro de SEPARADOS (um anestesista por pasta)
    for nome_pasta in os.listdir(BASE_SEPARADOS):
        caminho_pasta = os.path.join(BASE_SEPARADOS, nome_pasta)
        if os.path.isdir(caminho_pasta):
            print(f"\n=== Processando pasta: {caminho_pasta} ===")
            processar_pasta(caminho_pasta)
