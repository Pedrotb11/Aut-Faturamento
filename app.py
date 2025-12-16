import os
import fitz  # PyMuPDF
import shutil
from tqdm import tqdm
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import re
from tkinter import Tk, filedialog
import subprocess
import sys

# Abre o Dialog para selecionar a pasta
root = Tk()
root.withdraw()
folder_path = filedialog.askdirectory(title="Selecione a pasta com os PDFs")

if not folder_path:
    print("Nenhuma pasta selecionada.")
    exit()

# Caminho para o Poppler no Windows
POPPLER_PATH = r"C:/poppler/poppler-25.12.0/Library/bin"  # Atualize este caminho conforme necessário

# Assume que o Tesseract virá junto na pasta do projeto
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


PASTA_ORIGEM = folder_path #pasta selecionada pelo dialog
PASTA_DESTINO = "SEPARADOS"

os.makedirs(PASTA_DESTINO, exist_ok=True)

# Mapeia variações e erros comuns para o nome correto
NOMES_CORRETOS = {
    "DR DANIEL CAMPELO": "DR. DANIEL CAMPELO",
    "DR. DANIEL CAMPELO'": "DR. DANIEL CAMPELO",

    "DRA ANA PAULA SIMÕES": "DRA. ANA PAULA SIMÕES",
    "DRA. ANA PAULA SIMOES": "DRA. ANA PAULA SIMÕES",
    "DRA ANA PAULA SIMOES": "DRA. ANA PAULA SIMÕES",
    "DRA. ANA PAULA SIMÓES": "DRA. ANA PAULA SIMÕES",
    "DRA. ANA PAULA SIMÓES'" : "DRA. ANA PAULA SIMÕES",
    "DRA. ANA PAULA SIMÕES -": "DRA. ANA PAULA SIMÕES",
    "DRA, ANA PAULA SIMÕES" : "DRA. ANA PAULA SIMÕES",

    "DR PEDRO PESSANHA": "DR. PEDRO PESSANHA",

    "DRA CAMILA GALINDO": "DRA. CAMILA GALINDO",
    "DR. CAMILA GALINDO": "DRA. CAMILA GALINDO",
    "DR CAMILA GALINDO": "DRA. CAMILA GALINDO",
    "DRA, CAMILA GALINDO": "DRA. CAMILA GALINDO",
    "DRA. CAMILA GALINDO'": "DRA. CAMILA GALINDO",
}

def normalizar_nome(anestesista_extraido):
    nome_original = anestesista_extraido.strip()
    nome_limpo = re.sub(r"[^\w\s]", "", nome_original)  # Remove pontuação
    nome_limpo = nome_limpo.upper()  # Converte para maiúsculas para comparar com as chaves
    
    # Busca na lista de nomes corretos
    nome_corrigido = NOMES_CORRETOS.get(nome_limpo)
    if nome_corrigido:
        return nome_corrigido
    # Se não encontrou, retorna o nome original
    return nome_original

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

def extrair_anestesista(pdf_path):
    if not os.path.exists(pdf_path):
        return "Desconhecido"

    doc = fitz.open(pdf_path)
    texto = ""
    for page in doc:
        texto += page.get_text()
    doc.close()

    if not texto.strip():
        texto = extrair_texto_ocr(pdf_path)

    for linha in texto.split("\n"):
        if "anestesista:" in linha.lower():
            match = re.search(r"anestesista:\s*(.+)$", linha, re.IGNORECASE)
            if match:
                return match.group(1).strip()
    return "Desconhecido"

def mover_para_pasta(pdf_path, anestesista):
    if not pdf_path:
        return

    try:
        nome_normalizado = normalizar_nome(anestesista)
        nome_pasta = re.sub(r'[<>:"/\\|?*]', '', nome_normalizado)
        destino = os.path.join(PASTA_DESTINO, nome_pasta)
        os.makedirs(destino, exist_ok=True)
        shutil.move(pdf_path, os.path.join(destino, os.path.basename(pdf_path)))

    except FileNotFoundError as e:
        if e.winerror == 3:
            print(f"[WinError 3] Caminho inválido para {pdf_path}. Movendo para 'Desconhecidos'")
            destino = os.path.join(PASTA_DESTINO, "Desconhecidos")
            os.makedirs(destino, exist_ok=True)
            try:
                shutil.move(pdf_path, os.path.join(destino, os.path.basename(pdf_path)))
            except Exception as move_error:
                print(f"Erro ao mover para pasta 'Desconhecidos': {move_error}")
        else:
            raise  # Se não for WinError 3, relança o erro para você ver


def processar_pdfs():
    arquivos = [f for f in os.listdir(PASTA_ORIGEM) if f.lower().endswith(".pdf")]
    for nome_arquivo in tqdm(arquivos, desc="Processando PDFs"):
        caminho = os.path.join(PASTA_ORIGEM, nome_arquivo)
        anestesista = extrair_anestesista(caminho)
        mover_para_pasta(caminho, anestesista)

if __name__ == "__main__":
    processar_pdfs()
    try:
        rota_rotate = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rotate.py")
        subprocess.run([sys.executable, rota_rotate], check=True)
    except Exception as e:
        print(f"Erro ao chamar rotate.py: {e}")
