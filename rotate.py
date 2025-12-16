import os
import fitz  # PyMuPDF
import shutil
from tqdm import tqdm
import pytesseract
from PIL import Image
import re
from io import BytesIO
import subprocess
import sys

# Puxa o Tesseract do sistema
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Usa diretamente a pasta SEPARADOS/Desconhecido como origem (sem poppler)
PASTA_ORIGEM = os.path.join("SEPARADOS", "Desconhecido")
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
        # Gira o PDF em 180° e salva no próprio arquivo (sem duplicar)
        try:
            doc = fitz.open(pdf_path)
            for page in doc:
                page.set_rotation(180)  # define rotação absoluta para evitar acumular
            doc.save(pdf_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
            doc.close()
        except Exception as e:
            print(f"Erro ao rotacionar e salvar PDF: {e}")

        texto_total = ""
        doc = fitz.open(pdf_path)
        for page in doc:
            try:
                pix = page.get_pixmap()
                img = Image.open(BytesIO(pix.tobytes("png")))
                texto_total += pytesseract.image_to_string(img, lang='por') + "\n"
            except pytesseract.TesseractError as e:
                print(f"Erro no Tesseract ao processar imagem: {e}")
            except Exception as e:
                print(f"Erro ao gerar imagem para OCR: {e}")
        doc.close()
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
    if not os.path.isdir(PASTA_ORIGEM):
        print(f"Pasta de origem não encontrada: {PASTA_ORIGEM}")
        return

    arquivos = [f for f in os.listdir(PASTA_ORIGEM) if f.lower().endswith(".pdf")]
    for nome_arquivo in tqdm(arquivos, desc="Processando PDFs em Desconhecido"):
        caminho = os.path.join(PASTA_ORIGEM, nome_arquivo)
        anestesista = extrair_anestesista(caminho)
        mover_para_pasta(caminho, anestesista)

if __name__ == "__main__":
    processar_pdfs()
    # Ao terminar, chama protocolo.py
    try:
        caminho_protocolo = os.path.join(os.path.dirname(os.path.abspath(__file__)), "protocolo.py")
        subprocess.run([sys.executable, caminho_protocolo], check=True)
    except Exception as e:
        print(f"Erro ao chamar protocolo.py: {e}")
