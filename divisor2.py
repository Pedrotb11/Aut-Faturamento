import os
import fitz  # PyMuPDF
import shutil
from tqdm import tqdm
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import re
from tkinter import Tk, filedialog

# Abre o Dialog para selecionar a pasta
root = Tk()
root.withdraw()
folder_path = filedialog.askdirectory(title="Selecione a pasta com os PDFs")

if not folder_path:
    print("Nenhuma pasta selecionada.")
    exit()

# Caminho para o Poppler no Windows
POPPLER_PATH = r"C:/poppler/poppler-24.08.0/Library/bin"  # Atualize este caminho conforme necessário

# Assume que o Tesseract virá junto na pasta do projeto
pytesseract.pytesseract.tesseract_cmd = os.path.join(os.path.dirname(__file__), "Tesseract","Tesseract-OCR", "tesseract.exe")

PASTA_ORIGEM = folder_path #pasta selecionada pelo dialog
PASTA_DESTINO = "SEPARADOS"
DEBUG_MODE = False  # Ative como True para salvar textos extraídos em arquivos de debug

os.makedirs(PASTA_DESTINO, exist_ok=True)
if DEBUG_MODE:
    os.makedirs("DEBUG_TEXTO", exist_ok=True)

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

def extrair_texto_ocr(pdf_path, tentar_rotacao=True):
    try:
        imagens = convert_from_path(pdf_path, poppler_path=POPPLER_PATH, dpi=300)
        texto_total = ""
        for img in imagens:
            try:
                # Tenta primeiro sem rotação
                texto_sem_rotacao = pytesseract.image_to_string(img, lang='por')
                texto_total += texto_sem_rotacao + "\n"
                
                # Se solicitado, tenta também com rotação de 180 graus
                if tentar_rotacao:
                    img_rotacionada = img.rotate(180, expand=True)
                    texto_rotacionado = pytesseract.image_to_string(img_rotacionada, lang='por')
                    texto_total += texto_rotacionado + "\n"
            except pytesseract.TesseractError as e:
                print(f"Erro no Tesseract ao processar imagem: {e}")
        return texto_total
    except Exception as e:
        print(f"Erro ao converter PDF para imagem: {e}")
        return ""

def extrair_anestesista(pdf_path):
    if not os.path.exists(pdf_path):
        return "Desconhecido"

    # Primeiro tenta extrair texto diretamente do PDF
    doc = fitz.open(pdf_path)
    texto = ""
    for page in doc:
        texto += page.get_text()
    doc.close()

    # Busca pelo anestesista no texto extraído
    anestesista = buscar_anestesista_no_texto(texto)
    texto_ocr = None
    
    # Se não encontrou e o texto está vazio, usa OCR
    if not anestesista and not texto.strip():
        texto_ocr = extrair_texto_ocr(pdf_path)
        anestesista = buscar_anestesista_no_texto(texto_ocr)
    
    # Se ainda não encontrou, mas havia texto (pode ser texto incorreto), tenta OCR mesmo assim
    if not anestesista and texto.strip():
        print(f"Tentando OCR para {os.path.basename(pdf_path)} (texto extraído não contém 'anestesista')")
        texto_ocr = extrair_texto_ocr(pdf_path)
        anestesista = buscar_anestesista_no_texto(texto_ocr)
    
    if not anestesista:
        print(f"Não foi possível encontrar anestesista em {os.path.basename(pdf_path)}")
        # Salva o texto extraído para debug se estiver em modo debug
        if DEBUG_MODE:
            nome_arquivo_debug = os.path.basename(pdf_path).replace('.pdf', '_debug.txt')
            caminho_debug = os.path.join("DEBUG_TEXTO", nome_arquivo_debug)
            with open(caminho_debug, 'w', encoding='utf-8') as f:
                f.write("=== TEXTO EXTRAÍDO DO PDF ===\n")
                f.write(texto if texto else "(vazio)")
                f.write("\n\n=== TEXTO DO OCR ===\n")
                if texto_ocr:
                    f.write(texto_ocr)
                elif texto.strip():
                    # Se não tinha texto_ocr mas havia texto, faz OCR agora para debug
                    texto_ocr_debug = extrair_texto_ocr(pdf_path)
                    f.write(texto_ocr_debug)
                else:
                    f.write("(não processado)")
            print(f"  Texto salvo em: {caminho_debug}")
        return "Desconhecido"
    
    return anestesista

def buscar_anestesista_no_texto(texto):
    """Busca o nome do anestesista no texto usando múltiplos padrões"""
    if not texto:
        return None
    
    # Padrões mais flexíveis para encontrar "anestesista"
    padroes = [
        r"anestesista\s*:?\s*(.+?)(?:\n|$)",  # anestesista: nome
        r"anestesista\s*:?\s*(.+?)(?:\r|$)",  # anestesista: nome (com \r)
        r"anestesista\s*:?\s*([A-Z][A-Z\s\.]+)",  # anestesista: NOME EM MAIÚSCULAS
        r"(?:dr|dra|dr\.|dra\.)\s+([A-Z][A-Z\s\.]+?)(?:\n|$)",  # DR/DRA NOME
    ]
    
    # Primeiro tenta buscar linha por linha (método original)
    for linha in texto.split("\n"):
        linha_limpa = linha.strip()
        if "anestesista" in linha_limpa.lower():
            # Tenta múltiplos padrões na linha
            for padrao in padroes:
                match = re.search(padrao, linha_limpa, re.IGNORECASE)
                if match:
                    nome = match.group(1).strip()
                    if nome and len(nome) > 3:  # Nome deve ter pelo menos 3 caracteres
                        return nome
    
    # Se não encontrou linha por linha, busca em todo o texto
    for padrao in padroes:
        matches = re.finditer(padrao, texto, re.IGNORECASE | re.MULTILINE)
        for match in matches:
            nome = match.group(1).strip()
            if nome and len(nome) > 3:
                return nome
    
    return None

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
