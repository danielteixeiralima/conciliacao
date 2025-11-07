# -*- coding: utf-8 -*-

###############################################################################
#                              conciliacao.py                                 #
###############################################################################

import ctypes
import pyautogui
import logging
import time
import os
import base64
from anthropic import Anthropic
from gera_txt import generate_txts_from_xls
import pyexcel_xls
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError, find_windows
from pywinauto.timings import TimeoutError
import cv2
from pywinauto import Desktop
from datetime import date, timedelta
import calendar
from holidays import Brazil
import openai
import difflib
import sys
from utils import login
from datetime import datetime, time as dt_time, timedelta
import shutil
from itertools import count
import re
pyautogui.FAILSAFE = False  # CUIDADO: não encosta nos cantos da tela para abortar
pyautogui.PAUSE = 0.1       # pequeno delay entre ações (deixa mais estável)
import psutil
import signal
from dotenv import load_dotenv
import os
import tempfile
# Força o carregamento do .env da mesma pasta do script,
# mesmo quando o .exe é chamado via web, cron ou manual.
base_dir = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(base_dir, ".env")

if os.path.exists(env_path):
    load_dotenv(dotenv_path=env_path)
    print(f"[DEBUG] .env carregado de: {env_path}")
else:
    print(f"[ERRO] .env não encontrado em: {env_path}")

br_holidays = Brazil()

# ============================================================
# 🔧 Configuração global de diretórios centralizados
# Logs e prints sempre vão para o diretório raiz do projeto
# ============================================================

# Diretório raiz fixo (independente do local do script)
# ============================================================
# 🔧 Diretórios fixos absolutos (garante funcionamento via .exe)
# ============================================================

# Define explicitamente o caminho raiz fixo no Windows
root_dir = r"C:\AUTOMACAO\conciliacao"

# Garante as pastas principais
log_dir = os.path.join(root_dir, "Logs")
prints_folder = os.path.join(root_dir, "prints")

os.makedirs(log_dir, exist_ok=True)
os.makedirs(prints_folder, exist_ok=True)

print(f"[DEBUG] Logs fixos em: {log_dir}")
print(f"[DEBUG] Prints fixos em: {prints_folder}")

_screenshot_counter = count(1)



# determina qual shopping (vai sobrescrever o arquivo se já existir)
# shopping = sys.argv[1] if len(sys.argv) > 1 else 'default'
# print(f"[DEBUG] shopping = {shopping}")

# substitui caracteres impróprios para nome de arquivo, se for o caso
# safe_shopping = shopping.replace(' ', '_').replace('á','a').replace('ó','o')  # etc.

# log_path = os.path.join(log_dir, f"{safe_shopping}.log")

portador_map = {
    "SDI": [
        {"codigo": "008", "banco": "341", "agencia": "2938", "conta": "18524-2", "rubrica": "FPP", "folder": r"\\192.168.18.4\hnc\COBRANCA_SHOPPING\RETORNO\341\SDI\FPP"},
        {"codigo": "010", "banco": "341", "agencia": "2938", "conta": "20468-8", "rubrica": "CONDOMINIO", "folder": r"\\192.168.18.4\hnc\cobranca_shopping\RETORNO\341\SDI\CONDOMINIO"},
        {"codigo": "018", "banco": "237", "agencia": "2373", "conta": "8892-7", "rubrica": "EMPREENDEDOR", "folder": r"\\192.168.18.4\HNC\COBRANCA_SHOPPING\RETORNO\237\SDI\EMPREENDEDOR"}
    ],
    "SMA": [
        {"codigo": "001", "banco": "033", "agencia": "116", "conta": "13002857-6", "rubrica": "EMPREENDEDOR", "folder": r"\\192.168.18.4\hnc\COBRANCA_SHOPPING\RETORNO\033\SMA\EMPREENDEDOR"},
        {"codigo": "002", "banco": "033", "agencia": "116", "conta": "13002848-0", "rubrica": "CONDOMINIO", "folder": r"\\192.168.18.4\hnc\COBRANCA_SHOPPING\RETORNO\033\SMA\CONDOMINIO"},
        {"codigo": "003", "banco": "033", "agencia": "116", "conta": "13004439-5", "rubrica": "FUNDO", "folder": r"\\192.168.18.4\hnc\COBRANCA_SHOPPING\RETORNO\033\SMA\FPP"}
    ],
    "SMO": [
        {"codigo": "001", "banco": "033", "agencia": "3907", "conta": "13004445-0", "rubrica": "EMPREENDEDOR", "folder": r"\\192.168.18.4\HNC\COBRANCA_SHOPPING\RETORNO\033\SMO\EMPREENDEDOR"},
        {"codigo": "003", "banco": "033", "agencia": "1160", "conta": "13002858-3", "rubrica": "CONDOMINIO", "folder": r"\\192.168.18.4\HNC\COBRANCA_SHOPPING\RETORNO\033\SMO\CONDOMINIO"},
        {"codigo": "005", "banco": "033", "agencia": "1160", "conta": "13002847-3", "rubrica": "FUNDO", "folder": r"\\192.168.18.4\HNC\COBRANCA_SHOPPING\RETORNO\033\SMO\FPP"}
    ],
    "SMS": [
        {"codigo": "002", "banco": "033", "agencia": "3907", "conta": "13003227-7", "rubrica": "EMPREENDEDOR", "folder": r"\\192.168.18.4\hnc\COBRANCA_SHOPPING\RETORNO\033\SMS\EMPREENDEDOR"},
        {"codigo": "008", "banco": "033", "agencia": "1160", "conta": "13002854-5", "rubrica": "CONDOMINIO", "folder": r"\\192.168.18.4\hnc\COBRANCA_SHOPPING\RETORNO\033\SMS\CONDOMINIO"},
        {"codigo": "009", "banco": "033", "agencia": "1160", "conta": "13002849-7", "rubrica": "FPP", "folder": r"\\192.168.18.4\hnc\COBRANCA_SHOPPING\RETORNO\033\SMS\FPP"}
    ],
    "SMT": [
        {"codigo": "002", "banco": "341", "agencia": "2938", "conta": "24491-6", "rubrica": "EMPREENDEDOR", "folder": r"\\192.168.18.4\hnc\COBRANCA_SHOPPING\RETORNO\341\SMT\EMPREENDEDOR\61"},
        {"codigo": "006", "banco": "341", "agencia": "2938", "conta": "41069-9", "rubrica": "CONDOMINIO", "folder": r"\\192.168.18.4\hnc\COBRANCA_SHOPPING\RETORNO\341\SMT\CONDOMINIO\133"},
        {"codigo": "007", "banco": "341", "agencia": "2938", "conta": "43346-9", "rubrica": "FPP", "folder": r"\\192.168.18.4\hnc\COBRANCA_SHOPPING\RETORNO\341\SMT\FUNDONOVO"}
    ],
    "SRP": [
        {"codigo": "001", "banco": "004", "agencia": "56", "conta": "29043-5", "rubrica": "EMPREENDEDOR", "folder": r"\\192.168.18.4\hnc\COBRANCA_SHOPPING\RETORNO\004\SRP\EMPREENDEDOR"},
        {"codigo": "008", "banco": "341", "agencia": "2938", "conta": "56292-9", "rubrica": "FPP", "folder": r"\\192.168.18.4\hnc\COBRANCA_SHOPPING\RETORNO\341\SRP\FPP"},
        {"codigo": "010", "banco": "341", "agencia": "2938", "conta": "50450-9", "rubrica": "CONDOMINIO", "folder": r"\\192.168.18.4\HNC\COBRANCA_SHOPPING\RETORNO\341\SRP\CONDOMINIO"}
    ]
}

def determine_variant(shopping):
    """
    Determina o variant com base no nome do shopping.
    """
    mapping = {
        "Shopping da Ilha": "SDI",
        "Shopping Mestre Álvaro": "SMA",
        "Shopping Moxuara": "SMO",
        "Shopping Montserrat": "SMS",
        "Shopping Metrópole": "SMT",
        "Shopping Rio Poty": "SRP",
        "Shopping Praia da Costa": "SPC"
    }
    return mapping.get(shopping, "SDI")



# # configura o logging para usar esse arquivo e sobrescrevê-lo a cada execução
# logging.basicConfig(
#     filename=log_path,
#     filemode='w',
#     level=logging.INFO,
#     format='%(asctime)s %(levelname)s:%(message)s'
# )
 


for w in Desktop(backend="uia").windows():
    logging.info(w.window_text())

openai.api_key = os.getenv("OPENAI_API_KEY")
anthropic = Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))


shopping_fases_tipo2 = {
    "Shopping Mestre Álvaro": {
        "Antecipados": [24, 25, 5, 5, 24, 24, 7, 7],
        "Atípicos": [31, 32, 6, 4, 31, 31],
        "Postecipados": [11, 2, 8, 8, 11, 11, 41]
    },
    "Shopping Montserrat": {
        "Antecipados": [24, 25, 5, 5, 24, 24, 7, 7],
        "Atípicos": [31, 32, 6, 15, 31, 31],
        "Postecipados": [11, 4, 8, 8, 11, 11, 22]
    },
    "Shopping Metrópole": {
        "Antecipados": [12, 13, 7, 7, 12, 12, 18, 18],
        "Atípicos": [31, 32, 6, 11, 31, 31],
        "Postecipados": [8, 9, 2, 2, 8, 8, 36]
    },
    "Shopping Rio Poty": {
        "Antecipados": [12, 13, 7, 7, 12, 12, 18, 18],
        "Atípicos": [31, 32, 6, 11, 31, 31],
        "Postecipados": [8, 9, 2, 2, 8, 8, 23]
    },
    "Shopping da Ilha": {
        "Antecipados": [12, 13, 7, 7, 12, 12, 18, 18],
        "Atípicos": [31, 32, 6, 11, 31, 31],
        "Postecipados": [8, 9, 2, 2, 8, 8, 36, 37]
    },
    "Shopping Moxuara": {
        "Antecipados": [12, 13, 24, 24, 12, 12, 18, 18],
        "Atípicos": [31, 32, 6, 11, 31, 31],
        "Postecipados": [8, 9, 2, 2, 8, 8, 39]
    }
}



# Mapeamento dos portadores conforme a tabela:


def select_cnab_files(folder):
    import re
    from datetime import datetime, time as dt_time

    arquivos = []
    agora = datetime.now()
    corte = dt_time(6, 0)

    for f in os.listdir(folder):
        if not f.lower().endswith('.ret') or '_033_' not in f:
            continue
        m = re.search(r'_(\d{2})(\d{2})(\d{4})(\d{2})\D*\.ret$', f, re.IGNORECASE)
        if not m:
            continue
        dia, mes, ano, hora = m.groups()
        dt_arch = datetime.strptime(f"{dia}{mes}{ano}{hora}", "%d%m%Y%H")
        if dt_arch.date() < agora.date() or dt_arch.time() < corte:
            arquivos.append(f)

    arquivos.sort()
    return arquivos

# A cada screenshot, um nome único será gerado para acumular os prints
SCREENSHOT_PATH = os.path.join(prints_folder, "monitor_screenshot.png")
if not os.path.exists(prints_folder):
    os.makedirs(prints_folder)

IS_SEGURO = False

# 🕒 Cria pasta de prints com data/hora em formato brasileiro
RUN_ID = datetime.now().strftime("%d-%m-%Y_%H-%M")

RUN_PRINTS_DIR = os.path.join(r"C:\AUTOMACAO\conciliacao\prints", RUN_ID)

try:
    os.makedirs(RUN_PRINTS_DIR, exist_ok=True)
    print(f"[DEBUG] Pasta de prints criada com sucesso em: {RUN_PRINTS_DIR}")
except Exception as e:
    print(f"[ERRO] Falha ao criar pasta de prints: {e}")
    RUN_PRINTS_DIR = r"C:\AUTOMACAO\conciliacao\prints_fallback"
    os.makedirs(RUN_PRINTS_DIR, exist_ok=True)
    print(f"[DEBUG] Usando pasta alternativa: {RUN_PRINTS_DIR}")

globals()["RUN_PRINTS_DIR"] = RUN_PRINTS_DIR



def fuzzy_contains(text, sub, threshold=0.8):
    """
    Verifica se 'sub' está contido em 'text' com tolerância para pequenas variações.
    Se a correspondência exata não for encontrada, usa uma janela deslizante para comparar
    a similaridade. Retorna True se a similaridade máxima for maior ou igual ao limiar.
    """
    text = text.lower()
    sub = sub.lower()
    if sub in text:
        return True
    max_ratio = 0
    sub_len = len(sub)
    for i in range(len(text) - sub_len + 1):
        segment = text[i:i+sub_len]
        ratio = difflib.SequenceMatcher(None, segment, sub).ratio()
        if ratio > max_ratio:
            max_ratio = ratio
        if max_ratio >= threshold:
            return True
    return False

def build_fase_map(shopping):
    base_y = 33
    step_y = 14
    coords = {}
    missing = missing_phases_map.get(shopping, [])
    real_index = 1
    for fase in range(1, 46):
        if fase in missing:
            continue
        if real_index < 14:
            coords[fase] = base_y + (real_index - 1) * step_y
        else:
            coords[fase] = 215
        real_index += 1
    return coords

missing_phases_map = {
    "Shopping Montserrat": [29, 39, 40, 41, 42, 43, 44, 45],
    "Shopping da Ilha": [3],
    "Shopping Mestre Álvaro": [12, 13, 38, 43, 44, 46, 47, 48, 49],
    "Shopping Metrópole": [3],
    "Shopping Moxuara": [],
    "Shopping Praia da Costa": [27, 42],
    "Shopping Rio Poty": [3, 39, 43, 44, 45, 46, 47, 48, 49]
}

def find_and_click_button_with_retry(image_path, max_attempts=10, confidence_range=(0.95, 0.6)):
    try:
        for attempt in range(max_attempts):
            confidence = confidence_range[0] - (attempt * (confidence_range[0] - confidence_range[1]) / max_attempts)
            logging.info(f"Tentativa {attempt + 1}/10 com confiança {confidence:.2f}")
            try:
                button = pyautogui.locateOnScreen(image_path, confidence=confidence)
                if button:
                    x, y = pyautogui.center(button)
                    logging.info(f"Botão encontrado nas coordenadas: x={x}, y={y}")
                    pyautogui.click(x, y)
                    return True
            except Exception as e:
                logging.error(f"Erro na tentativa {attempt + 1}: {str(e)}")
            time.sleep(2)
        logging.error("Botão não encontrado após todas as tentativas")
        return False
    except Exception as e:
        logging.error(f"Erro ao tentar localizar botão: {e}")
        return False

def verify_image_visibility(image_path, confidence=0.7, max_retries=10):
    try:
        logging.info(f"Verificando visibilidade da imagem: {image_path} com confiança {confidence}")
        for attempt in range(max_retries):
            button_location = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if button_location is not None:
                x, y = pyautogui.center(button_location)
                logging.info(f"Imagem visível nas coordenadas: x={x}, y={y}")
                return button_location
            else:
                logging.error(f"Tentativa {attempt + 1}/10: Imagem não encontrada na tela.")
                time.sleep(2)
        return None
    except Exception as e:
        logging.error(f"Erro ao verificar visibilidade da imagem: {e}")
        return None

def find_and_click_button(image_path, confidence=0.95):
    try:
        while True:
            button_location = verify_image_visibility(image_path, confidence=confidence)
            if button_location is not None:
                x, y = pyautogui.center(button_location)
                logging.info(f"Botão encontrado nas coordenadas: x={x}, y={y}")
                pyautogui.click(x, y)
                break
            else:
                logging.info("Tentando localizar a imagem novamente...")
                time.sleep(2)
    except Exception as e:
        logging.error(f"Erro ao localizar o botão: {e}")

def setup_logging_for_shopping(variant):
    """
    Configura o logging para o shopping informado.
    Sempre sobrescreve o arquivo antigo (modo 'w').
    """
    os.makedirs(log_dir, exist_ok=True)
    logfile = os.path.join(log_dir, f"{variant}.txt")

    # Remove o arquivo anterior para garantir log limpo
    if os.path.exists(logfile):
        try:
            os.remove(logfile)
            print(f"[DEBUG] Log antigo removido: {logfile}")
        except Exception as e:
            print(f"[ERRO] Não foi possível remover {logfile}: {e}")

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # Remove handlers antigos
    for h in list(logger.handlers):
        logger.removeHandler(h)

    # Cria novo handler sobrescrevendo o arquivo
    fh = logging.FileHandler(logfile, mode='w', encoding='utf-8')
    fh.setFormatter(logging.Formatter('%(asctime)s %(levelname)s:%(message)s'))
    logger.addHandler(fh)

    print(f"[DEBUG] Novo log ativo em: {logfile}")



def delete_all_prints():
    """
    Deleta todos os arquivos na pasta de prints.
    """
    for filename in os.listdir(prints_folder):
        file_path = os.path.join(prints_folder, filename)
        try:
            os.remove(file_path)
            logging.info(f"Print deletado: {file_path}")
        except Exception as e:
            logging.error(f"Erro ao deletar {file_path}: {e}")

def get_visible_index(shopping, fase_val):
    """
    Retorna quantas fases visíveis existem de 1 até fase_val inclusive,
    desconsiderando as fases ausentes conforme missing_phases_map.
    """
    missing = missing_phases_map.get(shopping, [])
    index = 0
    for f in range(1, fase_val + 1):
        if f not in missing:
            index += 1
    return index

def click_fase_tipo1(shopping, fase):
    """
    Corrige o método de clicar na fase, respeitando as fases ausentes
    também para fases acima de 13.
    """
    if not fase:
        logging.error(f"Fase inválida para {shopping}. Verifique o mapeamento.")
        return
    vi = get_visible_index(shopping, fase)
    if vi <= 14:
        y = 33 + 14 * (vi - 1)
        pyautogui.moveRel(-100, y)
        pyautogui.click()
    else:
        times_to_scroll = vi - 14
        pyautogui.moveRel(2, 215)
        for _ in range(times_to_scroll):
            pyautogui.click()
            time.sleep(0.3)
        pyautogui.moveRel(-100, 0)
        pyautogui.click()




def capture_screenshot(prefix="monitor"):
    global RUN_PRINTS_DIR
    """
    Captura a tela e salva com nome único dentro do diretório da execução.
    Ex.: prints/20250811-163300/monitor_0001.png
    """
    idx = next(_screenshot_counter)
    # Garante que a pasta de execução ainda existe (caso o .exe reinicie ou o antivírus limpe)
    if not os.path.exists(RUN_PRINTS_DIR):
        try:
            os.makedirs(RUN_PRINTS_DIR, exist_ok=True)
            logging.info(f"[RECOVERY] Pasta de prints recriada dinamicamente: {RUN_PRINTS_DIR}")
        except Exception as e:
            import tempfile
            fallback = os.path.join(tempfile.gettempdir(), f"conc_prints_recovery_{datetime.now().strftime('%H%M%S')}")
            os.makedirs(fallback, exist_ok=True)
            logging.warning(f"[RECOVERY] Falha ao recriar pasta original, salvando prints em: {fallback} ({e})")
            RUN_PRINTS_DIR = fallback


    filename = f"{prefix}_{idx:04d}.png"
    path = os.path.join(RUN_PRINTS_DIR, filename)
    logging.info(f"Capturando screenshot em {path}")
    pyautogui.screenshot(path)
    return path

def analyze_screenshot(image_path):
    """ Converte a imagem para Base64 e envia para a API do GPT-4o, retornando o texto extraído. """
    if not os.path.exists(image_path):
        logging.error("Nenhuma imagem válida encontrada para análise.")
        return None

    with open(image_path, "rb") as img_file:
        base64_image = base64.b64encode(img_file.read()).decode("utf-8")

    while True:  # Loop para garantir que obtenha uma resposta válida
        try:
            response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": "Identifique a tela em foco no print, normalmente um modal ou um alerta, e me devolva o conteúdo COMPLETO da tela por escrito, normalmente esse modal, se existir, vai ter título como 'Gerando Área de Recibo' ou 'Alerta VSSC', lembrando que a análise deve ser feita principalmente se há um modal ou tela no meu aplicativo aberta além da tela principal. Se não houver nenhum modal ou tela, pode-se concluir que o sistema está no lobby. Muita atenção ao conteúdo de cada modal se houver mais de um modal detectado. EU PRECISO QUE O TEXTO DA RESPOSTA SEMPRE COMECE COM 'MODAL DETECTADO' OU 'MODAL NÃO DETECTADO'",
                            },
                            {
                                "type": "image_url",
                                "image_url": {"url": f"data:image/png;base64,{base64_image}"},
                            },
                        ],
                    }
                ],
                request_timeout=60
            )
            extracted_text = response["choices"][0]["message"]["content"].strip()
            if extracted_text:  # Se recebeu uma resposta válida, sai do loop
                return extracted_text

        except Exception as e:
            logging.error(f"Erro ao analisar imagem (tentando novamente): {e}")
            time.sleep(5)  # Aguarda antes de tentar novamente

def find_and_click_button_with_retry(image_path, max_attempts=10, confidence_range=(0.95, 0.6)):
    try:
        for attempt in range(max_attempts):
            confidence = confidence_range[0] - (attempt * (confidence_range[0] - confidence_range[1]) / max_attempts)
            logging.info("Tentativa %d/%d com confiança %.2f" % (attempt + 1, max_attempts, confidence))
            try:
                button = pyautogui.locateOnScreen(image_path, confidence=confidence)
                if button:
                    x, y = pyautogui.center(button)
                    logging.info("Botão encontrado nas coordenadas: x=%d, y=%d" % (x, y))
                    pyautogui.click(x, y)
                    return True
            except Exception as e:
                logging.error("Erro na tentativa %d: %s" % (attempt + 1, str(e)))
            time.sleep(2)
        logging.error("Botão não encontrado após todas as tentativas")
        return False
    except Exception as e:
        logging.error("Erro ao tentar localizar botão: %s" % str(e))
        return False

def verify_image_visibility(image_path, confidence=0.7, max_retries=10):
    try:
        logging.info("Verificando visibilidade da imagem: %s com confiança %f" % (image_path, confidence))
        for attempt in range(max_retries):
            button_location = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if button_location is not None:
                x, y = pyautogui.center(button_location)
                logging.info("Imagem visível nas coordenadas: x=%d, y=%d" % (x, y))
                return button_location
            else:
                logging.error("Tentativa %d/%d: Imagem não encontrada na tela." % (attempt + 1, max_retries))
                time.sleep(2)
        return None
    except Exception as e:
        logging.error("Erro ao verificar visibilidade da imagem: %s" % str(e))
        return None

def find_and_click_button(image_path, confidence=0.95):
    try:
        while True:
            button_location = verify_image_visibility(image_path, confidence=confidence)
            if button_location is not None:
                x, y = pyautogui.center(button_location)
                logging.info("Botão encontrado nas coordenadas: x=%d, y=%d" % (x, y))
                pyautogui.click(x, y)
                break
            else:
                logging.info("Tentando localizar a imagem novamente...")
                time.sleep(2)
    except Exception as e:
        logging.error("Erro ao localizar o botão: %s" % str(e))
def select_cnab_files(folder):
    import re
    from datetime import datetime, time as dt_time

    arquivos = []
    agora = datetime.now()
    corte = dt_time(6, 0)

    for f in os.listdir(folder):
        if not f.lower().endswith('.ret') or '_033_' not in f:
            continue
        m = re.search(r'_(\d{2})(\d{2})(\d{4})(\d{2})\D*\.ret$', f, re.IGNORECASE)
        if not m:
            continue
        dia, mes, ano, hora = m.groups()
        dt_arch = datetime.strptime(f"{dia}{mes}{ano}{hora}", "%d%m%Y%H")
        if dt_arch.date() < agora.date() or dt_arch.time() < corte:
            arquivos.append(f)

    arquivos.sort()
    return arquivos

folder_map = {
    "Shopping Montserrat": r"C:\Program Files\Victor & Schellenberger\VSSC_MONTSERRAT",
    "Shopping da Ilha": r"C:\Program Files\Victor & Schellenberger\VSSC_ILHA",
    "Shopping Rio Poty": r"C:\Program Files\Victor & Schellenberger\VSSC_TERESINA",
    "Shopping Metrópole": r"C:\Program Files\Victor & Schellenberger\VSSC_METROPOLE",
    "Shopping Moxuara": r"C:\Program Files\Victor & Schellenberger\VSSC_MOXUARA",
    "Shopping Praia da Costa": r"C:\Program Files\Victor & Schellenberger\VSSC_PRAIADACOSTA",
    "Shopping Mestre Álvaro": r"C:\Program Files\Victor & Schellenberger\VSSC_MESTREALVARO"
}

def kill_vsloader():
    """Força o fechamento do VSLoader.exe."""
    for proc in psutil.process_iter(['pid', 'name']):
        if 'VSLOADER' in proc.info['name'].upper():
            try:
                proc.kill()
                logging.warning(f"💀 VSLOADER.EXE finalizado forçadamente (PID={proc.info['pid']})")
            except Exception as e:
                logging.error(f"Erro ao tentar encerrar VSLOADER.EXE: {e}")
def execute_vsloader(shopping):
    print(f"[DEBUG] execute_vsloader recebeu shopping = {shopping}")

    # 🔧 Garante que o cwd é o mesmo da pasta do script (corrige execução via .exe / servidor)
    os.chdir(base_dir)
    print(r"[DEBUG] Diretório de trabalho forçado para: C:\AUTOMACAO\conciliacao")


    variant = determine_variant(shopping)
    setup_logging_for_shopping(variant)
    logging.info(f"Iniciando conciliação para shopping '{shopping}' (variant={variant})")

    # 🔍 Loga as janelas ativas para depuração
    logging.info("Listando janelas ativas (Desktop UIA):")
    try:
        for w in Desktop(backend="uia").windows():
            logging.info(w.window_text())
    except Exception as e:
        logging.info(f"Falha ao listar janelas: {e}")


    if variant in ("SMA", "SMO", "SMS"):

        # usa UTC-3 para "hoje" e o corte de 06h
        now = datetime.utcnow() - timedelta(hours=3)
        today = now.date()
        six_am = dt_time(6, 0)

        def previous_business_day(d):
            prev = d - timedelta(days=1)
            # pula fins de semana e feriados (nacionais + ES, conforme br_holidays)
            while prev.weekday() >= 5 or prev in br_holidays:
                prev -= timedelta(days=1)
            return prev

        # limpeza prévia das pastas de destino
        for sub in ("ret_emp", "ret_con", "ret_fpp"):
            shutil.rmtree(os.path.join(r"C:\AUTOMACAO\conciliacao", sub), ignore_errors=True)

        # calcula datas-alvo conforme as regras
        if today.weekday() == 0:  # segunda-feira
            saturday = today - timedelta(days=2)  # sábado calendário
            friday_bd = previous_business_day(today - timedelta(days=1))  # “sexta útil”

            def want_file(file_date, file_mtime):
                t = file_mtime.time()
                if file_date == today:
                    return t < six_am                       # hoje: só até 06h
                if (saturday not in br_holidays) and (file_date == saturday):
                    return True                             # sábado inteiro (se não for feriado)
                if file_date == friday_bd:
                    return t >= six_am                      # “sexta útil” após 06h
                return False
        else:
            prev_bd = previous_business_day(today)

            def want_file(file_date, file_mtime):
                t = file_mtime.time()
                if file_date == today:
                    return t < six_am                       # hoje: só até 06h
                if file_date == prev_bd:
                    return t >= six_am                      # dia útil anterior: após 06h
                return False

        def dst_dir_for_rubrica(rub):
            r = rub.upper()
            if "EMP" in r:            # EMPREENDEDOR
                return "ret_emp"
            if "CONDOM" in r:         # CONDOMINIO
                return "ret_con"
            # FPP / FUNDO / FUNDONOVO vão para ret_fpp
            return "ret_fpp"

        # varre cada portador e copia somente os .RET desejados
        # === LOG DE DEPURAÇÃO DA CÓPIA DOS ARQUIVOS .RET ===
        logging.info(f"[RET_FETCH] Iniciando busca de arquivos .RET para {variant}")

        for port in portador_map[variant]:
            src = port["folder"]
            dst = os.path.join(r"C:\AUTOMACAO\conciliacao", dst_dir_for_rubrica(port["rubrica"]))
            os.makedirs(dst, exist_ok=True)
            logging.info(f"[RET_FETCH] Verificando portador {port['codigo']} ({port['rubrica']})")
            logging.info(f"[RET_FETCH] Pasta origem: {src}")
            logging.info(f"[RET_FETCH] Pasta destino: {dst}")

            if not os.path.exists(src):
                logging.error(f"[RET_FETCH][ERRO] Pasta de origem não encontrada: {src}")
                continue

            try:
                arquivos = [f for f in os.listdir(src) if f.lower().endswith(".ret")]
            except PermissionError as e:
                logging.error(f"[RET_FETCH][ERRO] Sem permissão para acessar {src}: {e}")
                continue
            except Exception as e:
                logging.error(f"[RET_FETCH][ERRO] Falha ao listar {src}: {e}")
                continue

            logging.info(f"[RET_FETCH] {len(arquivos)} arquivo(s) encontrados em {src}")

            for fn in arquivos:
                m = re.search(r'_(\d{2})(\d{2})(\d{4})\d{2}', fn)
                if not m:
                    logging.warning(f"[RET_FETCH][SKIP] Nome fora do padrão: {fn}")
                    continue

                dd, mm, yyyy = m.groups()
                file_date = datetime.strptime(f"{dd}{mm}{yyyy}", "%d%m%Y").date()
                fullpath = os.path.join(src, fn)

                try:
                    file_mtime = datetime.fromtimestamp(os.path.getmtime(fullpath))
                except FileNotFoundError:
                    logging.warning(f"[RET_FETCH][SKIP] Arquivo sumiu: {fullpath}")
                    continue

                if want_file(file_date, file_mtime):
                    try:
                        shutil.copy2(fullpath, os.path.join(dst, fn))
                        logging.info(f"[RET_FETCH][COPIADO] {fn} -> {dst}")
                    except Exception as e:
                        logging.error(f"[RET_FETCH][ERRO] Falha ao copiar {fn}: {e}")
                else:
                    logging.debug(f"[RET_FETCH][IGNORADO] {fn} (data={file_date}, hora={file_mtime.time()}) não passou no filtro")







    
    try:
        
        user32 = ctypes.windll.user32

        def get_foreground_window():
            return user32.GetForegroundWindow()

        def wait_for_stable_focus(prev_handle, max_wait=15):
            start_time = time.time()
            while True:
                if time.time() - start_time > max_wait:
                    logging.info("Tempo máximo de espera por foco estável atingido.")
                    break
                time.sleep(1)
                current_handle = get_foreground_window()
                if current_handle != prev_handle:
                    time.sleep(1.5)
                    second_check = get_foreground_window()
                    if second_check != current_handle:
                        prev_handle = second_check
                        continue
                    break

        def wait_for_focus_change(prev_handle, max_wait=60):
            start_time = time.time()
            while True:
                if time.time() - start_time > max_wait:
                    logging.info("Tempo máximo de espera por mudança de foco atingido.")
                    break
                time.sleep(1)
                current_handle = get_foreground_window()
                if current_handle != prev_handle:
                    break

        

        folder = folder_map.get(
            shopping,
            r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_MESTREALVARO_HOM"
        )
        # chama o login do utils passando a pasta
        screen_width, screen_height = pyautogui.size()
        center_x = screen_width // 2
        center_y = screen_height // 2


        time.sleep(8)
        pyautogui.press('win')
        time.sleep(4)
        pyautogui.write('file explorer')
        pyautogui.press('enter')
        time.sleep(14)

        pyautogui.hotkey('alt', 'd')
        pyautogui.typewrite(folder)
        pyautogui.press('enter')
        time.sleep(3)

        pyautogui.typewrite("VSLOADER.EXE")
        pyautogui.press('enter')
        time.sleep(10)

        logging.info("VSLOADER.EXE executado.")

        pyautogui.typewrite("z8")
        pyautogui.typewrite("S@cavalcante")
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(5)
        pyautogui.press('enter')
        time.sleep(20)

        pyautogui.hotkey('win', 'up')
        time.sleep(2)
        pyautogui.hotkey('win', 'down')
        time.sleep(2)
        pyautogui.hotkey('win', 'up')
        time.sleep(5)
        repeated_text = None
        repeat_count = 0

        while True:
            screenshot1 = capture_screenshot()
            extracted1 = analyze_screenshot(screenshot1)
            time.sleep(2)
            screenshot2 = capture_screenshot()
            extracted2 = analyze_screenshot(screenshot2)

            if not extracted1 or not extracted2:
                time.sleep(8)
                continue

            extracted1_l = extracted1.lower() if extracted1 else ""
            extracted2_l = extracted2.lower() if extracted2 else ""

            if (
                (
                    ("não há nenhum modal" in extracted1_l)
                    or ("lobby" in extracted1_l)
                    or ("modal não detectado" in extracted1_l)
                )
                and (
                    ("não há nenhum modal" in extracted2_l)
                    or ("lobby" in extracted2_l)
                    or ("modal não detectado" in extracted2_l)
                )
            ):
                logging.info(
                    "execute_vsloader: Lobby identificado (nenhum modal detectado) em ambos os prints. "
                    "Texto 1: %s Texto 2: %s", extracted1, extracted2
                )
                break

            combined = (extracted1 + " " + extracted2) if extracted1 and extracted2 else (extracted1 or extracted2)
            combined = combined.strip().lower()

            if combined == repeated_text:
                repeat_count += 1
            else:
                repeated_text = combined
                repeat_count = 0

            if repeat_count >= 20:
                logging.error(f"⚠️ Travamento detectado no shopping {shopping}. Mesmo modal repetido 20x seguidas: {combined[:100]}...")
                kill_vsloader()
                raise RuntimeError(f"Travamento detectado no shopping {shopping}")

            if fuzzy_contains(combined, "<ESC>"):
                pyautogui.click(center_x, center_y)
                pyautogui.press('esc')
            elif fuzzy_contains(combined, "competência de trabalho") and fuzzy_contains(combined, "alerta vssc"):
                pyautogui.click(center_x, center_y)
                pyautogui.press('enter')
            elif fuzzy_contains(combined, "alerta vssc"):
                pyautogui.click(center_x, center_y)
                pyautogui.press('esc')
            elif fuzzy_contains(combined, "contratos com término") or fuzzy_contains(combined, "contratos com termino"):
                logging.info("Contrato com término detectado")
                for _ in range(5):
                    pyautogui.press('esc')
                time.sleep(2)
                break
            else:
                logging.info("Nenhuma tela detectada")
                logging.info(combined)

        if variant not in ("SMA", "SMO", "SMS"):
            print(f"Shopping de fora do estado")
            time.sleep(15)

            pyautogui.hotkey('alt', 's')
            time.sleep(0.5)
            
            pyautogui.press('down')
            time.sleep(0.3)
        
            pyautogui.press('enter')
            time.sleep(8)

            # Início do laço para cada portador do shopping conforme a tabela (iteração de 3 vezes para este shopping)
            for portador in portador_map.get(determine_variant(shopping), []):
                pyautogui.hotkey('alt', 's')
                time.sleep(0.5)
                for _ in range(1):
                    pyautogui.press('right')
                    time.sleep(0.3)
                for _ in range(2):
                    pyautogui.press('down')
                    time.sleep(0.3)
                pyautogui.press('right')
                time.sleep(0.3)
                for _ in range(3):
                    pyautogui.press('down')
                    time.sleep(0.3)
                pyautogui.press('enter')
                time.sleep(3)
                pyautogui.hotkey('alt', 'space')
                time.sleep(0.3)
                pyautogui.press('down')
                time.sleep(0.3)
                pyautogui.press('enter')
                pyautogui.moveRel(110, 50)
                # pyautogui.moveRel(110, 45)
                pyautogui.click()
                pyautogui.click()
                time.sleep(1)
                # Busca via ctrl+f para o portador atual
                pyautogui.hotkey('ctrl', 'f')
                time.sleep(0.5)
                pyautogui.typewrite(portador["codigo"])
                logging.info(portador["codigo"])
                pyautogui.press('enter')
                time.sleep(3)
                pyautogui.press('down')
                time.sleep(0.5)
                pyautogui.press('up')

                repeated_text = None
                repeat_count = 0

                while True:
                    screenshot1 = capture_screenshot()
                    extracted1 = analyze_screenshot(screenshot1)
                    time.sleep(2)
                    screenshot2 = capture_screenshot()
                    extracted2 = analyze_screenshot(screenshot2)
                    if extracted1 is None and extracted2 is None:
                        logging.info("execute_vsloader: Nenhum texto extraído nos prints. extracted1=None, extracted2=None.")
                        time.sleep(3)
                        continue
                    extracted1_l = extracted1.lower() if extracted1 else ""
                    extracted2_l = extracted2.lower() if extracted2 else ""

                    if (
                        (
                            ("não há nenhum modal" in extracted1_l)
                            or ("lobby" in extracted1_l)
                            or ("modal não detectado" in extracted1_l)
                        )
                        and (
                            ("não há nenhum modal" in extracted2_l)
                            or ("lobby" in extracted2_l)
                            or ("modal não detectado" in extracted2_l)
                        )
                    ):
                        logging.info(
                            "execute_vsloader: Lobby identificado (nenhum modal detectado) em ambos os prints. "
                            "Texto 1: %s Texto 2: %s", extracted1, extracted2
                        )
                        break

                    combined_extracted = (extracted1 + " " + extracted2) if extracted1 and extracted2 else (extracted1 or extracted2)
                    combined_extracted = combined_extracted.strip().lower()

                    if combined_extracted == repeated_text:
                        repeat_count += 1
                    else:
                        repeated_text = combined_extracted
                        repeat_count = 0

                    if repeat_count >= 20:
                        logging.error(f"⚠️ Travamento detectado no shopping {shopping}. Mesmo modal repetido 20x seguidas: {combined_extracted[:100]}...")
                        kill_vsloader()
                        raise RuntimeError(f"Travamento detectado no shopping {shopping}")

                    screen_width = pyautogui.size().width
                    screen_height = pyautogui.size().height
                    center_x = screen_width // 2
                    center_y = screen_height // 2

                    if fuzzy_contains(combined_extracted, "alerta vssc"):
                        logging.info("execute_vsloader: Print indica ação ENTER. Texto identificado: %s", combined_extracted)
                        pyautogui.moveTo(center_x, center_y)
                        pyautogui.click()
                        pyautogui.press("enter")
                        time.sleep(1)
                        break
                    elif fuzzy_contains(combined_extracted, "leitura do arquivo cnab"):
                        logging.info("execute_vsloader: Tela 'Leitura do Arquivo CNAB' identificada. Texto identificado: %s", combined_extracted)
                        break
                    else:
                        logging.info("execute_vsloader: Nenhuma condição modal identificada. Texto identificado: %s", combined_extracted)
                    time.sleep(3)

                time.sleep(2)
                
                # Seleção de arquivo: utiliza a pasta do portador atual conforme a tabela
                folder_selecionado = portador["folder"]
                

                pyautogui.hotkey('alt', 'space')
                time.sleep(0.3)
                pyautogui.press('down')
                time.sleep(0.3)
                pyautogui.press('enter')
                pyautogui.moveRel(210,110)
                pyautogui.click()
                pyautogui.click()
                # Vai direto para a barra de endereço (atalho universal)
                                    # Vai direto para a barra de endereço
                # 🔧 Garante que o seletor de arquivo está ativo
                time.sleep(3)

                # Abre o menu do sistema da janela ("Restaurar, Mover, Tamanho..." etc.)
                pyautogui.hotkey('alt', 'space')
                time.sleep(0.3)
                pyautogui.press('down')
                time.sleep(0.3)
                pyautogui.press('enter')

                # Move o foco para o campo de endereço (coordenadas seguras)
                pyautogui.moveRel(0, 200)  # <== AJUSTE: move o mouse até a barra de endereço
                pyautogui.click()
                pyautogui.click()
                time.sleep(0.5)

                time.sleep(3)
                pyautogui.hotkey('alt', 'space')
                time.sleep(0.3)                # ==========================================================
                # 🧩 SELEÇÃO DE ARQUIVO — CORREÇÃO DE CAMINHO E DEBUG
                # ==========================================================
                folder_selecionado = portador["folder"]
                logging.info(f"[CNAB] Portador {portador['codigo']} ({portador['rubrica']}) - Pasta esperada: {folder_selecionado}")

                # 🧭 Verifica se a pasta existe antes de tentar usá-la
                if not os.path.exists(folder_selecionado):
                    logging.error(f"[ERRO] Pasta não existe: {folder_selecionado}")
                else:
                    arquivos_ret = [f for f in os.listdir(folder_selecionado) if f.lower().endswith('.ret')]
                    logging.info(f"[CNAB] {len(arquivos_ret)} arquivo(s) .RET disponível(is) em {folder_selecionado}:")

                
                # 💾 Captura um print do caminho digitado para debug
                path_debug = os.path.join(RUN_PRINTS_DIR, f"path_debug_{portador['codigo']}_{datetime.now().strftime('%H%M%S')}.png")
                pyautogui.screenshot(path_debug)
                logging.info(f"[DEBUG] Screenshot do caminho digitado salvo em: {path_debug}")

                # Confirma o caminho
                pyautogui.press('enter')
                time.sleep(2)

                # 🧾 Seleciona o primeiro arquivo .RET encontrado (ou o mais recente)
                try:
                    arquivos_ret = sorted(
                        [f for f in os.listdir(folder_selecionado) if f.lower().endswith(".ret")],
                        key=lambda f: os.path.getmtime(os.path.join(folder_selecionado, f)),
                        reverse=True
                    )
                    if arquivos_ret:
                        arquivo_escolhido = arquivos_ret[0]
                        fullpath = os.path.join(folder_selecionado, arquivo_escolhido)
                        logging.info(f"[CNAB] Abrindo arquivo: {fullpath}")
                        pyautogui.typewrite(fullpath)
                        time.sleep(1)
                        pyautogui.press('enter')
                        time.sleep(2)
                    else:
                        logging.warning(f"[CNAB] Nenhum arquivo .RET encontrado em {folder_selecionado}")
                except Exception as e:
                    logging.error(f"[ERRO] Falha ao listar arquivos em {folder_selecionado}: {e}")

                pyautogui.press('down')
                time.sleep(0.3)
                pyautogui.press('enter')
                pyautogui.moveRel(-200,60)
                # pyautogui.moveRel(-200,65)
                pyautogui.click()
                pyautogui.click()
                time.sleep(2)

        

    # ... dentro de execute_vsloader, no lugar removido:

                if portador["banco"] == "033":
                    arquivos = select_cnab_files(folder_selecionado)
                    for nome_ret in arquivos:
                        pyautogui.moveRel(250, 0)
                        pyautogui.click()
                        time.sleep(0.3)

                        # ✅ Captura o print antes de digitar o nome do arquivo
                        debug_print_path = os.path.join(RUN_PRINTS_DIR, f"before_typing_filename_{datetime.now().strftime('%H%M%S')}.png")
                        pyautogui.screenshot(debug_print_path)
                        logging.info(f"[DEBUG] Screenshot antes de digitar nome do arquivo salvo em: {debug_print_path}")

                        fullpath = os.path.join(folder_selecionado, nome_ret)
                        pyautogui.typewrite(fullpath)
                        pyautogui.press('enter')

                        time.sleep(1)
                        pyautogui.press('enter')
                        repeated_text = None
                        repeat_count = 0

                        while True:
                            screenshot1 = capture_screenshot()
                            extracted1 = analyze_screenshot(screenshot1)
                            time.sleep(2)
                            screenshot2 = capture_screenshot()
                            extracted2 = analyze_screenshot(screenshot2)

                            if extracted1 is None and extracted2 is None:
                                logging.info("execute_vsloader: Nenhum texto extraído nos prints (033). extracted1=None, extracted2=None.")
                                time.sleep(3)
                                continue

                            combined_tmp = (extracted1 or "") + " " + (extracted2 or "")
                            combined_tmp = combined_tmp.strip().lower()

                            # 🚨 DETECTOR DE TRAVAMENTO
                            if combined_tmp == repeated_text:
                                repeat_count += 1
                            else:
                                repeated_text = combined_tmp
                                repeat_count = 0

                            if repeat_count >= 20:
                                logging.error(f"⚠️ Travamento detectado no shopping {shopping} (033). Mesmo modal repetido 20x seguidas: {combined_tmp[:100]}...")
                                kill_vsloader()
                                raise RuntimeError(f"Travamento detectado no shopping {shopping}")

                            if fuzzy_contains(combined_tmp, "LEITURA CONCLUÍDA"):
                                logging.info("execute_vsloader: Leitura concluída (033). Texto identificado: %s", combined_tmp)
                                cx = pyautogui.size().width // 2
                                cy = pyautogui.size().height // 2
                                pyautogui.moveTo(cx, cy)
                                pyautogui.click()
                                break
                            elif fuzzy_contains(combined_extracted, "Alerta VSSC") and fuzzy_contains(combined_extracted, "Recria"):
                                logging.info("execute_vsloader: Alerta na baixa (Recria). Texto identificado: %s", combined_extracted)
                                pyautogui.moveTo(center_x, center_y)
                                pyautogui.click()
                                pyautogui.press("enter")
                                time.sleep(3)
                            else:
                                logging.info("execute_vsloader: Aguardando 'LEITURA CONCLUÍDA' (033). Texto identificado: %s", combined_tmp)
                            time.sleep(3)
                        for _ in range(4):
                            pyautogui.press('esc')
                            time.sleep(0.3)
                else:
                    pyautogui.hotkey('alt', 'space')
                    time.sleep(0.3)
                    pyautogui.press('down')
                    time.sleep(0.3)
                    pyautogui.press('enter')
                    pyautogui.moveRel(0,35)
                    pyautogui.click()
                    pyautogui.click()
                    time.sleep(0.3)
                    pyautogui.press('down')
                    time.sleep(0.3)
                    pyautogui.press('enter')

                    time.sleep(2)
                    pyautogui.press('enter')
                    # Validação via prints
                    repeated_text = None
                    repeat_count = 0

                    while True:
                        screenshot1 = capture_screenshot()
                        extracted1 = analyze_screenshot(screenshot1)
                        time.sleep(2)
                        screenshot2 = capture_screenshot()
                        extracted2 = analyze_screenshot(screenshot2)
                        if extracted1 is None and extracted2 is None:
                            logging.info("execute_vsloader: Nenhum texto extraído nos prints (não 033). extracted1=None, extracted2=None.")
                            time.sleep(3)
                            continue
                        extracted1_l = extracted1.lower() if extracted1 else ""
                        extracted2_l = extracted2.lower() if extracted2 else ""

                        if (
                            (
                                ("não há nenhum modal" in extracted1_l)
                                or ("lobby" in extracted1_l)
                                or ("modal não detectado" in extracted1_l)
                            )
                            and (
                                ("não há nenhum modal" in extracted2_l)
                                or ("lobby" in extracted2_l)
                                or ("modal não detectado" in extracted2_l)
                            )
                        ):
                            logging.info(
                                "execute_vsloader: Lobby identificado (nenhum modal detectado) em ambos os prints. "
                                "Texto 1: %s Texto 2: %s", extracted1, extracted2
                            )
                            break

                        combined_extracted = (extracted1 + " " + extracted2) if extracted1 and extracted2 else (extracted1 or extracted2)
                        combined_extracted = combined_extracted.strip().lower()

                        if combined_extracted == repeated_text:
                            repeat_count += 1
                        else:
                            repeated_text = combined_extracted
                            repeat_count = 0

                        if repeat_count >= 20:
                            logging.error(f"⚠️ Travamento detectado no shopping {shopping} (não 033). Mesmo modal repetido 20x seguidas: {combined_extracted[:100]}...")
                            kill_vsloader()
                            raise RuntimeError(f"Travamento detectado no shopping {shopping}")

                        screen_width = pyautogui.size().width
                        screen_height = pyautogui.size().height
                        center_x = screen_width // 2
                        center_y = screen_height // 2

                        if fuzzy_contains(combined_extracted, "alerta vssc") and fuzzy_contains(combined_extracted, "o arquivo de integração para este banco"):
                            logging.info("execute_vsloader: Alerta na integração (não 033). Texto identificado: %s", combined_extracted)
                            pyautogui.moveTo(center_x, center_y)
                            pyautogui.click()
                            pyautogui.press("enter")
                            time.sleep(3)
                        elif fuzzy_contains(combined_extracted, "atenção") and fuzzy_contains(combined_extracted, "competência de trabalho será alterada"):
                            logging.info("execute_vsloader: Atenção/Competência será alterada (não 033) -> ENTER. Texto identificado: %s", combined_extracted)
                            pyautogui.moveTo(center_x, center_y)
                            pyautogui.click()
                            pyautogui.press("enter")
                            time.sleep(3)
                            break
                        elif fuzzy_contains(combined_extracted, "leitura concluída") or fuzzy_contains(combined_extracted, "gravação concluída"):
                            logging.info("execute_vsloader: Leitura/Gravação concluída (não 033). Texto identificado: %s", combined_extracted)
                            pyautogui.moveTo(center_x, center_y)
                            pyautogui.click()
                            for _ in range(3):
                                pyautogui.press('esc')
                                time.sleep(0.3)
                            time.sleep(3)
                            break
                        elif fuzzy_contains(combined_extracted, "pressione <esc>") or fuzzy_contains(combined_extracted, "alerta vssc"):
                            logging.info("execute_vsloader: Solicita ESC/Alerta VSSC (não 033). Texto identificado: %s", combined_extracted)
                            pyautogui.moveTo(center_x, center_y)
                            pyautogui.click()
                            pyautogui.press("esc")
                            time.sleep(3)
                        elif fuzzy_contains(combined_extracted, "contratos com término") or fuzzy_contains(combined_extracted, "contratos com termino"):
                            logging.info("execute_vsloader: Contratos com término (não 033). Texto identificado: %s", combined_extracted)
                            break
                        elif fuzzy_contains(combined_extracted, "emissão do recibo") and fuzzy_contains(combined_extracted, "competência"):
                            logging.info("execute_vsloader: Emissão do Recibo/Competência (não 033). Texto identificado: %s", combined_extracted)
                            break
                        elif (
                            fuzzy_contains(combined_extracted, "competência de trabalho:")
                            and fuzzy_contains(combined_extracted, "período fechado")
                            and fuzzy_contains(combined_extracted, "(faturamento)")
                        ):
                            logging.info("execute_vsloader: Lobby identificado (Competência/Período Fechado/Faturamento). Texto identificado: %s", combined_extracted)
                            break
                        else:
                            logging.info("execute_vsloader: Nenhuma condição modal identificada (não 033). Texto identificado: %s", combined_extracted)
                        time.sleep(3)

                    # No final de cada iteração do laço, realizar 4 cliques em ESC
                    for _ in range(4):
                        pyautogui.press("esc")
                        time.sleep(0.3)

    ##################### baixa ###########################

                pyautogui.hotkey('alt', 's')
                time.sleep(0.5)
                for _ in range(1):
                    pyautogui.press('right')
                    time.sleep(0.3)
                for _ in range(2):
                    pyautogui.press('down')
                    time.sleep(0.3)
                pyautogui.press('right')
                time.sleep(0.3)
                for _ in range(5):
                    pyautogui.press('down')
                    time.sleep(0.3)
                pyautogui.press('enter')

                time.sleep(2)
                pyautogui.press('down')
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'f')
                time.sleep(0.5)
                pyautogui.typewrite(portador["codigo"])
                pyautogui.press('enter')
                time.sleep(1)

                pyautogui.press('enter')
                time.sleep(3)
                pyautogui.press('down')
                time.sleep(1)
                pyautogui.press('up')
                time.sleep(1)
                pyautogui.hotkey('alt', 'space')
                time.sleep(0.3)
                pyautogui.press('down')
                time.sleep(0.3)
                pyautogui.press('enter')
                pyautogui.moveRel(30, 325)
                pyautogui.click()
                pyautogui.click()
                prev_handle = get_foreground_window()
                wait_for_focus_change(prev_handle)
                wait_for_stable_focus(prev_handle)
                time.sleep(2)
                repeated_text = None
                repeat_count = 0

                while True:
                    screenshot1 = capture_screenshot()
                    extracted1 = analyze_screenshot(screenshot1)
                    time.sleep(2)
                    screenshot2 = capture_screenshot()
                    extracted2 = analyze_screenshot(screenshot2)
                    if extracted1 is None and extracted2 is None:
                        logging.info("execute_vsloader: Nenhum texto extraído nos prints (baixa). extracted1=None, extracted2=None.")
                        time.sleep(3)
                        continue
                    extracted1_l = extracted1.lower() if extracted1 else ""
                    extracted2_l = extracted2.lower() if extracted2 else ""

                    if (
                        (
                            ("não há nenhum modal" in extracted1_l)
                            or ("lobby" in extracted1_l)
                            or ("modal não detectado" in extracted1_l)
                        )
                        and (
                            ("não há nenhum modal" in extracted2_l)
                            or ("lobby" in extracted2_l)
                            or ("modal não detectado" in extracted2_l)
                        )
                    ):
                        logging.info(
                            "execute_vsloader: Lobby identificado (nenhum modal detectado) em ambos os prints (baixa). "
                            "Texto 1: %s Texto 2: %s", extracted1, extracted2
                        )
                        break

                    combined_extracted = (extracted1 + " " + extracted2) if extracted1 and extracted2 else (extracted1 or extracted2)
                    combined_extracted = combined_extracted.strip().lower()

                    if combined_extracted == repeated_text:
                        repeat_count += 1
                    else:
                        repeated_text = combined_extracted
                        repeat_count = 0

                    if repeat_count >= 20:
                        logging.error(f"⚠️ Travamento detectado no shopping {shopping} (baixa). Mesmo modal repetido 20x seguidas: {combined_extracted[:100]}...")
                        kill_vsloader()
                        raise RuntimeError(f"Travamento detectado no shopping {shopping}")

                    screen_width = pyautogui.size().width
                    screen_height = pyautogui.size().height
                    center_x = screen_width // 2
                    center_y = screen_height // 2

                    if fuzzy_contains(combined_extracted, "alerta vssc") and fuzzy_contains(combined_extracted, "foram baixados"):
                        logging.info("execute_vsloader: Print indica ação ESC (baixa - LOG/ESC). Texto identificado: %s", combined_extracted)
                        for _ in range(7):
                            pyautogui.moveTo(center_x, center_y)
                            pyautogui.click()
                            pyautogui.press('esc')
                            time.sleep(1)
                        break
                    elif fuzzy_contains(combined_extracted, "alerta vssc") and fuzzy_contains(combined_extracted, "recria"):
                        logging.info("execute_vsloader: Alerta na baixa (Recria). Texto identificado: %s", combined_extracted)
                        pyautogui.moveTo(center_x, center_y)
                        pyautogui.click()
                        pyautogui.press("enter")
                        time.sleep(3)
                    elif fuzzy_contains(combined_extracted, "alerta vssc") and fuzzy_contains(combined_extracted, "não foi baixado"):
                        logging.info("execute_vsloader: Alerta na baixa (Recria). Texto identificado: %s", combined_extracted)
                        pyautogui.moveTo(center_x, center_y)
                        pyautogui.click()
                        for _ in range(8):
                            pyautogui.press("esc")
                            time.sleep(0.5)
                        
                        time.sleep(3)
                        break
                    elif fuzzy_contains(combined_extracted, "configurar impressão"):
                        logging.info("execute_vsloader: Configurar impressão (baixa). Texto identificado: %s", combined_extracted)
                        time.sleep(2)
                        pyautogui.press('tab')
                        time.sleep(0.3)
                        pyautogui.hotkey('alt', 'space')
                        time.sleep(0.3)
                        pyautogui.press('down')
                        time.sleep(0.3)
                        pyautogui.press('enter')
                        pyautogui.moveRel(100, 160)
                        pyautogui.click()
                        pyautogui.click()
                        time.sleep(1)
                        pyautogui.press('tab')
                        time.sleep(0.5)
                        pyautogui.press('enter')
                        time.sleep(3)
                        break
                    else:
                        logging.info("execute_vsloader: Nenhuma condição modal identificada (baixa). Texto identificado: %s", combined_extracted)
                    time.sleep(3)


                repeated_text = None
                repeat_count = 0

                while True:
                    screenshot1 = capture_screenshot()
                    extracted1 = analyze_screenshot(screenshot1)
                    time.sleep(2)
                    screenshot2 = capture_screenshot()
                    extracted2 = analyze_screenshot(screenshot2)
                    if extracted1 is None and extracted2 is None:
                        logging.info("execute_vsloader: Nenhum texto extraído nos prints (baixa - LOG/ESC). extracted1=None, extracted2=None.")
                        time.sleep(3)
                        continue
                    extracted1_l = extracted1.lower() if extracted1 else ""
                    extracted2_l = extracted2.lower() if extracted2 else ""

                    if (
                        (
                            ("não há nenhum modal" in extracted1_l)
                            or ("lobby" in extracted1_l)
                            or ("modal não detectado" in extracted1_l)
                        )
                        and (
                            ("não há nenhum modal" in extracted2_l)
                            or ("lobby" in extracted2_l)
                            or ("modal não detectado" in extracted2_l)
                        )
                    ):
                        logging.info(
                            "execute_vsloader: Lobby identificado (nenhum modal detectado) em ambos os prints (baixa - LOG/ESC). "
                            "Texto 1: %s Texto 2: %s", extracted1, extracted2
                        )
                        break

                    combined_extracted = (extracted1 + " " + extracted2) if extracted1 and extracted2 else (extracted1 or extracted2)
                    combined_extracted = combined_extracted.strip().lower()

                    if combined_extracted == repeated_text:
                        repeat_count += 1
                    else:
                        repeated_text = combined_extracted
                        repeat_count = 0

                    if repeat_count >= 20:
                        logging.error(f"⚠️ Travamento detectado no shopping {shopping} (baixa - LOG/ESC). Mesmo modal repetido 20x seguidas: {combined_extracted[:100]}...")
                        kill_vsloader()
                        raise RuntimeError(f"Travamento detectado no shopping {shopping}")

                    screen_width = pyautogui.size().width
                    screen_height = pyautogui.size().height
                    center_x = screen_width // 2
                    center_y = screen_height // 2

                    if fuzzy_contains(combined_extracted, "alerta vssc") or fuzzy_contains(combined_extracted, "log"):
                        logging.info("execute_vsloader: Print indica ação ESC (baixa - LOG/ESC). Texto identificado: %s", combined_extracted)
                        for _ in range(7):
                            pyautogui.moveTo(center_x, center_y)
                            pyautogui.click()
                            pyautogui.press('esc')
                            time.sleep(1)
                        break
                    else:
                        logging.info("execute_vsloader: Nenhuma condição modal identificada (baixa - LOG/ESC). Texto identificado: %s", combined_extracted)
                    time.sleep(3)


        else:
            print(f"Shopping de fora do estado")
            time.sleep(10)

            pyautogui.hotkey('alt', 's')
            time.sleep(0.5)
            pyautogui.press('down')
            time.sleep(0.3)
            pyautogui.press('enter')
            time.sleep(8)

            base_ret = r"C:\AUTOMACAO\conciliacao"
            ret_dirs = {
                "EMPREENDEDOR": os.path.join(base_ret, "ret_emp"),
                "CONDOMINIO":   os.path.join(base_ret, "ret_con"),
                "FPP":          os.path.join(base_ret, "ret_fpp"),
                "FUNDO":        os.path.join(base_ret, "ret_fpp"),
                "FUNDONOVO":    os.path.join(base_ret, "ret_fpp"),
            }

            # ---- ordem e seleção de rubricas conforme portador_map do shopping ----
            variant_ports = portador_map.get(determine_variant(shopping), [])
            ordered_rubricas = []
            selected_port_by_rubrica = {}
            for p in variant_ports:
                rub = p["rubrica"].upper()
                if rub not in selected_port_by_rubrica:
                    selected_port_by_rubrica[rub] = p           # usa o PRIMEIRO portador da rubrica
                    ordered_rubricas.append(rub)                # mantém a ordem de aparição

            # ---- monta lista de arquivos por rubrica e contabiliza em variável global ----
            global RUBRICA_COUNTS
            RUBRICA_COUNTS = {}
            arquivos_por_rubrica = {}
            for rub in ("EMPREENDEDOR", "CONDOMINIO", "FPP", "FUNDO", "FUNDONOVO"):
                pasta = ret_dirs[rub]
                if os.path.isdir(pasta):
                    lst = [f for f in os.listdir(pasta) if f.lower().endswith(".ret")]
                    lst.sort()
                else:
                    lst = []
                arquivos_por_rubrica[rub] = lst
                RUBRICA_COUNTS[rub] = len(lst)

            # ====================== PROCESSO POR RUBRICA ==========================
            for rub in ordered_rubricas:
                portador = selected_port_by_rubrica[rub]
                lista_arquivos = arquivos_por_rubrica.get(rub, [])
                if not lista_arquivos:
                    logging.info(f"Nenhum arquivo para rubrica {rub}. Pulando.")
                    continue

                logging.info(f"Iniciando rubrica {rub} (portador {portador['codigo']}) com {len(lista_arquivos)} arquivo(s).")

                for idx, nome_arquivo in enumerate(lista_arquivos, start=1):

                    fullpath_local = os.path.join(ret_dirs[rub], nome_arquivo)

                    # -------------------- INTEGRAÇÃO --------------------
                    pyautogui.hotkey('alt', 's')
                    time.sleep(0.5)
                    for _ in range(1):
                        pyautogui.press('right')
                        time.sleep(0.3)
                    for _ in range(2):
                        pyautogui.press('down')
                        time.sleep(0.3)
                    pyautogui.press('right')
                    time.sleep(0.3)
                    for _ in range(3):
                        pyautogui.press('down')
                        time.sleep(0.3)
                    pyautogui.press('enter')
                    time.sleep(3)
                    pyautogui.hotkey('alt', 'space')
                    time.sleep(0.3)
                    pyautogui.press('down')
                    time.sleep(0.3)
                    pyautogui.press('enter')
                    pyautogui.moveRel(110, 45)
                    pyautogui.click()
                    pyautogui.click()
                    time.sleep(1)
                    pyautogui.hotkey('ctrl', 'f')
                    time.sleep(0.5)
                    pyautogui.typewrite(portador["codigo"])
                    pyautogui.press('enter')
                    time.sleep(3)
                    pyautogui.press('down')
                    time.sleep(0.5)
                    pyautogui.press('up')

                    repeated_text = None
                    repeat_count = 0

                    while True:
                        screenshot1 = capture_screenshot()
                        extracted1 = analyze_screenshot(screenshot1)
                        time.sleep(2)
                        screenshot2 = capture_screenshot()
                        extracted2 = analyze_screenshot(screenshot2)
                        if extracted1 is None and extracted2 is None:
                            logging.info("execute_vsloader: Nenhum texto extraído nos prints (fora do estado). extracted1=None, extracted2=None.")
                            time.sleep(3)
                            continue
                        extracted1_l = extracted1.lower() if extracted1 else ""
                        extracted2_l = extracted2.lower() if extracted2 else ""
                        if (
                            ("não há nenhum modal" in extracted1_l or "lobby" in extracted1_l or "modal não detectado" in extracted1_l) and
                            ("não há nenhum modal" in extracted2_l or "lobby" in extracted2_l or "modal não detectado" in extracted2_l)
                        ):
                            logging.info("execute_vsloader: Lobby identificado (nenhum modal detectado) em ambos os prints (fora do estado). Texto 1: %s Texto 2: %s", extracted1, extracted2)
                            break
                        combined_extracted = (extracted1 + " " + extracted2) if extracted1 and extracted2 else (extracted1 or extracted2)
                        combined_extracted = combined_extracted.strip().lower()

                        if combined_extracted == repeated_text:
                            repeat_count += 1
                        else:
                            repeated_text = combined_extracted
                            repeat_count = 0

                        if repeat_count >= 20:
                            logging.error(f"⚠️ Travamento detectado no shopping {shopping} (fora do estado). Mesmo modal repetido 20x seguidas: {combined_extracted[:100]}...")
                            kill_vsloader()
                            raise RuntimeError(f"Travamento detectado no shopping {shopping}")

                        screen_width = pyautogui.size().width
                        screen_height = pyautogui.size().height
                        center_x = screen_width // 2
                        center_y = screen_height // 2
                        if fuzzy_contains(combined_extracted, "alerta vssc"):
                            logging.info("execute_vsloader: Print indica ação ENTER (fora do estado). Texto identificado: %s", combined_extracted)
                            pyautogui.moveTo(center_x, center_y)
                            pyautogui.click()
                            pyautogui.press("enter")
                            time.sleep(1)
                            break
                        elif fuzzy_contains(combined_extracted, "leitura do arquivo cnab"):
                            logging.info("execute_vsloader: Tela 'Leitura do Arquivo CNAB' identificada (fora do estado). Texto identificado: %s", combined_extracted)
                            break
                        else:
                            logging.info("execute_vsloader: Nenhuma condição modal identificada (fora do estado). Texto identificado: %s", combined_extracted)
                        time.sleep(3)

                    time.sleep(2)

                    folder_selecionado = os.path.dirname(fullpath_local)

                    # Garante que o seletor de arquivo está ativo
                    time.sleep(1)


                    pyautogui.hotkey('alt', 'space')
                    time.sleep(0.3)
                    pyautogui.press('down')
                    time.sleep(0.3)
                    pyautogui.press('enter')
                    pyautogui.moveRel(210,110)
                    pyautogui.click()
                    pyautogui.click()
                    # Vai direto para a barra de endereço (atalho universal)
                                        # Vai direto para a barra de endereço
                    # 🔧 Garante que o seletor de arquivo está ativo
                    time.sleep(3)

                    # Abre o menu do sistema da janela ("Restaurar, Mover, Tamanho..." etc.)
                    pyautogui.hotkey('alt', 'space')
                    time.sleep(0.3)
                    pyautogui.press('down')
                    time.sleep(0.3)
                    pyautogui.press('enter')

                    # Move o foco para o campo de endereço (coordenadas seguras)
                    pyautogui.moveRel(0, 200)  # <== AJUSTE: move o mouse até a barra de endereço
                    pyautogui.click()
                    pyautogui.click()
                    time.sleep(0.5)
                    for i in range(70):
                        pyautogui.press('backspace')
                   
                    # Digita o caminho completo da pasta correta
                    pyautogui.typewrite(folder_selecionado)
                    time.sleep(1)

                    # 💡 Faz o print antes de apertar Enter, pra registrar o caminho digitado
                    print_path_debug = os.path.join(RUN_PRINTS_DIR, f"debug_path_{datetime.now().strftime('%H%M%S')}.png")
                    pyautogui.screenshot(print_path_debug)
                    logging.info(f"[DEBUG] Screenshot do caminho digitado salvo em: {print_path_debug}")

                    # Pressiona Enter para confirmar
                    pyautogui.press('enter')
                    time.sleep(2)



                    time.sleep(3)
                    pyautogui.hotkey('alt', 'space')
                    time.sleep(0.3)
                    pyautogui.press('down')
                    time.sleep(0.3)
                    pyautogui.press('enter')
                    pyautogui.moveRel(-200,65)
                    pyautogui.click()
                    pyautogui.click()
                    time.sleep(2)

                    posicao_arquivo = idx  # 1, 2, 3...
                    pyautogui.moveRel(250, 0)
                    pyautogui.click()
                    logging.info(f"Selecionando o {posicao_arquivo}º arquivo da rubrica {rub}")
                    time.sleep(1)
                    pyautogui.press('down')
                    time.sleep(0.3)
                    for i in range(8):
                        pyautogui.press('up')
                        time.sleep(0.3)
                    time.sleep(1)

                    # agora seleciona o arquivo correto baseado na posição (idx)
                    for _ in range(posicao_arquivo - 1):
                        pyautogui.press('down')
                        time.sleep(0.3)

                    pyautogui.press('enter')
                    time.sleep(2)
                    pyautogui.press('enter')
                    repeated_text = None
                    repeat_count = 0

                    while True:
                        screenshot1 = capture_screenshot()
                        extracted1 = analyze_screenshot(screenshot1)
                        time.sleep(2)
                        screenshot2 = capture_screenshot()
                        extracted2 = analyze_screenshot(screenshot2)
                        if extracted1 is None and extracted2 is None:
                            logging.info("execute_vsloader: Nenhum texto extraído nos prints (fora do estado/033). extracted1=None, extracted2=None.")
                            time.sleep(3)
                            continue
                        combined_tmp = (extracted1 or "") + " " + (extracted2 or "")
                        combined_tmp = combined_tmp.strip().lower()

                        if combined_tmp == repeated_text:
                            repeat_count += 1
                        else:
                            repeated_text = combined_tmp
                            repeat_count = 0

                        if repeat_count >= 20:
                            logging.error(f"⚠️ Travamento detectado no shopping {shopping} (fora do estado/033). Mesmo modal repetido 20x seguidas: {combined_tmp[:100]}...")
                            kill_vsloader()
                            raise RuntimeError(f"Travamento detectado no shopping {shopping}")

                        if fuzzy_contains(combined_tmp, "leitura concluída"):
                            logging.info("execute_vsloader: Leitura concluída (fora do estado/033). Texto identificado: %s", combined_tmp)
                            cx = pyautogui.size().width // 2
                            cy = pyautogui.size().height // 2
                            pyautogui.moveTo(cx, cy)
                            pyautogui.click()
                            break
                        elif fuzzy_contains(combined_tmp, "alerta vssc"):
                            logging.info("execute_vsloader: Alerta na baixa (Recria). Texto identificado: %s", combined_tmp)
                            pyautogui.moveTo(center_x, center_y)
                            pyautogui.click()
                            pyautogui.press("enter")
                            time.sleep(3)
                        else:
                            logging.info("execute_vsloader: Aguardando 'LEITURA CONCLUÍDA' (fora do estado/033). Texto identificado: %s", combined_tmp)
                        time.sleep(3)

                    for _ in range(4):
                        pyautogui.press('esc')
                        time.sleep(0.3)
                    

                    # -------------------- BAIXA (do MESMO arquivo) --------------------
                    pyautogui.hotkey('alt', 's')
                    time.sleep(0.5)
                    for _ in range(1):
                        pyautogui.press('right')
                        time.sleep(0.3)
                    for _ in range(2):
                        pyautogui.press('down')
                        time.sleep(0.3)
                    pyautogui.press('right')
                    time.sleep(0.3)
                    for _ in range(5):
                        pyautogui.press('down')
                        time.sleep(0.3)
                    pyautogui.press('enter')
                    time.sleep(2)
                    pyautogui.press('down')
                    time.sleep(1)
                    pyautogui.hotkey('ctrl', 'f')
                    time.sleep(0.5)
                    pyautogui.typewrite(portador["codigo"])
                    pyautogui.press('enter')
                    time.sleep(1)
                    pyautogui.press('enter')
                    time.sleep(3)
                    pyautogui.press('down')
                    time.sleep(1)
                    pyautogui.press('up')
                    time.sleep(1)
                    pyautogui.hotkey('alt', 'space')
                    time.sleep(0.3)
                    pyautogui.press('down')
                    time.sleep(0.3)
                    pyautogui.press('enter')
                    pyautogui.moveRel(30, 325)
                    pyautogui.click()
                    pyautogui.click()
                    
                    time.sleep(2)
                    repeated_text = None
                    repeat_count = 0

                    while True:
                        screenshot1 = capture_screenshot()
                        extracted1 = analyze_screenshot(screenshot1)
                        time.sleep(2)
                        screenshot2 = capture_screenshot()
                        extracted2 = analyze_screenshot(screenshot2)
                        if extracted1 is None and extracted2 is None:
                            logging.info("execute_vsloader: Nenhum texto extraído nos prints (fora do estado/baixa). extracted1=None, extracted2=None.")
                            time.sleep(3)
                            continue
                        extracted1_l = extracted1.lower() if extracted1 else ""
                        extracted2_l = extracted2.lower() if extracted2 else ""
                        if (
                            ("não há nenhum modal" in extracted1_l or "lobby" in extracted1_l or "modal não detectado" in extracted1_l) and
                            ("não há nenhum modal" in extracted2_l or "lobby" in extracted2_l or "modal não detectado" in extracted2_l)
                        ):
                            logging.info("execute_vsloader: Lobby identificado (nenhum modal detectado) em ambos os prints (fora do estado/baixa). Texto 1: %s Texto 2: %s", extracted1, extracted2)
                            break
                        combined_extracted = (extracted1 + " " + extracted2) if extracted1 and extracted2 else (extracted1 or extracted2)
                        combined_extracted = combined_extracted.strip().lower()

                        if combined_extracted == repeated_text:
                            repeat_count += 1
                        else:
                            repeated_text = combined_extracted
                            repeat_count = 0

                        if repeat_count >= 20:
                            logging.error(f"⚠️ Travamento detectado no shopping {shopping} (fora do estado/baixa). Mesmo modal repetido 20x seguidas: {combined_extracted[:100]}...")
                            kill_vsloader()
                            raise RuntimeError(f"Travamento detectado no shopping {shopping}")

                        screen_width = pyautogui.size().width
                        screen_height = pyautogui.size().height
                        center_x = screen_width // 2
                        center_y = screen_height // 2
                        if fuzzy_contains(combined_extracted, "alerta vssc") and fuzzy_contains(combined_extracted, "foram baixados"):
                            logging.info("execute_vsloader: Print indica ação ESC (baixa - LOG/ESC). Texto identificado: %s", combined_extracted)
                            for _ in range(7):
                                pyautogui.moveTo(center_x, center_y)
                                pyautogui.click()
                                pyautogui.press('esc')
                                time.sleep(1)
                            break
                        elif fuzzy_contains(combined_extracted, "alerta vssc") and fuzzy_contains(combined_extracted, "recria"):
                            logging.info("execute_vsloader: Alerta na baixa (fora do estado/Recria). Texto identificado: %s", combined_extracted)
                            pyautogui.moveTo(center_x, center_y)
                            pyautogui.click()
                            pyautogui.press("enter")
                            time.sleep(3)
                        elif fuzzy_contains(combined_extracted, "alerta vssc"):
                            logging.info("execute_vsloader: Erro de alerta. Texto identificado: %s", combined_extracted)
                            for _ in range(5):
                                pyautogui.press('esc')
                                time.sleep(0.5)
                            time.sleep(3)
                            break
                        elif fuzzy_contains(combined_extracted, "configurar impressão"):
                            logging.info("execute_vsloader: Configurar impressão (fora do estado/baixa). Texto identificado: %s", combined_extracted)
                            time.sleep(2)
                            pyautogui.press('tab')
                            time.sleep(0.3)
                            pyautogui.hotkey('alt', 'space')
                            time.sleep(0.3)
                            pyautogui.press('down')
                            time.sleep(0.3)
                            pyautogui.press('enter')
                            pyautogui.moveRel(100, 160)
                            pyautogui.click()
                            pyautogui.click()
                            time.sleep(1)
                            pyautogui.press('tab')
                            time.sleep(0.5)
                            pyautogui.press('enter')
                            time.sleep(3)
                            break
                        else:
                            logging.info("execute_vsloader: Nenhuma condição modal identificada (fora do estado/baixa). Texto identificado: %s", combined_extracted)
                        time.sleep(3)


                    repeated_text = None
                    repeat_count = 0

                    while True:
                        screenshot1 = capture_screenshot()
                        extracted1 = analyze_screenshot(screenshot1)
                        time.sleep(2)
                        screenshot2 = capture_screenshot()
                        extracted2 = analyze_screenshot(screenshot2)
                        if extracted1 is None and extracted2 is None:
                            logging.info("execute_vsloader: Nenhum texto extraído nos prints (fora do estado/baixa - LOG/ESC). extracted1=None, extracted2=None.")
                            time.sleep(3)
                            continue
                        extracted1_l = extracted1.lower() if extracted1 else ""
                        extracted2_l = extracted2.lower() if extracted2 else ""
                        if (
                            ("não há nenhum modal" in extracted1_l or "lobby" in extracted1_l or "modal não detectado" in extracted1_l) and
                            ("não há nenhum modal" in extracted2_l or "lobby" in extracted2_l or "modal não detectado" in extracted2_l)
                        ):
                            logging.info("execute_vsloader: Lobby identificado (nenhum modal detectado) em ambos os prints (fora do estado/baixa - LOG/ESC). Texto 1: %s Texto 2: %s", extracted1, extracted2)
                            break
                        combined_extracted = (extracted1 + " " + extracted2) if extracted1 and extracted2 else (extracted1 or extracted2)
                        combined_extracted = combined_extracted.strip().lower()

                        if combined_extracted == repeated_text:
                            repeat_count += 1
                        else:
                            repeated_text = combined_extracted
                            repeat_count = 0

                        if repeat_count >= 20:
                            logging.error(f"⚠️ Travamento detectado no shopping {shopping} (fora do estado/baixa - LOG/ESC). Mesmo modal repetido 20x seguidas: {combined_extracted[:100]}...")
                            kill_vsloader()
                            raise RuntimeError(f"Travamento detectado no shopping {shopping}")

                        screen_width = pyautogui.size().width
                        screen_height = pyautogui.size().height
                        center_x = screen_width // 2
                        center_y = screen_height // 2
                        if fuzzy_contains(combined_extracted, "alerta vssc") or fuzzy_contains(combined_extracted, "log"):
                            logging.info("execute_vsloader: Print indica ação ESC (fora do estado/baixa - LOG/ESC). Texto identificado: %s", combined_extracted)
                            for _ in range(7):
                                pyautogui.moveTo(center_x, center_y)
                                pyautogui.click()
                                pyautogui.press('esc')
                                time.sleep(1)
                            break
                        else:
                            logging.info("execute_vsloader: Nenhuma condição modal identificada (fora do estado/baixa - LOG/ESC). Texto identificado: %s", combined_extracted)
                        time.sleep(3)




    


                
                
        delete_all_prints()
        time.sleep(2)
        pyautogui.hotkey('alt', 'F4')
        time.sleep(2)
        pyautogui.moveTo(center_x, center_y)
        pyautogui.click()
        pyautogui.hotkey('alt', 'F4')

            
        # Fim do laço para cada portador
    except ElementNotFoundError as e:
        logging.error(f"Elemento não encontrado: {e}")
    except TimeoutError as e:
        logging.error(f"Tempo limite excedido: {e}")
    except Exception as e:
        logging.error(f"Erro inesperado: {e}")

# Novo ponto de entrada para chamar apenas com o shopping como argumento
if __name__ == "__main__":
    if len(sys.argv) > 1:
        shopping = sys.argv[1]
        print(f"[DEBUG] Argumento recebido: {shopping}")

        # Configura o log DO SHOPPING antes de qualquer logging.info
        variant = determine_variant(shopping)
        setup_logging_for_shopping(variant)

        logging.info(f"Iniciando automação para: {shopping}")
        execute_vsloader(shopping)
    else:
        print("[ERRO] Nenhum shopping informado ao chamar o executável.")
