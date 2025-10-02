# utils.py
import os
import sys
import time
import logging
import pyautogui
import ctypes
import base64
import difflib
import openai
from datetime import date, datetime


# Cria uma pasta para os prints, se não existir.
prints_folder = os.path.join(os.getcwd(), "prints")
if not os.path.exists(prints_folder):
    os.makedirs(prints_folder)

# Variável global (caso seja necessária em outras partes do código)
IS_SEGURO = False


folder_map = {
    "Shopping da Ilha": r"C:\Program Files\Victor & Schellenberger\VSSC_ILHA",
    "Shopping Mestre Álvaro": r"C:\Program Files\Victor & Schellenberger\VSSC_MESTREALVARO",
    "Shopping Metrópole": r"C:\Program Files\Victor & Schellenberger\VSSC_METROPOLE",
    "Shopping Montserrat": r"C:\Program Files\Victor & Schellenberger\VSSC_MONTSERRAT",
    "Shopping Moxuara": r"C:\Program Files\Victor & Schellenberger\VSSC_MOXUARA",
    "Shopping Praia da Costa": r"C:\Program Files\Victor & Schellenberger\VSSC_PRAIADACOSTA",
    "Shopping Rio Poty": r"C:\Program Files\Victor & Schellenberger\VSSC_TERESINA"
}

# Se o módulo for executado diretamente, lê os parâmetros da linha de comando.
if len(sys.argv) >= 3:
    shopping_escolhido = sys.argv[1]
    tipo_escolhido = sys.argv[2]
else:
    shopping_escolhido = None
    tipo_escolhido = None

# Define a variável 'folder' com base no shopping escolhido.
folder = folder_map.get(
    shopping_escolhido,
    r"C:\Program Files\Victor & Schellenberger_FAT\VSSC_MONTSERRAT"
) if shopping_escolhido else None

# Funções para manipulação de foco usando ctypes
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

def wait_for_focus_change(prev_handle, max_wait=40):
    start_time = time.time()
    while True:
        if time.time() - start_time > max_wait:
            logging.info("Tempo máximo de espera por mudança de foco atingido.")
            break
        time.sleep(1)
        current_handle = get_foreground_window()
        if current_handle != prev_handle:
            break

def login(custom_folder: str = None) -> None:
    
    screen_width, screen_height = pyautogui.size()
    center_x = screen_width // 2
    center_y = screen_height // 2
    folder_to_use = custom_folder if custom_folder is not None else folder



    time.sleep(8)
    pyautogui.press('win')
    time.sleep(4)
    pyautogui.write('file explorer')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(14)

    pyautogui.hotkey('alt', 'd')
    pyautogui.typewrite(folder_to_use)
    pyautogui.press('enter')
    time.sleep(3)

    pyautogui.typewrite("VSLOADER.EXE")
    pyautogui.press('enter')
    # time.sleep(10)

    logging.info("VSLOADER.EXE executado.")
    time.sleep(3)

    while True:
        screenshot1 = capture_screenshot()
        extracted1 = analyze_screenshot(screenshot1)
        time.sleep(2)
        screenshot2 = capture_screenshot()
        extracted2 = analyze_screenshot(screenshot2)

        if not extracted1 or not extracted2:
            time.sleep(3)
            continue

        extracted1_l = extracted1.lower() if extracted1 else ""
        extracted2_l = extracted2.lower() if extracted2 else ""

        

        combined = (extracted1 + " " + extracted2) if extracted1 and extracted2 else (extracted1 or extracted2)
        if fuzzy_contains(combined, "<ESCjljjlklkjjkl>"):
            pyautogui.click(center_x, center_y)
            pyautogui.press('esc')
        
        
        elif fuzzy_contains(combined, "Usuário") and fuzzy_contains(combined, "Senha"):
            logging.info("Usuário e senha")
            time.sleep(2)
            break
            
        else:
            logging.info("Nenhuma tela detectada")
            logging.info(combined)
            time.sleep(5)


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
    while True:
        screenshot = capture_screenshot()
        extracted = analyze_screenshot(screenshot)

        if not extracted:
            time.sleep(3)
            continue

        extracted_l = extracted.lower()

        if fuzzy_contains(extracted, "<ESC>"):
            pyautogui.click(center_x, center_y)
            pyautogui.press('esc')

        elif fuzzy_contains(extracted, "Competência de Trabalho") and fuzzy_contains(extracted, "Alerta VSSC"):
            logging.info("competencia de trabalho detectada")
            time.sleep(12)
            break

        elif fuzzy_contains(extracted, "Competência de Trabalho"):
            logging.info("competencia de trabalho detectada")
            break

        elif fuzzy_contains(extracted, "Alerta VSSC"):
            logging.info("Alerta VSSC detectado")
            pyautogui.click(center_x, center_y)
            pyautogui.press('esc')

        elif fuzzy_contains(extracted, "Contratos com término"):
            logging.info("Contrato com término detectado")
            for _ in range(5):
                pyautogui.press('esc')
            time.sleep(2)
            break

        elif fuzzy_contains(extracted, "Demonstrativos"):
            logging.info("Abriu 100%")
            break

        else:
            logging.info("Nenhuma tela detectada")
            logging.info(extracted)
            time.sleep(5)




    time.sleep(3)

def gerar_competencia(tipo: str, from_calculos: bool = False) -> None:
    screen_width, screen_height = pyautogui.size()
    center_x = screen_width // 2
    center_y = screen_height // 2

    if tipo in ["Postecipados", "Atípicos"]:
        logging.info("Alterando para a competência correta")
        time.sleep(2)
        pyautogui.click(center_x, center_y)
        pyautogui.hotkey('alt', 's')
        for _ in range(8):
            pyautogui.press('right'); time.sleep(0.3)
        pyautogui.press('down'); time.sleep(0.3)
        pyautogui.press('right'); time.sleep(0.3)
        pyautogui.press('down'); time.sleep(0.3)
        pyautogui.press('enter'); time.sleep(5)
        pyautogui.click(center_x, center_y)
        time.sleep(0.5)
        pyautogui.keyDown('ctrl'); pyautogui.press('f'); pyautogui.keyUp('ctrl')
        time.sleep(1)
        now = datetime.now()
        mes_ano = now.strftime("%m/%Y")
        pyautogui.typewrite(mes_ano, interval=0.1)
        pyautogui.press('enter') 
        time.sleep(2)
        while True:
            extracted1 = analyze_screenshot(capture_screenshot()) or ""
            time.sleep(2)
            extracted2 = analyze_screenshot(capture_screenshot()) or ""
            if not extracted1 and not extracted2:
                time.sleep(3)
                continue
            combined = f"{extracted1} {extracted2}".strip()
            if fuzzy_contains(combined, "Alerta VSSC"):
                pyautogui.click(center_x, center_y)
                pyautogui.press('enter')
                time.sleep(0.5)
                for _ in range(4):
                    pyautogui.press('esc'); time.sleep(0.5)
                time.sleep(3)
                return
            else:
                for _ in range(2):
                    pyautogui.press('enter'); time.sleep(1)
                for _ in range(5):
                    pyautogui.press('esc'); time.sleep(0.5)
                time.sleep(3)
                return

    elif tipo == "Antecipados" and from_calculos:

        logging.info("Atualizando para a competência correta. Antecipados do gerar cálculos")
        time.sleep(2)
        pyautogui.click(center_x, center_y)
        pyautogui.hotkey('alt', 's')
        for _ in range(8):
            pyautogui.press('right'); time.sleep(0.3)
        pyautogui.press('down'); time.sleep(0.3)
        pyautogui.press('right'); time.sleep(0.3)
        pyautogui.press('enter'); time.sleep(2)
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
            if fuzzy_contains(combined, "<ESC>"):
                pyautogui.click(center_x, center_y)
                pyautogui.press('esc')
            elif fuzzy_contains(combined, "Alerta VSSC"):
                pyautogui.click(center_x, center_y)
                pyautogui.press('enter')
                break
                
            else:
                logging.info("Nenhuma tela detectada")
                logging.info(combined)
        time.sleep(2)
        logging.info("Dormiu 8 segundos para esperar o carregamento do sistema")

        pyautogui.hotkey('alt', 'space'); time.sleep(0.3)
        pyautogui.press('down'); time.sleep(0.3)
        pyautogui.press('enter')
        pyautogui.moveRel(130, 70)
        pyautogui.click(); pyautogui.click()
        time.sleep(1)
        current_month = date.today().month
        if current_month in [4, 8]:
            pyautogui.press('down'); time.sleep(0.3)
            for _ in range(3):
                pyautogui.press('left'); time.sleep(0.3)
        else:
            pyautogui.press('right'); time.sleep(0.3)
        pyautogui.press('enter'); time.sleep(0.3)
        for _ in range(2):
            time.sleep(0.5); pyautogui.press('enter')

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
            if fuzzy_contains(combined, "<ESC>"):
                pyautogui.click(center_x, center_y)
                pyautogui.press('esc')
            elif fuzzy_contains(combined, "Competência de Trabalho") and fuzzy_contains(combined, "Alerta VSSC"):
                logging.info("Alerta com competencia")
                pyautogui.click(center_x, center_y)
                pyautogui.press('enter')
            elif fuzzy_contains(combined, "Regerar") and fuzzy_contains(combined, "Alerta VSSC"):
                logging.info("Alerta pra regerar")
                pyautogui.click(center_x, center_y)
                pyautogui.press('enter')
            elif fuzzy_contains(combined, "Alerta VSSC"):
                logging.info("Alerta puro")
                pyautogui.click(center_x, center_y)
                pyautogui.press('esc')
            
            elif fuzzy_contains(combined, "Contratos com término"):
                logging.info("Contrato com término detectado")
                for _ in range(5):
                    pyautogui.press('esc')
                time.sleep(2)
                break
                
            else:
                logging.info("Nenhuma tela detectada")
                logging.info(combined)

        return

    elif tipo == "Antecipados" and not from_calculos:
        logging.info("Antecipado de gerar boletos ou enviar emails")
        time.sleep(2)
        pyautogui.click(center_x, center_y)
        pyautogui.hotkey('alt', 's')
        for _ in range(8):
            pyautogui.press('right'); time.sleep(0.3)
        pyautogui.press('down'); time.sleep(0.3)
        pyautogui.press('right'); time.sleep(0.3)
        pyautogui.press('down'); time.sleep(0.3)
        pyautogui.press('enter'); time.sleep(5)
        pyautogui.click(center_x, center_y)
        time.sleep(0.5)
        pyautogui.keyDown('ctrl'); pyautogui.press('f'); pyautogui.keyUp('ctrl')
        time.sleep(1)
        # --- começa alteração ---
        now = datetime.now()
        # calcula mês seguinte, ajustando ano se for dezembro
        next_year = now.year + (now.month == 12)
        next_month = now.month % 12 + 1
        mes_ano = f"{next_month:02d}/{next_year}"
        # --- termina alteração ---
        pyautogui.typewrite(mes_ano, interval=0.1)
        pyautogui.press('enter')
        time.sleep(2)
        while True:
            extracted1 = analyze_screenshot(capture_screenshot()) or ""
            time.sleep(2)
            extracted2 = analyze_screenshot(capture_screenshot()) or ""
            if not extracted1 and not extracted2:
                time.sleep(3)
                continue
            combined = f"{extracted1} {extracted2}".strip()
            if fuzzy_contains(combined, "Alerta VSSC"):
                pyautogui.click(center_x, center_y)
                pyautogui.press('enter')
                time.sleep(0.5)
                for _ in range(4):
                    pyautogui.press('esc'); time.sleep(0.5)
                time.sleep(3)
                return
            else:
                for _ in range(2):
                    pyautogui.press('enter'); time.sleep(1)
                for _ in range(5):
                    pyautogui.press('esc'); time.sleep(0.5)
                time.sleep(3)
                return


    else:
        logging.warning("Tipo de competência não reconhecido. Verifique se 'tipo' está correto.")
        return


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

def analyze_screenshot(image_path):
    """
    Converte a imagem para Base64 e envia para a API do GPT-4o, retornando o texto extraído.
    """
    if not os.path.exists(image_path):
        logging.error("Nenhuma imagem válida encontrada para análise.")
        return None
    with open(image_path, "rb") as img_file:
        base64_image = base64.b64encode(img_file.read()).decode("utf-8")
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4-turbo",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": ("Identifique a tela"
                                     "e me devolva seu conteúdo por escrito, em todas as janelas abertas no print, não apenas a que está em foco. Normalmente essa janela vai ter título "
                                     "como 'Gerando Área de Recibo' ou 'Alerta VSSC', lembrando que a análise deve ser feita "
                                     "principalmente se há um modal ou tela no meu aplicativo aberta além da tela principal. "
                                     "Se não houver nenhum modal ou tela, pode-se concluir que o sistema está no lobby. "
                                     "Muita atenção ao conteúdo de cada modal. EU PRECISO QUE O TEXTO DA RESPOSTA SEMPRE COMECE "
                                     "COM 'MODAL DETECTADO' OU 'MODAL NÃO DETECTADO'. Além disso me retorne TODO conteúdo visto na tela, não só o modal 'principal', mas tudo que for visível, ")
                        },
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:image/png;base64,{base64_image}"}
                        },
                    ],
                }
            ],
            request_timeout=60
        )
        extracted_text = response["choices"][0]["message"]["content"].strip()
        return extracted_text
    except Exception as e:
        if "base64" not in str(e):
            logging.error(f"Erro ao analisar imagem: {e}")
        return None
    
def analyze_screenshot_login(image_path):
    """
    Converte a imagem para Base64 e envia para a API do GPT-4o, retornando o texto extraído.
    """
    if not os.path.exists(image_path):
        logging.error("Nenhuma imagem válida encontrada para análise.")
        return None
    with open(image_path, "rb") as img_file:
        base64_image = base64.b64encode(img_file.read()).decode("utf-8")
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": ("Identifique a tela em foco no print, normalmente um modal ou um alerta, "
                                     "e me devolva seu conteúdo por escrito, muita atenção a TODOS OS COMPONENTES DA PÁGINA. Eu preciso de TUDO que for escrito, até as menores coisas")
                        },
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:image/png;base64,{base64_image}"}
                        },
                    ],
                }
            ],
            request_timeout=60
        )
        extracted_text = response["choices"][0]["message"]["content"].strip()
        logging.info("Texto retornado pelo GPT: %s", extracted_text)
        return extracted_text
    except Exception as e:
        if "base64" not in str(e):
            logging.error(f"Erro ao analisar imagem: {e}")
        return None


def capture_screenshot(retries: int = 3, delay: float = 0.5) -> str:
    """
    Tenta capturar a tela até `retries` vezes caso ocorra OSError.
    Se ainda falhar, propaga o erro.
    """
    for attempt in range(1, retries + 1):
        timestamp = int(time.time() * 1000)
        file_path = os.path.join(prints_folder, f"monitor_screenshot_{timestamp}.png")
        try:
            logging.info(f"[Screenshot] tentativa {attempt}/{retries}: salvando em {file_path}")
            pyautogui.screenshot(file_path)
            return file_path
        except OSError as e:
            logging.warning(f"[Screenshot] falhou na tentativa {attempt}: {e}")
            if attempt < retries:
                time.sleep(delay)
            else:
                logging.error(f"[Screenshot] não foi possível capturar após {retries} tentativas")
                raise


# Bloco de teste para executar o módulo diretamente
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s %(levelname)s: %(message)s",
                        datefmt='%d/%m/%Y %H:%M:%S')
    try:
        # Teste da função login:
        login()
        # Teste da função gerar_competencia: passando o tipo desejado.
        # Exemplo: "Antecipados" ou "Postecipados" ou "Atípicos"
        gerar_competencia("Antecipados")
    except Exception as e:
        logging.error("Erro: %s", e)
