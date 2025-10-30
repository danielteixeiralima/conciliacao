# -*- coding: utf-8 -*-

###############################################################################
#                              enviar_email.py                                #
###############################################################################

import ctypes
import pyautogui
import logging
import time
import os
import base64
from anthropic import Anthropic

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
from openpyxl import Workbook
from openpyxl import load_workbook
from hom_utils import login, gerar_competencia, folder_map
from rapidfuzz import fuzz
import unicodedata
import re

br_holidays = Brazil()

for w in Desktop(backend="uia").windows():
    logging.info(w.window_text())

openai.api_key = "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"

anthropic = Anthropic(api_key='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')

shopping_fases_tipo2 = {
    "Shopping Mestre Álvaro": {
        "Antecipados": [24, 25, 5, 7],
        "Atípicos": [31, 32, 6, 4],
        "Postecipados": [11, 2, 8, 41]
    },
    "Shopping Montserrat": {
        "Antecipados": [24, 25, 5, 7],
        "Atípicos": [31, 32, 6, 15],
        "Postecipados": [11, 4, 8, 22]
    },
    "Shopping Metrópole": {
        "Antecipados": [12, 13, 7, 18],
        "Atípicos": [31, 32, 6, 11],
        "Postecipados": [8, 9, 2, 36]
    },
    "Shopping Rio Poty": {
        "Antecipados": [12, 13, 7, 18],
        "Atípicos": [31, 32, 6, 11],
        "Postecipados": [8, 9, 2, 23]
    },
    "Shopping da Ilha": {
        "Antecipados": [12, 13, 7, 18],
        "Atípicos": [31, 32, 6, 11],
        "Postecipados": [8, 9, 2, 36, 37, 44]
    },
    "Shopping Moxuara": {
        "Antecipados": [12, 13, 24, 18],
        "Atípicos": [31, 32, 6, 11],
        "Postecipados": [8, 9, 2, 39]
    }
}


prints_folder = os.path.join(os.getcwd(), "prints")
# A cada screenshot, um nome único será gerado para acumular os prints
SCREENSHOT_PATH = os.path.join(prints_folder, "monitor_screenshot.png")
if not os.path.exists(prints_folder):
    os.makedirs(prints_folder)

IS_SEGURO = False

def normalize_text(s):
    s = s.lower()
    s = ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))
    s = re.sub(r'[^a-z0-9\s]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    # remove plural simples no final
    s = re.sub(r'\b(\w+)s\b', r'\1', s)
    return s

def fuzzy_contains(text, sub, threshold=0.75):  # já baixei um pouco o threshold
    text = normalize_text(text)
    sub = normalize_text(sub)

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
    # Todas as fases de 1 até 45 são consideradas, sem exclusão de fases
    for fase in range(1, 46):
        if fase < 14:
            coords[fase] = base_y + (fase - 1) * step_y
        else:
            coords[fase] = 215
    return coords

def click_fase_tipo1(shopping, fase):
    """
    Corrige o método de clicar na fase, considerando todas as fases.
    """
    if not fase:
        logging.error(f"Fase inválida para {shopping}. Verifique o mapeamento.")
        return
    vi = fase  # Considera que todas as fases são válidas
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
                                "text": (
                                    "Extraia **exatamente** o texto visível da imagem abaixo, como um OCR literal. "
                                    "Não interprete, não resuma, não complete informações, não descreva. "
                                    "Se houver um modal ou janela em primeiro plano, escreva apenas na primeira linha "
                                    "'MODAL DETECTADO'. Se não houver modal, escreva apenas na primeira linha "
                                    "'MODAL NÃO DETECTADO'. "
                                    "Depois dessa primeira linha, cole fielmente o texto visível do modal/tela, "
                                    "sem comentários adicionais, sem observações ou explicações."
                                ),
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
            time.sleep(5)

def find_and_click_button_with_retry(image_path, max_attempts=10, confidence_range=(0.95, 0.6)):
    try:
        for attempt in range(max_attempts):
            confidence = confidence_range[0] - (attempt * (confidence_range[0] - confidence_range[1]) / max_attempts)
            logging.info(f"Tentativa {attempt + 1}/{max_attempts} com confiança {confidence:.2f}")
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
    # Como não há exclusões, o índice visível é simplesmente o número da fase
    valid_fases = list(range(1, fase_val + 1))
    return len(valid_fases)

def execute_vsloader(shopping, tipo):
    login()

    # gerar_competencia(tipo_escolhido)
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

        fases = shopping_fases_tipo2.get(shopping, {}).get(tipo, [])
        folder = folder_map.get(shopping, r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_MESTREALVARO_HOM")

        screen_width, screen_height = pyautogui.size()
        center_x = screen_width // 2
        center_y = screen_height // 2

        
        
        for _ in range(5):
            pyautogui.press('esc')
            time.sleep(0.3)
        pyautogui.moveTo(center_x, center_y)
        pyautogui.click()
        current_phase = None
        time.sleep(2)

        excel_filename = f"{shopping}_envio_email.xlsx"
        # Dentro de execute_vsloader, antes de usar excel_filename:
        variant = determine_variant(shopping)
        logs_dir = os.path.join(os.getcwd(), "logs", variant)
        os.makedirs(logs_dir, exist_ok=True)

        excel_filename = os.path.join(logs_dir, f"{shopping}_envio_email.xlsx")
        


        
        for fase in fases:
            logging.info(f"Iniciando processo de envio de e-mails.")

            pyautogui.hotkey('alt', 's')
            for _ in range(8):
                pyautogui.press('right')
                time.sleep(0.3)
            for _ in range(7):
                pyautogui.press('down')
            pyautogui.press('right')
            time.sleep(0.3)
            pyautogui.press('down')
            time.sleep(0.3)
            pyautogui.press('enter')
            prev_handle = get_foreground_window()
            wait_for_focus_change(prev_handle)
            wait_for_stable_focus(prev_handle)
            pyautogui.hotkey('alt', 'space')
            time.sleep(0.3)
            pyautogui.press('down')
            time.sleep(0.3)
            pyautogui.press('enter')
            pyautogui.moveRel(350, 510)
            pyautogui.click()
            pyautogui.click()
            time.sleep(5)
            pyautogui.press('enter')
            prev_handle = get_foreground_window()
            wait_for_focus_change(prev_handle)
            wait_for_stable_focus(prev_handle)
            time.sleep(1)
            screen_width = pyautogui.size().width
            screen_height = pyautogui.size().height
            center_x = screen_width // 2
            center_y = screen_height // 2
            pyautogui.moveTo(center_x, center_y)
            pyautogui.click()
            time.sleep(0.5)
            pyautogui.hotkey('alt', 'space')
            time.sleep(0.3)
            pyautogui.press('down')
            time.sleep(0.3)
            pyautogui.press('enter')
            pyautogui.moveRel(180, 90)
            pyautogui.click()
            pyautogui.click()
            
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'f')
            time.sleep(1)
            pyautogui.press('end')
            time.sleep(1)
            for _ in range(5):
                pyautogui.press('backspace')
                time.sleep(0.3)

            # ---------- Substitua apenas a função get_visible_index por esta versão -----------
            def get_visible_index(shopping, fase_val):
                valid_fases = list(range(1, fase_val + 1))
                return len(valid_fases)

            visible_index_current = get_visible_index(shopping, current_phase) if current_phase is not None else 0
            visible_index_target = get_visible_index(shopping, fase)
            diff = visible_index_target - visible_index_current

            # Em vez de mover a seleção com setas, identificamos a fase e digitamos o nome dela
            # Se a fase for menor que 10, formata com dois dígitos (ex: "03 -"); caso contrário, mantém o número
            formatted_phase = f"{fase:02d} -"
            pyautogui.typewrite(formatted_phase)
            pyautogui.press('enter')
            time.sleep(2)
            logging.info(f"Selecionando a fase: {formatted_phase} ")
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

                combined_extracted = (extracted1 + " " + extracted2) if extracted1 and extracted2 else (extracted1 or extracted2)
                screen_width, screen_height = pyautogui.size()
                center_x = screen_width // 2
                center_y = screen_height // 2
                if fuzzy_contains(combined_extracted, "Alerta VSSC"):
                    logging.info("execute_vsloader: Print indica ação ENTER. Texto identificado: %s", combined_extracted)
                    pyautogui.moveTo(center_x, center_y)
                    pyautogui.click()
                    pyautogui.press("enter")
                    break
                else:
                    break
            time.sleep(1)
            pyautogui.press('down')
            time.sleep(1)
            pyautogui.press('up')
            time.sleep(1)
            logging.info("Escrevendo data de emissão")
            current_phase = fase
            time.sleep(2)
            pyautogui.hotkey('alt', 'space')
            time.sleep(0.3)
            pyautogui.press('down')
            time.sleep(0.3)
            pyautogui.press('enter')
            pyautogui.moveRel(93, 220)
            pyautogui.click()
            pyautogui.click()
            time.sleep(0.3)
            hoje = date.today()
            data_hoje_formatada = f"{hoje.day:02d}{hoje.month:02d}{hoje.year:04d}"
            pyautogui.write(data_hoje_formatada)
            time.sleep(2)
            pyautogui.hotkey('alt', 'space')
            time.sleep(0.3)
            pyautogui.press('down')
            time.sleep(0.3)
            pyautogui.press('enter')
            pyautogui.moveRel(93, 195)
            pyautogui.click()
            pyautogui.click()
            logging.info("Escrevendo data de vencimento")

            def is_business_day(d):
                """
                Retorna True se a data 'd' for um dia útil (não sábado, não domingo e não feriado).
                """
                return d.weekday() < 5 and d not in br_holidays

            def adjust_to_business_day(d):
                """
                Se a data 'd' cair em fim de semana ou feriado, avança para o próximo dia útil.
                """
                while not is_business_day(d):
                    d += timedelta(days=1)
                return d

            def get_nth_business_day(year, month, n):
                """
                Retorna o n-ésimo dia útil do mês e ano informados.
                Caso não haja n dias úteis no mês, continua a contagem no mês seguinte.
                """
                count = 0
                day = 1
                last_day = calendar.monthrange(year, month)[1]
                while day <= last_day:
                    current_date = date(year, month, day)
                    if is_business_day(current_date):
                        count += 1
                        if count == n:
                            return current_date
                    day += 1
                d = date(year, month, last_day) + timedelta(days=1)
                while count < n:
                    if is_business_day(d):
                        count += 1
                        if count == n:
                            return d
                    d += timedelta(days=1)
                return d

            def calculate_due_date(shopping, tipo, fase):
                """
                Calcula a data de vencimento de acordo com as regras:
                
                - Para faturamento Atípicos (todos os shoppings, todas as fases):
                    vencimento é no dia 15 do mês vigente (ajustado para o próximo dia útil se cair em fim de semana/feriado).
                
                - Para faturamento Postecipados:
                    * Se o shopping for "Shopping Mestre Álvaro" e a fase for 22:
                        vencimento é fixo no dia 5 do mês seguinte (ajustado se necessário).
                    * Se o shopping for "Shopping Metrópole" e a fase for 36:
                        vencimento é o 10º dia útil do mês seguinte.
                
                - Para faturamento Antecipados:
                    * Se o shopping for "Shopping Mestre Álvaro" (todas as fases):
                        vencimento é fixo no dia 5 do mês seguinte (ajustado se necessário).
                    * Se o shopping for "Shopping Metrópole" e a fase for 18:
                        vencimento é fixo no dia 20 do mês seguinte (ajustado se necessário).
                
                - Caso nenhuma regra específica seja atendida, o vencimento padrão será o primeiro dia útil do mês seguinte.
                """
                today = date.today()
                if tipo == "Atípicos":
                    d = date(today.year, today.month, 15)
                    return adjust_to_business_day(d)
                if today.month == 12:
                    next_year = today.year + 1
                    next_month = 1
                else:
                    next_year = today.year
                    next_month = today.month + 1
                if shopping == "Shopping Mestre Álvaro" and tipo == "Postecipados" and fase == 41:
                    d = date(next_year, next_month, 5)
                    return adjust_to_business_day(d)
                if shopping == "Shopping Metrópole" and tipo == "Postecipados" and fase == 36:
                    return get_nth_business_day(next_year, next_month, 10)
                if shopping == "Shopping Mestre Álvaro" and tipo == "Antecipados":
                    d = date(next_year, next_month, 5)
                    return adjust_to_business_day(d)
                if shopping == "Shopping Metrópole" and tipo == "Antecipados" and fase == 18:
                    d = date(next_year, next_month, 20)
                    return adjust_to_business_day(d)
                d = date(next_year, next_month, 1)
                return adjust_to_business_day(d)

            due_date = calculate_due_date(shopping, tipo, fase)
            data_formatada = f"{due_date.day:02d}{due_date.month:02d}{due_date.year:04d}"
            pyautogui.write(data_formatada)
            logging.info("Clicando para selecionar LUC's")
            time.sleep(2)
            pyautogui.hotkey('alt', 'space')
            time.sleep(0.3)
            pyautogui.press('down')
            time.sleep(0.3)
            pyautogui.press('enter')
            pyautogui.moveRel(-80, 195)
            pyautogui.click()
            pyautogui.click()
            logging.info("Clicando para selecionar LUC's")

            prev_handle = get_foreground_window()
            wait_for_focus_change(prev_handle)
            wait_for_stable_focus(prev_handle)
            time.sleep(2)
            
            # pyautogui.hotkey('alt', 'space')
            # time.sleep(0.3)
            # pyautogui.press('down')
            # time.sleep(0.3)
            # pyautogui.press('enter')
            # pyautogui.moveRel(-150, 365)
            # pyautogui.click()
            # pyautogui.click()
            # time.sleep(1)
            # for _ in range(4):
            #     pyautogui.press('tab')
            #     time.sleep(0.1)
            
           

            # time.sleep(1)
            # pyautogui.hotkey('ctrl', 'f')
            # time.sleep(1)
            # company_map = {
            #     "SMT": "Machida",
            #     "SRP": "Bodytech",
            #     "SDI": "Bodytech",
            #     "SMO": "Formula",
            #     "SMA": "Formula",
            #     "SMS": "BOB S",
            # }

            # # determine_variant(shopping) já retorna a abreviação (“SMT”, “SRP” etc.)
            # variant = determine_variant(shopping)
            # texto_para_digitar = company_map.get(variant, "Bodytech")

            # pyautogui.typewrite(texto_para_digitar)
            # pyautogui.press("enter")
            # time.sleep(0.2)
            # pyautogui.press('enter')
            # time.sleep(0.5)
            # for _ in range(2):
            #     pyautogui.press('left')
            #     time.sleep(0.1)
            # pyautogui.press('enter')
            # time.sleep(0.5)
            # for _ in range(4):
            #     pyautogui.press('tab')
            #     time.sleep(0.1)
            pyautogui.hotkey('alt', 'space')
            time.sleep(0.3)
            pyautogui.press('down')
            time.sleep(0.3)
            pyautogui.press('enter')
            pyautogui.moveRel(-220, 365)
            pyautogui.click()
            pyautogui.click()
            logging.info("Selecionando todas as empresas")
            time.sleep(0.5)
            for _ in range(3):
                pyautogui.press('tab')
                time.sleep(0.1)
            pyautogui.press('enter')
            time.sleep(1)
            pyautogui.hotkey('alt', 'space')
            time.sleep(0.3)
            pyautogui.press('down')
            time.sleep(0.3)
            pyautogui.press('enter')
            pyautogui.moveRel(180, 300)
            pyautogui.click()
            pyautogui.click()
            logging.info("Clicando no formato do email gerado")
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(0.5)
            for _ in range(6):
                pyautogui.press('tab')
                time.sleep(0.3)
            pyautogui.press('enter')
            time.sleep(0.3)
            pyautogui.press('enter')
            time.sleep(0.3)
            pyautogui.press('tab')
            time.sleep(0.3)
            pyautogui.press('enter')
            time.sleep(2)
            
            pyautogui.press('tab')
            time.sleep(2)
            pyautogui.hotkey('alt', 'space')
            time.sleep(0.3)
            pyautogui.press('down')
            time.sleep(0.3)
            pyautogui.press('enter')
            pyautogui.moveRel(130, 111)
            pyautogui.click()
            pyautogui.click()
            logging.info("Clicando para selecionar o envio de e-mail")
            pyautogui.press('tab')
            pyautogui.press('enter')
            time.sleep(0.3)
            pyautogui.press('enter')
            time.sleep(2)

            logging.info(f"Enviando e-mail.")

            global IS_SEGURO
            # Unifica a validação por prints em um único while; para cada verificação, realiza 2 prints com 2 segundos de intervalo
            while True:
                screenshot1 = capture_screenshot()
                extracted1 = analyze_screenshot(screenshot1)
                time.sleep(2)
                screenshot2 = capture_screenshot()
                extracted2 = analyze_screenshot(screenshot2)

                if not extracted1 or not extracted2:
                    time.sleep(8)
                    continue

                # Primeiro, trate o caso de None convertendo para "" (string vazia):
                extracted1_l = extracted1.lower() if extracted1 else ""
                extracted2_l = extracted2.lower() if extracted2 else ""

                # Agora, o if fica assim:
                # if (
                #     (
                #         ("não há nenhum modal" in extracted1_l)
                #         or ("lobby" in extracted1_l)
                #         or ("modal não detectado" in extracted1_l)
                #     )
                #     and (
                #         ("não há nenhum modal" in extracted2_l)
                #         or ("lobby" in extracted2_l)
                #         or ("modal não detectado" in extracted2_l)
                #     )
                # ):
                #     logging.info(
                #         "execute_vsloader: Lobby identificado (nenhum modal detectado) em ambos os prints. "
                #         "Texto 1: %s Texto 2: %s", extracted1, extracted2
                #     )
                #     break


                combined_extracted = (extracted1 + " " + extracted2) if extracted1 and extracted2 else (extracted1 or extracted2)
                screen_width, screen_height = pyautogui.size()
                center_x = screen_width // 2
                center_y = screen_height // 2
                if fuzzy_contains(combined_extracted, "Alerta VSSC") and fuzzy_contains(combined_extracted, "A área de recibo já foi gerada"):
                    logging.info("execute_vsloader: Alerta VSSC. Texto identificado: %s", combined_extracted)
                    pyautogui.moveTo(center_x, center_y)
                    pyautogui.click()
                    pyautogui.press("enter")
                elif fuzzy_contains(combined_extracted, "Pressione <ESC>"):
                    logging.info("execute_vsloader: Print indica ação ESC. Texto identificado: %s", combined_extracted)
                    pyautogui.moveTo(center_x, center_y)
                    pyautogui.click()
                    time.sleep(0.5)
                    pyautogui.press("esc")
                elif fuzzy_contains(combined_extracted, "Processando código"):
                    logging.info("execute_vsloader: Processando código. Texto identificado: %s", combined_extracted)
                    time.sleep(15)
                elif fuzzy_contains(combined_extracted, "Lista de Recibos Já Emitidos"):
                    logging.info("execute_vsloader: Lista de recibos já emitidos. Texto identificado: %s", combined_extracted)
                    pyautogui.moveTo(center_x, center_y)
                    pyautogui.click()
                    pyautogui.hotkey('alt', 'space')
                    time.sleep(0.3)
                    pyautogui.press('down')
                    time.sleep(0.3)
                    pyautogui.press('enter')
                    pyautogui.moveRel(-370, 370)
                    pyautogui.click()
                    pyautogui.click()
                    for _ in range(2):
                        pyautogui.press('tab')
                        time.sleep(0.3)
                    pyautogui.press('enter')
                elif fuzzy_contains(combined_extracted, "Emitindo Recibos"):
                    logging.info("execute_vsloader: Emitindo recibos. Texto identificado: %s", combined_extracted)
                    time.sleep(8)
                
                elif fuzzy_contains(combined_extracted, "Alerta VSSC"):
                    logging.info("execute_vsloader: Print indica ação ESC. Texto identificado: %s", combined_extracted) 
                    pyautogui.moveTo(center_x, center_y)
                    pyautogui.click()
                    time.sleep(0.5)
                    pyautogui.press("esc")
                
                    
                elif fuzzy_contains(combined_extracted, "ATENÇÃO") and fuzzy_contains(combined_extracted, "Competência de Trabalho será alterada"):
                    logging.info("execute_vsloader: Competência de Trabalho será alterada. Texto identificado: %s", combined_extracted)
                    pyautogui.moveTo(center_x, center_y)
                    pyautogui.click()
                    pyautogui.press("enter")
                    time.sleep(3)
                    break
                
                elif fuzzy_contains(combined_extracted, "Emissão do Recibo") and fuzzy_contains(combined_extracted, "competência"):
                    logging.info("execute_vsloader: Tela de Emissão de Recibo identificada, desativando validação. Texto identificado: %s", combined_extracted)
                    pyautogui.press("esc")
                    time.sleep(3)
                    break
                    
                
                elif fuzzy_contains(combined_extracted, "Pessoas Físicas e Jurídicas") \
                    or fuzzy_contains(combined_extracted, "Corpo do email") \
                    or fuzzy_contains(combined_extracted, "Corpo do e mail") \
                    or fuzzy_contains(combined_extracted, "Corpo do e-mail"):
                    logging.info("execute_vsloader: Corpo do email. Texto identificado: %s", combined_extracted)
                    pyautogui.moveTo(center_x, center_y)
                    pyautogui.click()
                    time.sleep(3)
                    break
                    
                elif fuzzy_contains(combined_extracted, "Contratos com término"):
                    logging.info("execute_vsloader: Contratos com término, desativando validação. Texto identificado: %s", combined_extracted)
                    time.sleep(3)
                    break
                
                elif (fuzzy_contains(combined_extracted, "Competência de trabalho:") and 
                      fuzzy_contains(combined_extracted, "Período Fechado") and 
                      fuzzy_contains(combined_extracted, "(Faturamento)")):
                    logging.info("execute_vsloader: Lobby identificado. Texto identificado: %s", combined_extracted)
                    time.sleep(3)
                    break
                else:
                    logging.info("execute_vsloader: Nenhuma condição modal identificada. Texto identificado: %s", combined_extracted)
                time.sleep(3)
        
            pyautogui.moveTo(center_x, center_y)
            pyautogui.click()

            pyautogui.hotkey('alt', 'space')
            time.sleep(0.3)
            pyautogui.press('down')
            time.sleep(0.3)
            pyautogui.press('enter')
            pyautogui.moveRel(172, 60)
            pyautogui.click()
            pyautogui.click()
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'f')
            time.sleep(1)
            pyautogui.press('end')
            time.sleep(1)
            for _ in range(5):
                pyautogui.press('backspace')
                time.sleep(0.3)

            # ---------- Substitua apenas a função get_visible_index por esta versão -----------
            def get_visible_index(shopping, fase_val):
                valid_fases = list(range(1, fase_val + 1))
                return len(valid_fases)

            visible_index_current = get_visible_index(shopping, current_phase) if current_phase is not None else 0
            visible_index_target = get_visible_index(shopping, fase)
            diff = visible_index_target - visible_index_current

            # Em vez de mover a seleção com setas, identificamos a fase e digitamos o nome dela
            # Se a fase for menor que 10, formata com dois dígitos (ex: "03 -"); caso contrário, mantém o número
            formatted_phase = f"{fase:02d} -"
            pyautogui.typewrite(formatted_phase)
            pyautogui.press('enter')
            

            time.sleep(3)
            pyautogui.press('down')
            time.sleep(2)

            while True:
            # Em vez de apagar os prints imediatamente, capturamos e acumulamos-os.
                screenshot1 = capture_screenshot()
                extracted1 = analyze_screenshot(screenshot1)
                time.sleep(2)
                screenshot2 = capture_screenshot()
                extracted2 = analyze_screenshot(screenshot2)
                if extracted1 is None and extracted2 is None:
                    time.sleep(3)
                    continue
                extracted1_l = extracted1.lower() if extracted1 else ""
                extracted2_l = extracted2.lower() if extracted2 else ""

                # Agora, o if fica assim:
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
                screen_width, screen_height = pyautogui.size()
                center_x = screen_width // 2
                center_y = screen_height // 2
                if fuzzy_contains(combined_extracted, "Alerta VSSC"):
                    logging.info("execute_vsloader: Não há emissão para essa fase. Texto identificado: %s", combined_extracted)
                    pyautogui.moveTo(center_x, center_y)
                    pyautogui.click()
                    time.sleep(0.5)
                    pyautogui.press('esc')
                    time.sleep(2)
                    break
                else:
                    logging.info("execute_vsloader: não identificou o alerta. Texto identificado: %s", combined_extracted)
                    break
                time.sleep(3)

            
            time.sleep(1)
            pyautogui.moveTo(center_x, center_y)
            pyautogui.click()
            time.sleep(0.3)
            pyautogui.hotkey('alt', 'space')
            time.sleep(0.3)
            pyautogui.press('down')
            time.sleep(0.3)
            pyautogui.press('enter')
            pyautogui.moveRel(172, 60)
            pyautogui.click()
            pyautogui.click()
            pyautogui.press('up')
            logging.info("Clicando no up")

            while True:
            # Em vez de apagar os prints imediatamente, capturamos e acumulamos-os.
                screenshot1 = capture_screenshot()
                extracted1 = analyze_screenshot(screenshot1)
                time.sleep(2)
                screenshot2 = capture_screenshot()
                extracted2 = analyze_screenshot(screenshot2)
                if extracted1 is None and extracted2 is None:
                    time.sleep(3)
                    continue
                extracted1_l = extracted1.lower() if extracted1 else ""
                extracted2_l = extracted2.lower() if extracted2 else ""

                # Agora, o if fica assim:
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
                screen_width, screen_height = pyautogui.size()
                center_x = screen_width // 2
                center_y = screen_height // 2
                if fuzzy_contains(combined_extracted, "Alerta VSSC") and fuzzy_contains(combined_extracted, "Não há emissão"):
                    logging.info("execute_vsloader: Não há emissão para essa fase. Texto identificado: %s", combined_extracted)
                    time.sleep(1)
                    for _ in range(5):
                        pyautogui.press('esc')
                        time.sleep(0.3)
                    if tipo == "Antecipados":
                        base = 0
                    elif tipo == "Postecipados":
                        base = 3
                    elif tipo == "Atípicos":
                        base = 6
                    else:
                        base = 0
                    updated = False
                   
                    break

                
                else:
                    while True:
                        screenshot1 = capture_screenshot()
                        extracted1 = analyze_screenshot(screenshot1)
                        time.sleep(2)
                        screenshot2 = capture_screenshot()
                        extracted2 = analyze_screenshot(screenshot2)

                        if not extracted1 or not extracted2:
                            time.sleep(8)
                            continue

                        # Primeiro, trate o caso de None convertendo para "" (string vazia):
                        extracted1_l = extracted1.lower() if extracted1 else ""
                        extracted2_l = extracted2.lower() if extracted2 else ""

                        # Agora, o if fica assim:
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
                        screen_width, screen_height = pyautogui.size()
                        center_x = screen_width // 2
                        center_y = screen_height // 2
                        if fuzzy_contains(combined_extracted, "Alerta VSSC") and fuzzy_contains(combined_extracted, "Não há emissão"):
                            logging.info("execute_vsloader: Não há emissão para essa fase. Texto identificado: %s", combined_extracted)
                            
                            time.sleep(1)
                            for _ in range(5):
                                pyautogui.press('esc')
                                time.sleep(0.3)
                            break
                        elif fuzzy_contains(combined_extracted, "Alerta VSSC") :
                            pyautogui.press('esc')
                            time.sleep(0.3)
                        
                        elif not fuzzy_contains(combined_extracted, "Alerta VSSC") and not fuzzy_contains(combined_extracted, "Não há emissão"):
                            logging.info("execute_vsloader: Alerta VSSC Texto identificado: %s", combined_extracted)
                            time.sleep(3)
                            pyautogui.press('enter')
                            time.sleep(12)
                            pyautogui.hotkey('alt', 'space')
                            time.sleep(0.3)
                            pyautogui.press('down')
                            time.sleep(0.3)
                            pyautogui.press('enter')
                            pyautogui.moveRel(170,145)
                            pyautogui.click()
                            pyautogui.click()
                            logging.info("Selecionando o formato do email")
                            time.sleep(1)
                            pyautogui.hotkey('ctrl', 'f')
                            time.sleep(1)
                            pyautogui.write('textoemail.html')
                            time.sleep(0.5)
                            pyautogui.press('enter')
                            time.sleep(0.5)
                            pyautogui.press('enter')
                            time.sleep(2)
                            pyautogui.hotkey('alt', 'space')
                            time.sleep(0.3)
                            pyautogui.press('down')
                            time.sleep(0.3)
                            pyautogui.press('enter')
                            pyautogui.moveRel(145,560)
                            pyautogui.click()
                            pyautogui.click()
                            logging.info("Selecionando o tipo do email (última opção)")
                            time.sleep(1)
                            for _ in range(5):
                                pyautogui.press('down')
                                time.sleep(0.3) 
                            time.sleep(5)               
                            pyautogui.press('enter')
                            time.sleep(1)
                            pyautogui.press('tab')
                            pyautogui.press('enter')
                            break
                 
                        else:
                            logging.info("execute_vsloader: Nenhuma condição modal identificada. Texto identificado: %s", combined_extracted)
                        time.sleep(3)
                    
                    time.sleep(3)
                    while True:
                        screenshot1 = capture_screenshot()
                        extracted1 = analyze_screenshot(screenshot1)
                        time.sleep(2)
                        screenshot2 = capture_screenshot()
                        extracted2 = analyze_screenshot(screenshot2)

                        if not extracted1 or not extracted2:
                            time.sleep(8)
                            continue

                        extracted1 = extracted1.lower() if isinstance(extracted1, str) else ""
                        extracted2 = extracted2.lower() if isinstance(extracted2, str) else ""

                        if (("não há nenhum modal" in extracted1 or "lobby" in extracted1 or "modal não detectado" in extracted1) and
                            ("não há nenhum modal" in extracted2 or "lobby" in extracted2 or "modal não detectado" in extracted2)):
                            logging.info("execute_vsloader: Lobby identificado (nenhum modal detectado) em ambos os prints.")
                            time.sleep(3)
                            break

                        combined_extracted = (extracted1 + " " + extracted2) if extracted1 and extracted2 else (extracted1 or extracted2)
                        screen_width, screen_height = pyautogui.size()
                        center_x = screen_width // 2
                        center_y = screen_height // 2
                        if fuzzy_contains(combined_extracted, "Processando código"):
                            logging.info("execute_vsloader: 1111 Processando código. Texto identificado: %s", combined_extracted)
                            time.sleep(15)
                            
                        elif fuzzy_contains(combined_extracted, "Alerta VSSC") and fuzzy_contains(combined_extracted, "A área de recibo já foi gerada"):
                            logging.info("execute_vsloader: 1115 Alerta VSSC. Texto identificado: %s", combined_extracted)
                            pyautogui.moveTo(center_x, center_y)
                            pyautogui.click()
                            pyautogui.press("enter")
                        
                    
                        elif fuzzy_contains(combined_extracted, "envio") \
                            and fuzzy_contains(combined_extracted, "boleto") \
                            and fuzzy_contains(combined_extracted, "concluido") \
                            and fuzzy_contains(combined_extracted, "alerta vssc"):

                            logging.info("execute_vsloader: Envio de boletos por email concluído. Texto identificado: %s", combined_extracted)
                            pyautogui.moveTo(center_x, center_y)
                            pyautogui.click()
                            for _ in range(6):
                                
                                time.sleep(0.5)
                                pyautogui.press('esc')
                                time.sleep(0.3)
                            if tipo == "Antecipados":
                                base = 0
                            elif tipo == "Postecipados":
                                base = 3
                            elif tipo == "Atípicos":
                                base = 6
                            else:
                                base = 0
                            updated = False
                            
                            break

                        
                        else:
                            logging.info("execute_vsloader: Nenhuma condição modal identificada. Texto identificado: %s", combined_extracted)
                        time.sleep(3)
                    break
                
                    
            for _ in range(5):
                pyautogui.moveTo(center_x, center_y)
                pyautogui.click()
                pyautogui.press('esc') 
                time.sleep(0.3)
            time.sleep(3)
            
            

            
        for _ in range(5):
            pyautogui.press('esc')
            time.sleep(0.3)

#################### CHAMADAS #####################

        

   #########################################
        # Ao final da execução, deleta todos os prints acumulados e encerra o aplicativo
        delete_all_prints()
        time.sleep(2)
        pyautogui.hotkey('alt', 'F4')
        time.sleep(2)
        pyautogui.moveTo(center_x, center_y)
        pyautogui.click()
        pyautogui.hotkey('alt', 'F4')

    except ElementNotFoundError as e:
        logging.error(f"Elemento não encontrado: {e}")
    except TimeoutError as e:
        logging.error(f"Tempo limite excedido: {e}")
    except Exception as e:
        logging.error(f"Erro inesperado: {e}")

if __name__ == "__main__":
    import sys
    shopping_escolhido = sys.argv[1]
    tipo_escolhido = sys.argv[2]

    # Abreviação dos nomes dos shoppings
    abreviacoes = {
        "Shopping da Ilha": "SDI",
        "Shopping Moxuara": "SMO",
        "Shopping Mestre Álvaro": "SMA",
        "Shopping Montserrat": "SMS",
        "Shopping Rio Poty": "SRP",
        "Shopping Metrópole": "SMT"
    }
    shopping_abreviado = abreviacoes.get(shopping_escolhido, shopping_escolhido.replace(' ', '_').upper())

    # --- cria pasta de logs para este shopping ---
        # --- cria pasta de logs para este shopping ---
    log_dir = os.path.join(os.getcwd(), "logs", shopping_abreviado)
    os.makedirs(log_dir, exist_ok=True)

    # Cria o nome do arquivo de log dentro de logs/<abreviatura>/
    log_filename = os.path.join(log_dir, f"Hom_enviar_email_{shopping_abreviado}_{tipo_escolhido}.txt")

    # Se já existir um log com mesmo nome, remove para não acumular
    if os.path.exists(log_filename):
        os.remove(log_filename)

    # Remove os handlers já configurados
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)

    logging.basicConfig(
        filename=log_filename,
        level=logging.INFO,
        format=f"({shopping_escolhido} | {tipo_escolhido}) %(asctime)s %(levelname)s: %(message)s",
        datefmt='%d/%m/%Y %H:%M:%S'
    )

    execute_vsloader(shopping_escolhido, tipo_escolhido)
