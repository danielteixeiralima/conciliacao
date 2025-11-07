# -*- coding: utf-8 -*-

###############################################################################
#                              teste.py                                       #
###############################################################################

import time
import logging
import pyautogui
import os
from statistics import mean

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.1

def _rect_from_string(s):
    """
    Converte 'L,T,R,B' (string) para tuple de ints (L, T, R, B).
    """
    parts = [p.strip() for p in s.split(",")]
    if len(parts) != 4:
        raise ValueError("rect deve ser 'L,T,R,B'")
    return tuple(int(x) for x in parts)

def _center_of_rect(rect):
    l, t, r, b = rect
    return ((l + r) // 2, (t + b) // 2)

def _click_and_type_rect(rect, text):
    """
    Move o mouse para o centro do rect e digita o texto.
    Usa pequenos delays para estabilidade.
    """
    cx, cy = _center_of_rect(rect)
    pyautogui.moveTo(cx, cy, duration=0.15)
    pyautogui.click()
    time.sleep(0.12)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.05)
    pyautogui.press('backspace')
    time.sleep(0.05)
    pyautogui.typewrite(text, interval=0.03)

def _avg_pixel_brightness(img):
    """
    Recebe uma PIL Image e retorna a média de brilho (0..255)
    calculada a partir dos canais RGB.
    """
    pixels = list(img.getdata())
    if not pixels:
        return 0
    br = [mean(px[:3]) for px in pixels]
    return mean(br)

def wait_and_click_rect_when_visible(rect_str, timeout=60, check_interval=3, brightness_diff_threshold=10):
    """
    Aguarda até que a região definida por rect_str (L,T,R,B) apresente
    diferença visível (brilho) em relação ao estado inicial e então clica
    no centro. Retorna True se clicou, False em timeout.
    - timeout em segundos
    - check_interval em segundos entre tentativas
    - brightness_diff_threshold: diferença mínima de brilho para considerar "visível"
    """
    rect = _rect_from_string(rect_str)
    l, t, r, b = rect
    w = r - l
    h = b - t
    if w <= 0 or h <= 0:
        logging.error("wait_and_click_rect_when_visible: rect inválido %s", rect_str)
        return False

    try:
        baseline = pyautogui.screenshot(region=(l, t, w, h))
        baseline_b = _avg_pixel_brightness(baseline)
    except Exception as e:
        logging.error("Erro ao capturar baseline da região: %s", e)
        baseline_b = None

    start = time.time()
    while time.time() - start < timeout:
        try:
            cur = pyautogui.screenshot(region=(l, t, w, h))
            cur_b = _avg_pixel_brightness(cur)
            if baseline_b is None:
                if cur_b > 1:
                    cx, cy = _center_of_rect(rect)
                    pyautogui.moveTo(cx, cy, duration=0.12)
                    pyautogui.click()
                    logging.info("Região detectada (sem baseline) e clicada em %s", (cx, cy))
                    return True
            else:
                diff = abs(cur_b - baseline_b)
                logging.debug("Brilho baseline=%.2f atual=%.2f diff=%.2f", baseline_b, cur_b, diff)
                if diff >= brightness_diff_threshold:
                    cx, cy = _center_of_rect(rect)
                    pyautogui.moveTo(cx, cy, duration=0.12)
                    pyautogui.click()
                    logging.info("Região detectada (diff=%.2f) e clicada em %s", diff, (cx, cy))
                    return True
        except Exception as e:
            logging.debug("Erro ao capturar/avaliar região: %s", e)

        time.sleep(check_interval)

    logging.warning("Timeout aguardando região %s ficar disponível.", rect_str)
    return False

def wait_for_window_title_appears(title_substr, timeout=60, check_interval=2):
    """
    Método simples que aguarda a aparição de um título de janela na tela
    consultando pyautogui.getWindowsWithTitle (se disponível).
    Retorna True se apareceu, False em timeout.
    """
    start = time.time()
    while time.time() - start < timeout:
        try:
            wins = pyautogui.getWindowsWithTitle(title_substr)
            if wins:
                return True
        except Exception:
            pass
        time.sleep(check_interval)
    return False

def fill_credentials(username, password, user_rect, pass_rect):
    """
    Usa pyautogui puro para clicar nos campos e preencher usuário e senha.
    """
    try:
        rect_u = _rect_from_string(user_rect)
        rect_p = _rect_from_string(pass_rect)

        _click_and_type_rect(rect_u, username)
        time.sleep(0.5)
        _click_and_type_rect(rect_p, password)
        time.sleep(0.5)

        pyautogui.press('enter')
        logging.info("Credenciais digitadas e ENTER pressionado.")
        return True
    except Exception as e:
        logging.error(f"Erro ao preencher credenciais: {e}")
        return False

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s:%(message)s")

    # ===================================================================
    # CONFIGURAÇÕES FIXAS (específicas do cliente) - defina abaixo
    # ===================================================================
    folder = r"C:\Program Files\Victor & Schellenberger\VSSC_ILHA"
    username = "z8"
    password = "S@cavalcante"
    user_rect = "1035,652,1051,668"      # coordenadas do campo de USUÁRIO (L,T,R,B)
    pass_rect = "1035,670,1122,686"      # coordenadas do campo de SENHA  (L,T,R,B)

    # Retângulo do botão "Ciente" (capturado via Inspect em uma execução anterior).
    # Se esse botão aparecer até 5s após o ENTER, será clicado; senão o robo segue.
    ciente_rect = "1246,579,1319,603"

    # Elemento que você quer clicar APÓS o login (ex.: um menu item detectado pelo Inspect)
    # Use o rect que o Inspect mostrou (L,T,R,B). Ajuste se necessário.
    post_login_click_rect = "1220,138,1278,157"

    # ===================================================================
    # Fluxo principal
    # ===================================================================

    # abrir explorer e navegar até a pasta
    pyautogui.press("win")
    time.sleep(2)
    pyautogui.write("file explorer")
    pyautogui.press("enter")
    time.sleep(6)

    pyautogui.hotkey("alt", "d")
    pyautogui.typewrite(folder)
    pyautogui.press("enter")
    time.sleep(3)

    pyautogui.typewrite("VSLOADER.EXE")
    pyautogui.press("enter")

    logging.info("VSLOADER.EXE iniciado, aguardando janela...")

    # Em vez de um sleep fixo longo, aguardamos a janela do VSLoader surgir por um tempo
    found_title = wait_for_window_title_appears("VSLoader", timeout=30, check_interval=2)
    if not found_title:
        logging.info("Não detectei título 'VSLoader' (ou método indisponível). Aguardando 8s extra.")
        time.sleep(8)
    else:
        logging.info("Janela com 'VSLoader' detectada.")

    logging.info("Tentando preencher credenciais...")

    ok = fill_credentials(username, password, user_rect, pass_rect)

    if not ok:
        logging.error("Falha ao preencher credenciais automaticamente.")
    else:
        logging.info("Credenciais enviadas com sucesso. Aguardando possível diálogo/aviso 'Ciente' por até 5s...")

        # tenta detectar e clicar no botão "Ciente" por até 5 segundos; se não aparecer, segue o fluxo.
        clicked_ciente = wait_and_click_rect_when_visible(ciente_rect, timeout=5, check_interval=0.5, brightness_diff_threshold=6)
        if clicked_ciente:
            logging.info("Botão 'Ciente' detectado e clicado.")
            # dá um pequeno tempo após clicar no alerta para a aplicação estabilizar
            time.sleep(1.0)
        else:
            logging.info("Botão 'Ciente' não apareceu dentro de 5s. Continuando execução.")

        logging.info("Agora aguardando e clicando no elemento pós-login (se necessário)...")

        # aguardar e clicar no elemento pós-login apenas quando disponível (timeout 60s)
        clicked_post = wait_and_click_rect_when_visible(post_login_click_rect, timeout=60, check_interval=3, brightness_diff_threshold=8)
        if clicked_post:
            logging.info("Elemento pós-login detectado e clicado com sucesso.")
        else:
            logging.error("Não foi possível detectar/acionar o elemento pós-login dentro do timeout.")
