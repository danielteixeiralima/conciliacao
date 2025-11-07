# FILE: inspector_pywinauto.py
# -*- coding: utf-8 -*-
"""
Ferramenta auxiliar para inspecionar janelas e controles com pywinauto (backend UIA).
- Lista janelas top-level (índice, título, pid).
- Permite imprimir control_identifiers() de uma janela escolhida e salvar em arquivo.
- Permite capturar informações do elemento sob o cursor (nome, automation_id, control_type, class_name, rectangle).

Uso:
  python inspector_pywinauto.py
(O script é interativo para facilitar uso direto)
"""

import sys
import time
from datetime import datetime
from pywinauto import Desktop
from pywinauto import Application
from pywinauto.uia_element_info import UIAElementInfo

def list_windows():
    desktop = Desktop(backend="uia")
    wins = desktop.windows()
    listing = []
    for i, w in enumerate(wins, start=1):
        try:
            title = w.window_text()
        except Exception:
            title = "<sem título>"
        try:
            pid = w.process_id()
        except Exception:
            pid = None
        print(f"[{i}] PID={pid} Título='{title}'")
        listing.append((i, pid, title, w))
    return listing

def print_and_save_window_identifiers(window_wrapper):
    """
    Chama print_control_identifiers() e salva a versão textual em arquivo para análise.
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"control_identifiers_{ts}.txt"
    try:
        # redirect print to file by capturing the output
        with open(filename, "w", encoding="utf-8") as f:
            window_wrapper.print_control_identifiers(indent=2, file=f)
        print(f"Arquivo salvo: {filename}")
    except Exception as e:
        print(f"Erro ao salvar control_identifiers: {e}")

def inspect_element_under_cursor():
    """
    Retorna dados do elemento UIA sob o cursor usando UIAElementInfo.from_point.
    """
    try:
        import ctypes
        user32 = ctypes.windll.user32
        pt = ctypes.wintypes.POINT()
        user32.GetCursorPos(ctypes.byref(pt))
        x, y = pt.x, pt.y
    except Exception:
        from pywinauto.mouse import get_position
        x, y = get_position()

    try:
        ei = UIAElementInfo.from_point(x, y)  # <=== CORRIGIDO (antes passava uma tupla)
        info = {
            "name": ei.name,
            "automation_id": ei.automation_id,
            "control_type": ei.control_type,
            "class_name": ei.class_name,
            "rectangle": ei.rectangle,
            "process_id": ei.process_id,
        }
        print("Elemento sob cursor:")
        for k, v in info.items():
            print(f"  {k}: {v}")
    except Exception as e:
        print(f"Erro ao obter elemento sob cursor: {e}")


def interactive_loop():
    print("=== Pywinauto Inspector (UIA) ===")
    while True:
        print("\nOpções:")
        print("  1 - Listar janelas top-level")
        print("  2 - Imprimir control_identifiers() de uma janela (e salvar em arquivo)")
        print("  3 - Inspecionar elemento sob o cursor")
        print("  4 - Sair")
        choice = input("Escolha: ").strip()
        if choice == "1":
            list_windows()
        elif choice == "2":
            listing = list_windows()
            idx = input("Digite o índice da janela a inspecionar: ").strip()
            try:
                idx_i = int(idx)
                found = None
                for item in listing:
                    if item[0] == idx_i:
                        found = item
                        break
                if not found:
                    print("Índice inválido.")
                    continue
                wrapper = found[3]
                print(f"Imprimindo control_identifiers() para janela '{found[2]}' (PID={found[1]})")
                # salva em arquivo
                print_and_save_window_identifiers(wrapper)
            except Exception as e:
                print(f"Erro: {e}")
        elif choice == "3":
            print("Posicione o cursor sobre o elemento desejado e pressione Enter...")
            input("Enter pra capturar posição atual do cursor...")
            inspect_element_under_cursor()
        elif choice == "4":
            print("Saindo.")
            break
        else:
            print("Opção inválida. Tente novamente.")

if __name__ == "__main__":
    interactive_loop()
