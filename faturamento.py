# -*- coding: utf-8 -*-
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk
import logging

#### FASES tipo 1



# Mapeamento dos shoppings para suas pastas (removido o "Shopping Praia da Costa")
folder_map = {
    "Shopping Montserrat": r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_MONTSERRAT_HOM",
    "Shopping da Ilha": r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_ILHA_HOM",
    "Shopping Rio Poty": r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_TERESINA_HOM",
    "Shopping Metrópole": r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_METROPOLE_HOM",
    "Shopping Moxuara": r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_MOXUARA_HOM",
    "Shopping Mestre Álvaro": r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_MESTREALVARO_HOM"
}


logging.info("Executando:", os.path.abspath(__file__))

# Tipos de faturamento disponíveis
tipos_faturamento = ["Postecipados","Antecipados", "Atípicos",]

# Scripts disponíveis para execução
scripts = {
    "Gerar cálculos": r"C:\AUTOMACAO\faturamento\bots\calculos.exe",
    "Gerar boletos": r"C:\AUTOMACAO\faturamento\bots\gerar_boletos.exe",
    "Enviar e-mails": r"C:\AUTOMACAO\faturamento\bots\enviar_email.exe"
}


def main():
    # Verifica quantos argumentos foram passados
    if len(sys.argv) == 4:
        shopping = sys.argv[1]
        action = sys.argv[2]
        payment = sys.argv[3]

        # Loga os parâmetros recebidos e o diretório de trabalho atual
        logging.info("Parâmetros recebidos: Shopping: %s, Action: %s, Payment: %s", shopping, action, payment)
        logging.info("Diretório atual: %s", os.getcwd())

        if shopping in folder_map and action in scripts and payment in tipos_faturamento:
            caminho_script = scripts[action]
            # Loga o caminho do executável
            logging.info("Caminho do executável: %s", caminho_script)
            if not os.path.exists(caminho_script):
                logging.error("Executável NÃO encontrado: %s", caminho_script)
            else:
                logging.info("Executável encontrado.")

            if caminho_script.endswith(".py"):
                cmd = f'python "{caminho_script}" "{shopping}" "{payment}"'
            else:
                cmd = f'"{caminho_script}" "{shopping}" "{payment}"'

            # Loga o comando que será executado
            logging.info("Comando a ser executado: %s", cmd)
            try:
                subprocess.Popen(cmd, shell=True)
            except Exception as e:
                logging.error("Erro ao executar o comando: %s", e)
        else:
            print("Parâmetros inválidos ou não reconhecidos.")
            sys.exit(1)
    else:
        criar_interface()

def center_window(window, width, height):
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")

def executar_script(shopping, script, faturamento):
    caminho_script = scripts[script]
    if caminho_script.endswith(".py"):
        cmd = f'python "{caminho_script}" "{shopping}" "{faturamento}"'
    else:
        cmd = f'"{caminho_script}" "{shopping}" "{faturamento}"'
    print("Comando a ser executado:", cmd)

    subprocess.Popen(cmd, shell=True)


def selecionar_tipo_faturamento(shopping, script):
    tipo_window = tk.Toplevel(root)
    tipo_window.title(f"Selecionar Tipo - {shopping}")
    center_window(tipo_window, 500, 300)
    tipo_window.resizable(False, False)
    tipo_window.configure(bg="#e7f0ff")

    main_frame = ttk.Frame(tipo_window, style="Main.TFrame", padding=20)
    main_frame.pack(expand=True, fill='both')

    label = ttk.Label(main_frame, text="Escolha o tipo de faturamento:", font=("Segoe UI", 12))
    label.pack(pady=10)

    for tipo in tipos_faturamento:
        btn = ttk.Button(
            main_frame,
            text=tipo,
            style="Accent.TButton",
            command=lambda t=tipo: [tipo_window.destroy(), executar_script(shopping, script, t)]
        )
        btn.pack(pady=5, fill='x')

def selecionar_script(shopping):
    script_window = tk.Toplevel(root)
    script_window.title(f"Selecionar Ação - {shopping}")
    center_window(script_window, 500, 300)
    script_window.resizable(False, False)
    script_window.configure(bg="#e7f0ff")

    main_frame = ttk.Frame(script_window, style="Main.TFrame", padding=20)
    main_frame.pack(expand=True, fill='both')

    label = ttk.Label(main_frame, text="Escolha a ação desejada:", font=("Segoe UI", 12))
    label.pack(pady=10)

    for acao in scripts.keys():
        btn = ttk.Button(
            main_frame,
            text=acao,
            style="Accent.TButton",
            command=lambda s=acao: [script_window.destroy(), selecionar_tipo_faturamento(shopping, s)]
        )
        btn.pack(pady=5, fill='x')

def criar_interface():
    global root
    root = tk.Tk()
    root.title("Sistema de Automação - Faturamento")
    center_window(root, 500, 400)
    root.resizable(False, False)
    root.configure(bg="#e7f0ff")

    style = ttk.Style(root)
    style.theme_use('clam')

    style.configure('Main.TFrame', background='#e7f0ff')
    style.configure('TLabel', background='#e7f0ff', font=('Segoe UI', 12))
    style.configure('Accent.TButton',
                    font=('Segoe UI', 12),
                    foreground='white',
                    background='#3366CC',
                    padding=6)
    style.map('Accent.TButton',
              background=[('active', '#2853a3'), ('disabled', '#a6a6a6')])

    main_frame = ttk.Frame(root, style="Main.TFrame", padding=20)
    main_frame.pack(expand=True, fill='both')

    label = ttk.Label(main_frame, text="Selecione um Shopping:", font=("Segoe UI", 14, "bold"))
    label.pack(pady=10)

    for shopping in folder_map.keys():
        btn = ttk.Button(
            main_frame,
            text=shopping,
            style="Accent.TButton",
            command=lambda s=shopping: selecionar_script(s)
        )
        btn.pack(pady=5, fill='x')

    root.mainloop()

if __name__ == "__main__":
    main()
