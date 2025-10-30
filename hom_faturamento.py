# -*- coding: utf-8 -*-
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk
import logging

#### FASES tipo 1

# y 20 - fase 1
# y 33 - fase 2
# y 46 - fase 3
# y 59 - fase 4
# y 72 - fase 5
# y 85 - fase 6
# y 98 - fase 7
# y 111 - fase 8
# y 124 - fase 9
# y 137 - fase 10
# y 150 - fase 11
# y 163 - fase 12
# y 176 - fase 13
# y 189 - fase 14
# y 202 - fase 15
# y 215 - fase 16
# 1 click em (97, 215) depois no (-100, 215) - fase 17
# 2 clicks em (97, 215) depois no (-100, 215) - fase 18
# 3 clicks em (97, 215) depois no (-100, 215) - fase 19
# 4 clicks em (97, 215) depois no (-100, 215) - fase 20
# 5 clicks em (97, 215) depois no (-100, 215) - fase 21
# 6 clicks em (97, 215) depois no (-100, 215) - fase 22
# 7 clicks em (97, 215) depois no (-100, 215) - fase 23
# 8 clicks em (97, 215) depois no (-100, 215) - fase 24
# 9 clicks em (97, 215) depois no (-100, 215) - fase 25
# 10 clicks em (97, 215) depois no (-100, 215) - fase 26
# 11 clicks em (97, 215) depois no (-100, 215) - fase 27
# 12 clicks em (97, 215) depois no (-100, 215) - fase 28
# 13 clicks em (97, 215) depois no (-100, 215) - fase 29
# 14 clicks em (97, 215) depois no (-100, 215) - fase 30
# 15 clicks em (97, 215) depois no (-100, 215) - fase 31
# 16 clicks em (97, 215) depois no (-100, 215) - fase 32
# 17 clicks em (97, 215) depois no (-100, 215) - fase 33
# 18 clicks em (97, 215) depois no (-100, 215) - fase 34


#### FASES tipo 2

# press down 1 time - fase 1
# press down 2 times - fase 2
# press down 3 times - fase 3
# press down 4 times - fase 4
# press down 5 times - fase 5
# press down 6 times - fase 6
# press down 7 times - fase 7
# press down 8 times - fase 8
# press down 9 times - fase 9
# press down 10 times - fase 10
# press down 11 times - fase 11
# press down 12 times - fase 12
# press down 13 times - fase 13
# press down 14 times - fase 14
# press down 15 times - fase 15
# press down 16 times - fase 16
# press down 17 times - fase 17
# press down 18 times - fase 18
# press down 19 times - fase 19
# press down 20 times - fase 20
# press down 21 times - fase 21
# press down 22 times - fase 22
# press down 23 times - fase 23
# press down 24 times - fase 24
# press down 25 times - fase 25
# press down 26 times - fase 26
# press down 27 times - fase 27
# press down 28 times - fase 28
# press down 29 times - fase 29
# press down 30 times - fase 30
# press down 31 times - fase 31
# press down 32 times - fase 32
# press down 33 times - fase 33
# press down 34 times - fase 34
# press down 35 times - fase 35
# press down 36 times - fase 36
# press down 37 times - fase 37
# press down 38 times - fase 38

#### FASES tipo 3

# y 25 - fase 1
# y 39 - fase 2
# y 53 - fase 3
# y 67 - fase 4
# y 81 - fase 5
# y 95 - fase 6
# y 109 - fase 7
# y 123 - fase 8
# y 137 - fase 9
# y 151 - fase 10
# y 165 - fase 11
# y 179 - fase 12
# y 193 - fase 13
# click em (2, 193) depois no (-100, 193) - fase 14
# 2 clicks em (2, 193) depois no (-100, 193) - fase 15
# 3 clicks em (2, 193) depois no (-100, 193) - fase 16
# 4 clicks em (2, 193) depois no (-100, 193) - fase 17
# 5 clicks em (2, 193) depois no (-100, 193) - fase 18
# 6 clicks em (2, 193) depois no (-100, 193) - fase 19
# 7 clicks em (2, 193) depois no (-100, 193) - fase 20
# 8 clicks em (2, 193) depois no (-100, 193) - fase 21
# 9 clicks em (2, 193) depois no (-100, 193) - fase 22
# 10 clicks em (2, 193) depois no (-100, 193) - fase 23
# 11 clicks em (2, 193) depois no (-100, 193) - fase 24
# 12 clicks em (2, 193) depois no (-100, 193) - fase 25
# 13 clicks em (2, 193) depois no (-100, 193) - fase 26
# 14 clicks em (2, 193) depois no (-100, 193) - fase 27
# 15 clicks em (2, 193) depois no (-100, 193) - fase 28
# 16 clicks em (2, 193) depois no (-100, 193) - fase 29
# 17 clicks em (2, 193) depois no (-100, 193) - fase 30
# 18 clicks em (2, 193) depois no (-100, 193) - fase 31
# 19 clicks em (2, 193) depois no (-100, 193) - fase 32
# 20 clicks em (2, 193) depois no (-100, 193) - fase 33
# 21 clicks em (2, 193) depois no (-100, 193) - fase 34
# 22 clicks em (2, 193) depois no (-100, 193) - fase 35
# 23 clicks em (2, 193) depois no (-100, 193) - fase 36
# 24 clicks em (2, 193) depois no (-100, 193) - fase 37
# 25 clicks em (2, 193) depois no (-100, 193) - fase 38


# Mapeamento dos shoppings para suas pastas (removido o "Shopping Praia da Costa")
folder_map = {
    "Shopping Montserrat": r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_MONTSERRAT_HOM",
    "Shopping da Ilha": r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_ILHA_HOM",
    "Shopping Rio Poty": r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_TERESINA_HOM",
    "Shopping Metrópole": r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_METROPOLE_HOM",
    "Shopping Moxuara": r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_MOXUARA_HOM",
    "Shopping Mestre Álvaro": r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_MESTREALVARO_HOM"
}

# Tipos de faturamento disponíveis
tipos_faturamento = ["Postecipados","Antecipados", "Atípicos",]

# Scripts disponíveis para execução
scripts = {
    "Gerar cálculos": r"C:\AUTOMACAO\faturamento\bots\hom_calculos.exe",
    "Gerar boletos": r"C:\AUTOMACAO\faturamento\bots\hom_gerar_boletos.exe",
    "Enviar e-mails": r"C:\AUTOMACAO\faturamento\bots\hom_enviar_email.exe"
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
