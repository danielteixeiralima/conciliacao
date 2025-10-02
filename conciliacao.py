# -*- coding: utf-8 -*-
import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk
import pyautogui
import pygetwindow
import pyscreeze
import pytweening
import mouseinfo
#### FASES tipo 1
# y 20 - fase 1
# y 33 - fase 2
# ...

folder_map = {
    "Shopping Montserrat": r"C:\\Program Files\\Victor & Schellenberger_FAT_HOM\\VSSC_MONTSERRAT_HOM",
    "Shopping da Ilha":       r"C:\\Program Files\\Victor & Schellenberger_FAT_HOM\\VSSC_ILHA_HOM",
    "Shopping Rio Poty":      r"C:\\Program Files\\Victor & Schellenberger_FAT_HOM\\VSSC_TERESINA_HOM",
    "Shopping Metrópole":     r"C:\\Program Files\\Victor & Schellenberger_FAT_HOM\\VSSC_METROPOLE_HOM",
    "Shopping Moxuara":       r"C:\\Program Files\\Victor & Schellenberger_FAT_HOM\\VSSC_MOXUARA_HOM",
    "Shopping Mestre Álvaro": r"C:\\Program Files\\Victor & Schellenberger_FAT_HOM\\VSSC_MESTREALVARO_HOM"
}

# caminhos externos (ajuste se necessário)
BASE_AUTOMACAO = r"C:\AUTOMACAO\conciliacao\bots"
PATH_CONC_SHOP_EXE = os.path.join(BASE_AUTOMACAO, "conc_shopping.exe")
PATH_CONC_SHOP_PY  = os.path.join(BASE_AUTOMACAO, "conc_shopping.py")

PATH_CONC_INC_EXE  = os.path.join(BASE_AUTOMACAO, "conc_incorporadora.exe")
PATH_CONC_INC_PY   = os.path.join(BASE_AUTOMACAO, "conc_incorporadora.py")

def main():
    if len(sys.argv) == 2:
        shopping = sys.argv[1]
        if shopping in folder_map:
            executar_conciliacao(shopping)
        else:
            print("Shopping inválido ou não reconhecido.")
            sys.exit(1)
    else:
        criar_interface()

def center_window(window, width, height):
    window.update_idletasks()
    sw = window.winfo_screenwidth()
    sh = window.winfo_screenheight()
    x = (sw // 2) - (width // 2)
    y = (sh // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")

def run_external(path_exe, path_py, *args):
    if os.path.exists(path_exe):
        cmd = [path_exe, *args]
    else:
        cmd = ["python", path_py, *args]
    subprocess.Popen(cmd, shell=False)

def executar_conciliacao(shopping):
    run_external(PATH_CONC_SHOP_EXE, PATH_CONC_SHOP_PY, shopping)

def executar_incorporadora():
    run_external(PATH_CONC_INC_EXE, PATH_CONC_INC_PY)

def show_shopping_menu():
    global buttons
    for widget in main_frame.winfo_children():
        widget.destroy()
    label = ttk.Label(
        main_frame,
        text="Selecione um Shopping para Conciliação:",
        font=("Segoe UI", 11, "bold")
    )
    label.pack(pady=8)

    buttons = []
    for shopping in folder_map.keys():
        btn = ttk.Button(
            main_frame,
            text=shopping,
            style="Accent.TButton",
            command=lambda s=shopping: [executar_conciliacao(s), root.destroy()]
        )
        btn.pack(pady=4, fill='x')
        buttons.append(btn)

def criar_interface():
    global root, buttons, main_frame
    root = tk.Tk()
    root.title("Sistema de Conciliação")
    center_window(root, 400, 300)

    root.resizable(False, False)
    root.configure(bg="#e7f0ff")

    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('Main.TFrame', background='#e7f0ff')
    style.configure('TLabel', background='#e7f0ff', font=('Segoe UI', 11))
    style.configure('Accent.TButton',
                    font=('Segoe UI', 11),
                    foreground='white',
                    background='#3366CC',
                    padding=4)
    style.map('Accent.TButton',
              background=[('active', '#2853a3'), ('disabled', '#a6a6a6')])

    main_frame = ttk.Frame(root, style="Main.TFrame", padding=10)
    main_frame.place(relx=0.5, rely=0.5, anchor='center')

    label = ttk.Label(
        main_frame,
        text="Selecione o tipo de conciliação:",
        font=("Segoe UI", 11, "bold")
    )
    label.pack(pady=8)

    btn_shop = ttk.Button(
        main_frame,
        text="Shopping",
        style="Accent.TButton",
        command=show_shopping_menu
    )
    btn_shop.pack(pady=4, fill='x')

    btn_inc = ttk.Button(
        main_frame,
        text="Incorporadora",
        style="Accent.TButton",
        command=lambda: [executar_incorporadora(), root.destroy()]
    )
    btn_inc.pack(pady=4, fill='x')

    buttons = []
    root.mainloop()

if __name__ == "__main__":
    main()
