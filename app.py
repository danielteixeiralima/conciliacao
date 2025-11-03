import os
import time
import subprocess
import logging
import requests
from threading import Thread
from flask import Flask, render_template, request, jsonify
import json

logging.basicConfig(level=logging.DEBUG)
app = Flask(__name__)

# ====== Configuração ======
logging.getLogger('werkzeug').setLevel(logging.INFO)
COMMANDS_FILE = "commands.json"

# ====== Funções utilitárias ======
def load_commands():
    if os.path.exists(COMMANDS_FILE):
        try:
            with open(COMMANDS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logging.error(f"Erro ao carregar comandos: {e}")
            return []
    return []

def save_commands(data):
    try:
        with open(COMMANDS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logging.error(f"Erro ao salvar comandos: {e}")

# ====== Banco persistente ======
commands = load_commands()
command_counter = len(commands) + 1

commands = []        # limpa tudo ao iniciar
command_counter = 1  # reinicia IDs
save_commands(commands)
logging.info(">>> Lista de comandos limpa ao iniciar servidor.")

# ====== Agente cliente ======
SERVER_URL = "http://127.0.0.1:5000"

def execute_command():
    """Agente: busca e executa comando pendente no servidor"""
    global commands
    try:
        response = requests.get(f"{SERVER_URL}/get_command")
        data = response.json()
        command = data.get("command")
        command_id = data.get("id")

        if not command:
            logging.debug("Nenhum comando pendente.")
            return

        # --- Conciliação ---
        if command.startswith("execute_conciliacao::"):
            shopping = command.split("::", 1)[1]
            exe_path = r"C:\AUTOMACAO\conciliacao\bots\conc_shopping.exe"
            exe_dir = os.path.dirname(exe_path)
            if os.path.exists(exe_path):
                subprocess.Popen(
                    [exe_path, shopping],
                    cwd=exe_dir,  # 🔧 força o diretório correto
                )
                logging.info(f"Conciliação iniciada para {shopping} (cwd={exe_dir})")
            else:
                logging.error(f"conc_shopping.exe não encontrado em {exe_path}")


        # --- VSLoader ---
        elif command == "execute_vsloader":
            app_path = r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_ILHA_HOM\VSLOADER.exe"
            if os.path.exists(app_path):
                subprocess.Popen([app_path])
                logging.info("VSLOADER iniciado com sucesso!")
            else:
                logging.error(f"VSLOADER não encontrado em {app_path}")

        # --- Faturamento ---
        elif command.startswith("execute_faturamento::"):
            try:
                _, acao, shopping, tipo = command.split("::")
                exe_path = r"C:\AUTOMACAO\faturamento\bots\faturamento.exe"
                if os.path.exists(exe_path):
                    subprocess.Popen([exe_path, acao, shopping, tipo])
                    logging.info(f"Faturamento ({acao}) iniciado para {shopping} ({tipo})")
                else:
                    logging.error(f"faturamento.exe não encontrado em {exe_path}")
            except Exception as err:
                logging.error(f"Erro ao interpretar comando de faturamento: {err}")

        # Atualiza status
        try:
            requests.post(f"{SERVER_URL}/update_command", json={"command_id": command_id})
        except Exception as ex:
            logging.error(f"Erro ao atualizar status: {ex}")

    except Exception as e:
        logging.error(f"Erro no agente: {e}")

def run_client_agent():
    logging.info("Agente de automação iniciado.")
    while True:
        execute_command()
        time.sleep(5)

# Se quiser rodar o agente junto com o servidor:
# agent_thread = Thread(target=run_client_agent, daemon=True)
# agent_thread.start()

# ====== Rotas Web ======

@app.route('/')
def index():
    return render_template('conciliacao.html', title="Conciliação", active="conciliacao")

@app.route('/faturamento')
def faturamento():
    return render_template('faturamento.html', title="Faturamento", active="faturamento")

@app.route('/start_conciliacao')
def start_conciliacao():
    """Enfileira conciliação para todos os shoppings automaticamente"""
    global commands, command_counter

    # lista de shoppings
    shoppings = [
        "Shopping da Ilha",
        "Shopping Mestre Álvaro",
        "Shopping Moxuara",
        "Shopping Montserrat",
        "Shopping Metrópole",
        "Shopping Rio Poty",
        "Shopping Praia da Costa"
    ]

    for shopping in shoppings:
        cmd = {
            "id": command_counter,
            "command": f"execute_conciliacao::{shopping}",
            "status": "pending",
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
        }
        commands.append(cmd)
        command_counter += 1

    save_commands(commands)
    return f"Comandos de conciliação enfileirados para {len(shoppings)} shoppings."

@app.route('/start_faturamento')
def start_faturamento():
    """Enfileira comandos de faturamento"""
    global commands, command_counter
    shopping = request.args.get('shopping')
    tipo = request.args.get('tipo')
    acao = request.args.get('acao')
    if not (shopping and tipo and acao):
        return "Parâmetros incompletos", 400

    cmd = {
        "id": command_counter,
        "command": f"execute_faturamento::{acao}::{shopping}::{tipo}",
        "status": "pending",
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
    }
    commands.append(cmd)
    command_counter += 1
    save_commands(commands)
    return f"Comando de {acao} ({tipo}) enviado para {shopping}"

@app.route('/start_vsloader')
def start_vsloader():
    """Enfileira execução do VSLOADER"""
    global commands, command_counter
    cmd = {
        "id": command_counter,
        "command": "execute_vsloader",
        "status": "pending"
    }
    commands.append(cmd)
    command_counter += 1
    return "Comando enviado para executar VSLOADER.exe"

@app.route('/get_command')
def get_command():
    """Agente cliente consulta o próximo comando"""
    global commands
    for cmd in commands:
        if cmd["status"] == "pending":
            cmd["status"] = "in_progress"
            return jsonify({"command": cmd["command"], "id": cmd["id"]})
    return jsonify({"command": None, "id": None})

@app.route('/get_commands')
def get_commands():
    global commands
    ativos = [c for c in commands if c["status"] != "completed"]
    return jsonify(ativos)

@app.route('/update_command', methods=['POST'])
def update_command():
    """Agente marca comando como concluído"""
    global commands
    data = request.get_json()
    command_id = data.get("command_id")
    updated = False
    for cmd in commands:
        if cmd["id"] == command_id:
            cmd["status"] = "completed"
            cmd["completed_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
            updated = True
            break
    if updated:
        save_commands(commands)
        return jsonify({"message": "Command completed successfully."})
    return jsonify({"error": "Command not found."}), 404

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000, debug=True)
