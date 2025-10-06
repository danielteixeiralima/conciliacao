# app.py

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

# ====== Banco simples em memória (lista) ======
# ====== Banco persistente em disco ======


COMMANDS_FILE = "commands.json"

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

# Carrega a fila persistente
commands = load_commands()
command_counter = len(commands) + 1



# ====== Funções do agente (lado cliente) ======
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

        if command.startswith("execute_conciliacao::"):
            shopping = command.split("::", 1)[1]
            exe_path = r"C:\AUTOMACAO\conciliacao\bots\conc_shopping.exe"
            if os.path.exists(exe_path):
                subprocess.Popen([exe_path, shopping])
                logging.info(f"Conciliação iniciada para {shopping}")
            else:
                logging.error(f"conc_shopping.exe não encontrado em {exe_path}")

        elif command == "execute_vsloader":
            app_path = r"C:\Program Files\Victor & Schellenberger_FAT_HOM\VSSC_ILHA_HOM\VSLOADER.exe"
            if os.path.exists(app_path):
                subprocess.Popen([app_path])
                logging.info("VSLOADER iniciado com sucesso!")
            else:
                logging.error(f"VSLOADER não encontrado em {app_path}")

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


# Se você quiser rodar o agente no mesmo processo do servidor:
# agent_thread = Thread(target=run_client_agent, daemon=True)
# agent_thread.start()


# ====== Rotas Web ======
@app.route('/')
def index():
    return render_template('dashboard.html')


@app.route('/start_conciliacao')
def start_conciliacao():
    """Enfileira conciliação para um shopping"""
    global commands, command_counter
    shopping = request.args.get('shopping')
    if not shopping:
        return "Shopping não informado", 400

    cmd = {
        "id": command_counter,
        "command": f"execute_conciliacao::{shopping}",
        "status": "pending",
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
    }
    commands.append(cmd)
    command_counter += 1
    save_commands(commands)
    return f"Comando enviado para executar conc_shopping com {shopping}"




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
    """Lista todos os comandos com status"""
    global commands
    return jsonify(commands)

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
    # Roda local em http://127.0.0.1:5000
    app.run(host="0.0.0.0", port=5000, debug=True)
