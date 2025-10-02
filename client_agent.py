# client_agent.py
import time
import os
import subprocess
import logging
import requests

# Configura��o de log em arquivo
logging.basicConfig(
    filename='client_agent.log',
    level=logging.DEBUG,
    format='%(asctime)s %(levelname)s:%(message)s'
)

# URL do seu servidor Flask (onde est� rodando o dashboard/app.py)
SERVER_URL = "http://192.168.0.5:5000"


# Caminho fixo do conc_shopping.exe na m�quina do cliente
EXE_PATH = r"C:\AUTOMACAO\conciliacao\bots\conc_shopping.exe"

def execute_command():
    try:
        resp = requests.get(f"{SERVER_URL}/get_command", timeout=10)
        data = resp.json()
        command = data.get("command")
        command_id = data.get("id")

        if not command:
            logging.debug("Nenhum comando pendente.")
            return

        if command.startswith("execute_conciliacao::"):
            shopping = command.split("::", 1)[1]
            if os.path.exists(EXE_PATH):
                subprocess.Popen([EXE_PATH, shopping])
                logging.info(f"Concilia��o iniciada para {shopping}")
            else:
                logging.error(f"conc_shopping.exe n�o encontrado em {EXE_PATH}")
        else:
            logging.warning(f"Comando desconhecido recebido: {command}")

        # Atualiza status no servidor
        try:
            requests.post(f"{SERVER_URL}/update_command", json={"command_id": command_id})
        except Exception as ex:
            logging.error(f"Erro ao atualizar status: {ex}")

    except Exception as e:
        logging.error(f"Erro no agente: {e}")

def main():
    logging.info("Agente iniciado e rodando...")
    while True:
        execute_command()
        time.sleep(5)  # checa a cada 5s

if __name__ == "__main__":
    main()
