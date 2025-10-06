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
SERVER_URL = "http://34.67.108.173"

logging.info(f"Conectando ao servidor: {SERVER_URL}")


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
                try:
                    process = subprocess.Popen([EXE_PATH, shopping])
                    logging.info(f"Conciliação iniciada para {shopping} (PID: {process.pid})")
                    # Espera breve pra confirmar se abriu sem erro
                    time.sleep(3)
                    if process.poll() is None:
                        logging.info(f"Processo conc_shopping.exe está ativo para {shopping}")
                    else:
                        logging.error(f"Processo conc_shopping.exe terminou imediatamente (erro possível).")
                except Exception as sub_err:
                    logging.error(f"Erro ao iniciar conc_shopping.exe: {sub_err}")
                    return
            else:
                logging.error(f"conc_shopping.exe não encontrado em {EXE_PATH}")
                return
        else:
            logging.warning(f"Comando desconhecido recebido: {command}")

        # Atualiza status no servidor apenas se passou sem erro
        try:
            requests.post(f"{SERVER_URL}/update_command", json={"command_id": command_id})
            logging.debug(f"Status atualizado para concluído (command_id={command_id})")
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
