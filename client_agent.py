# client_agent.py
import time
import os
import subprocess
import logging
import requests
from datetime import datetime
import getpass
import platform
import psutil
import time as _time
import sys

# ===========================
# IDENTIDADE / LOG
# ===========================
current_user = getpass.getuser()
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
LOG_FILE = os.path.join(BASE_DIR, 'client_agent.log')
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s'
)
logging.info(f"👤 Agente rodando como usuário: {current_user} - Sistema: {platform.node()}")

# ===========================
# CONFIG
# ===========================
SERVER_URL = "http://34.67.108.173"  # URL do servidor Flask (dashboard/app.py)
EXE_PATH = r"C:\AUTOMACAO\conciliacao\bots\conc_shopping.exe"  # Caminho do executável do robô
CHECK_INTERVAL = 5  # segundos entre checagens

logging.info("=" * 80)
logging.info(f"🚀 Agente iniciado - Conectando ao servidor: {SERVER_URL}")
logging.info(f"🔍 Monitorando comandos a cada {CHECK_INTERVAL}s")
logging.info(f"🧠 Executável configurado em: {EXE_PATH}")
logging.info("=" * 80)

def start_conc_shopping(shopping: str) -> bool:
    """
    Inicia o conc_shopping.exe passando o nome do shopping como argumento.
    Retorna True se conseguiu iniciar, False caso contrário.
    """
    try:
        if not os.path.exists(EXE_PATH):
            logging.error(f"❌ Executável não encontrado em: {EXE_PATH}")
            return False

        exe_dir = os.path.dirname(EXE_PATH)
        logging.debug(f"🚀 Iniciando conc_shopping.exe com argumento: {shopping}")
        subprocess.Popen(
            [EXE_PATH, shopping],
            cwd=exe_dir,
            shell=False
        )
        return True
    except Exception as e:
        logging.exception(f"💥 Erro ao iniciar conc_shopping.exe: {e}")
        return False

def run_shopping(shopping):
    print(f"🏬 Iniciando {shopping}...")
    process = subprocess.Popen([r"C:\AUTOMACAO\conciliacao\bots\conc_shopping.exe", shopping])
    process.wait()  # 🔥 Aqui ele fica parado até o robô fechar o VSLoader e encerrar o exe
    print(f"✅ {shopping} finalizado às {time.strftime('%H:%M:%S')}")
# ===========================
# EXECUÇÃO DO COMANDO
# ===========================
def execute_command():
    try:
        logging.debug("➡️ Solicitando comando pendente do servidor...")
        resp = requests.get(f"{SERVER_URL}/get_command", timeout=10)

        logging.debug(f"🔁 Resposta HTTP: {resp.status_code}")
        if resp.status_code != 200:
            logging.error(f"❌ Erro HTTP ao consultar servidor: {resp.status_code} - {resp.text}")
            return

        data = resp.json()
        command = data.get("command")
        command_id = data.get("id")

        if not command:
            logging.debug("⏳ Nenhum comando pendente no servidor.")
            return

        logging.info(f"🆕 Novo comando recebido: {command} (ID: {command_id})")

        if command.startswith("execute_conciliacao::"):
            shopping = command.split("::", 1)[1].strip()
            logging.info(f"🏬 Iniciando conciliação para: {shopping}")

            start_time = datetime.now()

            # === Inicia direto o EXE com o argumento do shopping ===
            ok = start_conc_shopping(shopping)
            if not ok:
                return

            # Aguarda o processo conc_shopping.exe iniciar (até 60s)
            logging.info("🕓 Aguardando o processo conc_shopping.exe iniciar...")
            for _ in range(60):
                if any("conc_shopping.exe" in p.name().lower() for p in psutil.process_iter()):
                    logging.info("🟢 Processo conc_shopping.exe detectado — aguardando finalização...")
                    break
                _time.sleep(1)
            else:
                logging.warning("⚠️ O processo conc_shopping.exe não foi detectado em 60s.")
                return

            # Aguarda finalizar (verifica a cada 5s)
            while any("conc_shopping.exe" in p.name().lower() for p in psutil.process_iter()):
                _time.sleep(5)

            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()
            logging.info(f"✅ Conciliação finalizada para {shopping} (Tempo total: {duration:.2f}s)")

            # Atualiza status no servidor
            try:
                logging.debug(f"📡 Atualizando status no servidor para 'completed' (command_id={command_id})...")
                post_resp = requests.post(
                    f"{SERVER_URL}/update_command",
                    json={"command_id": command_id},
                    timeout=10
                )
                if post_resp.status_code == 200:
                    logging.info(f"🟢 Status atualizado com sucesso no servidor (command_id={command_id})")
                else:
                    logging.error(f"❌ Falha ao atualizar status no servidor ({post_resp.status_code}): {post_resp.text}")
            except Exception as ex:
                logging.exception(f"💥 Erro de conexão ao atualizar status: {ex}")
         # ======================================================
        # FATURAMENTO → executa hom_calculos.exe
        # ======================================================
        elif command.startswith("execute_faturamento::"):
            try:
                _, acao, shopping, tipo = command.split("::")
                shopping = shopping.strip()
                tipo = tipo.strip()

                logging.info(f"🧮 Iniciando FATURAMENTO ({acao}) para {shopping} | Tipo: {tipo}")

                exe_path = r"C:\AUTOMACAO\faturamento\bots\hom_calculos.exe"
                exe_dir = os.path.dirname(exe_path)

                if not os.path.exists(exe_path):
                    logging.error(f"❌ hom_calculos.exe não encontrado em: {exe_path}")
                    return

                # Inicia o processo
                proc = subprocess.Popen(
                    [exe_path, shopping, tipo],
                    cwd=exe_dir,
                    shell=False
                )

                logging.info("🕓 Aguardando término do processo de FATURAMENTO...")
                proc.wait()
                logging.info("🟢 Processo de FATURAMENTO finalizado.")

                # Atualiza o comando como concluído
                try:
                    post_resp = requests.post(
                        f"{SERVER_URL}/update_command",
                        json={"command_id": command_id},
                        timeout=10
                    )
                    if post_resp.status_code == 200:
                        logging.info(f"🟢 Status do FATURAMENTO atualizado no servidor (command_id={command_id})")
                    else:
                        logging.error(f"❌ Falha ao atualizar status do FATURAMENTO ({post_resp.status_code})")
                except Exception as ex:
                    logging.exception(f"💥 Erro ao atualizar status do FATURAMENTO: {ex}")

            except Exception as e:
                logging.exception(f"💥 Erro ao processar comando de FATURAMENTO: {e}")
                return
        else:
            logging.warning(f"⚠️ Comando desconhecido recebido: {command}")
            return

    except requests.exceptions.ConnectionError:
        logging.error("🌐 Conexão recusada - servidor indisponível.")
    except requests.exceptions.Timeout:
        logging.error("⏰ Timeout ao tentar se comunicar com o servidor.")
    except ValueError as e:
        logging.exception(f"💥 Erro ao processar resposta JSON do servidor: {e}")
    except Exception as e:
        logging.exception(f"💥 Erro inesperado no agente: {e}")


# ===========================
# LOOP PRINCIPAL DO AGENTE
# ===========================
def main():
    logging.info("🟦 Agente de automação iniciado e monitorando servidor...")
    while True:
        try:
            execute_command()
        except Exception as loop_err:
            logging.exception(f"🔥 Erro no loop principal: {loop_err}")
        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logging.warning("🟥 Execução interrompida manualmente pelo usuário.")
    except Exception as fatal:
        logging.critical(f"💀 Falha fatal no agente: {fatal}")
