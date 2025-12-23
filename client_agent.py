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

EXE_PATH = r"C:\AUTOMACAO\conciliacao\bots\conc_shopping.exe"  # Caminho do executável do robô de conciliação

EXE_CALCULOS = r"C:\AUTOMACAO\faturamento\bots\hom_calculos.exe"       # Robô de cálculos
EXE_BOLETOS  = r"C:\AUTOMACAO\faturamento\bots\hom_gerar_boletos.exe"  # Robô de boletos
EXE_EMAIL    = r"C:\AUTOMACAO\faturamento\bots\hom_enviar_email.exe"   # Robô de envio de email

CHECK_INTERVAL = 5  # segundos entre checagens

logging.info("=" * 80)
logging.info(f"🚀 Agente iniciado - Conectando ao servidor: {SERVER_URL}")
logging.info(f"🔍 Monitorando comandos a cada {CHECK_INTERVAL}s")
logging.info(f"🧠 Executável de conciliação configurado em: {EXE_PATH}")
logging.info(f"🧮 Executável de cálculos configurado em: {EXE_CALCULOS}")
logging.info(f"💳 Executável de boletos configurado em: {EXE_BOLETOS}")
logging.info(f"📧 Executável de email configurado em: {EXE_EMAIL}")
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

def start_faturamento(exe_path: str, shopping: str, tipo: str) -> bool:
    """
    Inicia o robô de faturamento correto (cálculos ou boletos).
    """
    try:
        logging.info(f"📌 Tentando iniciar faturamento: {exe_path}")

        if not os.path.exists(exe_path):
            logging.error(f"❌ Executável NÃO encontrado em: {exe_path}")
            return False

        exe_dir = os.path.dirname(exe_path)
        exe_name = os.path.basename(exe_path)

        logging.info(f"🚀 Executando {exe_name} com argumentos: '{shopping}', '{tipo}'")
        subprocess.Popen(
            [exe_path, shopping, tipo],
            cwd=exe_dir,
            shell=False
        )
        return True

    except Exception as e:
        logging.exception(f"💥 Erro ao iniciar faturamento: {e}")
        return False

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

        # ======================================================
        # CONCILIAÇÃO
        # ======================================================
        if command.startswith("execute_conciliacao::"):
            shopping = command.split("::", 1)[1].strip()
            logging.info(f"🏬 Iniciando conciliação para: {shopping}")

            start_time = datetime.now()

            ok = start_conc_shopping(shopping)
            if not ok:
                return

            logging.info("🕓 Aguardando o processo conc_shopping.exe iniciar...")
            for _ in range(60):
                if any("conc_shopping.exe" in p.name().lower() for p in psutil.process_iter()):
                    logging.info("🟢 Processo conc_shopping.exe detectado — aguardando finalização...")
                    break
                _time.sleep(1)
            else:
                logging.warning("⚠️ O processo conc_shopping.exe não foi detectado em 60s.")
                return

            while any("conc_shopping.exe" in p.name().lower() for p in psutil.process_iter()):
                _time.sleep(5)

            end_time = datetime.now()
            duration = (end_time - start_time).total_seconds()
            logging.info(f"✅ Conciliação finalizada para {shopping} (Tempo total: {duration:.2f}s)")

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
        # FATURAMENTO
        # ======================================================
        elif command.startswith("execute_faturamento::"):
            try:
                _, acao, shopping, tipo = command.split("::")
                shopping = shopping.strip()
                tipo = tipo.strip()
                acao = acao.strip()

                logging.info(f"🧮 Iniciando FATURAMENTO ({acao}) para {shopping} | Tipo: {tipo}")

                if acao == "calculo":
                    exe_path = EXE_CALCULOS
                elif acao == "boletos":
                    exe_path = EXE_BOLETOS
                elif acao == "email":
                    exe_path = EXE_EMAIL
                else:
                    logging.error(f"❌ Ação de faturamento desconhecida: {acao}")
                    return

                ok = start_faturamento(exe_path, shopping, tipo)
                if not ok:
                    return

                process_name = os.path.basename(exe_path).lower()

                logging.info(f"🕓 Aguardando inicialização do {process_name}...")
                for _ in range(60):
                    if any(process_name in p.name().lower() for p in psutil.process_iter()):
                        logging.info(f"🟢 Processo {process_name} detectado — aguardando finalização...")
                        break
                    _time.sleep(1)
                else:
                    logging.warning(f"⚠️ O processo {process_name} não foi detectado em 60s.")
                    return

                while any(process_name in p.name().lower() for p in psutil.process_iter()):
                    _time.sleep(5)

                logging.info(f"🟢 Processo de FATURAMENTO ({acao}) finalizado.")

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
