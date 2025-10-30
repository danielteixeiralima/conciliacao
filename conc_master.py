# conc_master.py
import subprocess
import time
import logging
import psutil

# Configuração de log geral
logging.basicConfig(
    filename=r"C:\AUTOMACAO\conciliacao\bots\conc_master.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

SHOPPINGS = [
    "Shopping da Ilha",
    "Shopping Mestre Álvaro",
    "Shopping Moxuara",
    "Shopping Montserrat",
    "Shopping Metrópole",
    "Shopping Rio Poty"
]

EXE_PATH = r"C:\AUTOMACAO\conciliacao\bots\conc_shopping.exe"

def kill_vsloader():
    """Força o fechamento do VSLOADER.exe."""
    for proc in psutil.process_iter(['pid', 'name']):
        if 'VSLOADER' in proc.info['name'].upper():
            try:
                proc.kill()
                logging.warning(f"💀 VSLOADER.EXE finalizado (PID={proc.info['pid']})")
            except Exception as e:
                logging.error(f"Erro ao tentar encerrar VSLOADER.EXE: {e}")

def run_shopping(shopping):
    logging.info(f"🏬 Iniciando conciliação para {shopping}")
    start_time = time.time()

    try:
        process = subprocess.Popen([EXE_PATH, shopping])
        while True:
            # Timeout de segurança: 2 horas por shopping
            if time.time() - start_time > 7200:
                logging.error(f"⏰ Timeout atingido (2h) para {shopping}, finalizando VSLOADER.")
                kill_vsloader()
                process.kill()
                raise TimeoutError(f"{shopping} excedeu tempo máximo.")
            
            # Verifica se terminou
            retcode = process.poll()
            if retcode is not None:
                logging.info(f"✅ {shopping} finalizado com código {retcode}")
                break
            time.sleep(10)

    except Exception as e:
        logging.error(f"❌ Erro ao executar {shopping}: {e}")
        kill_vsloader()

    time.sleep(10)

def main():
    logging.info("🚀 Início do processo de conciliação em sequência")
    for s in SHOPPINGS:
        run_shopping(s)
    logging.info("🎉 Todos os shoppings processados.")

if __name__ == "__main__":
    main()
