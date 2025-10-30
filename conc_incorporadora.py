# filename: conc_incorporadora.py
# -*- coding: utf-8 -*-
import os
import sys
import datetime
import zipfile

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

# Garante que a pasta de trabalho seja a mesma do executável
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

# === Configuração do Drive ===
# Para conseguir listar/apagar arquivos antigos use escopo completo
SCOPES = ['https://www.googleapis.com/auth/drive']

# Se quiser restringir a uma pasta no Drive, defina DRIVE_FOLDER_ID no .env ou ambiente
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID", "").strip() or None

# --- Resolve caminho base (suporta PyInstaller) ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Caminhos dos arquivos sensíveis (sempre junto ao exe)
TOKEN_PATH = os.path.join(BASE_DIR, 'token.json')
CREDENTIALS_PATH = os.path.join(BASE_DIR, 'credentials.json')


def criar_zip(network_path, output_dir):
    """
    Cria 'arquivos.zip' contendo todos os arquivos .RET do dia atual,
    exceto quando for segunda-feira — nesse caso, pega os arquivos de sábado.
    """

    hoje = datetime.date.today()
    # Se for segunda, pega sábado (2 dias antes)
    if hoje.weekday() == 0:  # 0 = segunda-feira
        target_date = hoje - datetime.timedelta(days=2)
        print(f"[info] Hoje é segunda-feira → buscando arquivos de {target_date} (sábado).")
    else:
        target_date = hoje
        print(f"[info] Buscando arquivos de {target_date} (dia atual).")

    zip_filename = os.path.join(output_dir, "arquivos.zip")

    # Remove arquivo ZIP antigo
    if os.path.exists(zip_filename):
        try:
            os.remove(zip_filename)
            print(f"[diag] Arquivo antigo removido: {zip_filename}")
        except Exception as e:
            print(f"[aviso] Não foi possível remover arquivo antigo: {e}")

    if not os.path.isdir(output_dir):
        os.makedirs(output_dir)

    count = 0
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for empresa in os.listdir(network_path):
            empresa_path = os.path.join(network_path, empresa)
            if not os.path.isdir(empresa_path):
                continue
            for banco in os.listdir(empresa_path):
                banco_path = os.path.join(empresa_path, banco)
                if not os.path.isdir(banco_path):
                    continue
                for arquivo in os.listdir(banco_path):
                    if arquivo.lower().endswith(".ret"):
                        arquivo_path = os.path.join(banco_path, arquivo)
                        try:
                            mod_date = datetime.date.fromtimestamp(os.path.getmtime(arquivo_path))
                            if mod_date == target_date:
                                zipf.write(arquivo_path, arcname=arquivo)
                                count += 1
                                print(f"[zip] Incluído: {arquivo}")
                        except Exception as e:
                            print(f"[aviso] Erro ao processar {arquivo}: {e}")

    print(f"[OK] ZIP criado com {count} arquivos do dia {target_date}: {zip_filename}")
    return zip_filename



def autenticar_drive():
    """
    Usa token.json se existir; se não existir roda fluxo OAuth
    """
    creds = None
    if os.path.exists(TOKEN_PATH):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
        except Exception as e:
            print("[aviso] token.json inválido, vai refazer login:", e)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            print("[diag] token expirado — refresh feito.")
        else:
            print("[diag] Abrindo navegador para login OAuth...")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, 'w') as token:
            token.write(creds.to_json())
            print(f"[diag] token.json salvo em {TOKEN_PATH}")

    service = build('drive', 'v3', credentials=creds)
    return service


def delete_old_drive_zips(service, filename="arquivos.zip", folder_id=None):
    """
    Remove do Google Drive todos os arquivos com 'filename'.
    Se 'folder_id' for informado, limita a busca àquela pasta.
    """
    q_parts = [f"name = '{filename}'", "trashed = false"]
    if folder_id:
        q_parts.append(f"'{folder_id}' in parents")
    query = " and ".join(q_parts)

    deleted = 0
    page_token = None
    while True:
        resp = service.files().list(
            q=query,
            spaces='drive',
            fields='nextPageToken, files(id, name, parents)',
            pageToken=page_token
        ).execute()

        for f in resp.get('files', []):
            try:
                service.files().delete(fileId=f['id']).execute()
                deleted += 1
                print(f"[diag] Removido do Drive: {f['name']} ({f['id']})")
            except Exception as e:
                print(f"[aviso] Falha ao remover {f.get('id')}: {e}")

        page_token = resp.get('nextPageToken')
        if not page_token:
            break

    print(f"[OK] Remoção no Drive concluída. Arquivos apagados: {deleted}")


def upload_to_drive(service, file_path, folder_id=None):
    """
    Sobe o arquivo para o Drive. Se passar folder_id, coloca dentro da pasta.
    """
    file_metadata = {'name': os.path.basename(file_path)}
    if folder_id:
        file_metadata['parents'] = [folder_id]

    media = MediaFileUpload(file_path, mimetype='application/zip', resumable=True)
    uploaded = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"[OK] Arquivo enviado ao Drive. ID: {uploaded['id']}")
    return uploaded['id']


def main():
    # Caminho na rede (Windows ou mapeado)
    network_path = r"\\192.168.18.4\hnc\COBRANCA_INCORPORACAO\RETORNO"

    # Diretório local onde salvar o ZIP
    output_dir = r"C:\AUTOMACAO\conciliacao"

    # 1. Cria ZIP (já limpa antigo local)
    zip_path = criar_zip(network_path, output_dir)

    # 2. Autentica Drive
    service = autenticar_drive()

    # 3. Remove ZIPs antigos no Drive
    delete_old_drive_zips(service, filename="arquivos.zip", folder_id=DRIVE_FOLDER_ID)

    # 4. Sobe ZIP para o Drive
    upload_to_drive(service, zip_path, folder_id=DRIVE_FOLDER_ID)



if __name__ == "__main__":
    try:
        with open("conciliacao_log.txt", "a", encoding="utf-8") as log:
            log.write(f"[{datetime.datetime.now()}] Iniciando execução\n")
        main()
        with open("conciliacao_log.txt", "a", encoding="utf-8") as log:
            log.write(f"[{datetime.datetime.now()}] Execução concluída com sucesso\n")
    except Exception as e:
        with open("conciliacao_log.txt", "a", encoding="utf-8") as log:
            log.write(f"[{datetime.datetime.now()}] ERRO: {e}\n")
