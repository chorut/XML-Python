import imaplib
import email
import os
from datetime import datetime
from email.utils import parsedate_to_datetime
from dotenv import load_dotenv

# Caminho onde os arquivos XML serão salvos
PASTA_DESTINO = r"C:\teste\teste\teste\teste"

# Data do dia atual
DATA_HOJE = datetime.now().date()

# Carrega as variáveis do .env
load_dotenv(dotenv_path="credenciais.env")
IMAP_SERVER = os.getenv("IMAP_SERVER")

def carregar_credenciais():
    contas = []
    for i in range(1, 21):
        usuario = os.getenv(f"EMAIL_{i}")
        senha = os.getenv(f"PASS_{i}")
        if usuario and senha:
            contas.append((usuario, senha))
    return contas

def conectar_imap(usuario, senha):
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(usuario, senha)
    return mail

def data_valida(msg):
    try:
        data_email = parsedate_to_datetime(msg["Date"])
        if data_email.tzinfo:
            data_email = data_email.replace(tzinfo=None)
        return data_email.date() == DATA_HOJE
    except:
        return False

def baixar_anexos_xml(usuario, senha):
    print(f"\n📥 Verificando conta: {usuario}")
    try:
        mail = conectar_imap(usuario, senha)
        mail.select("inbox")

        status, mensagens = mail.search(None, f'SINCE {DATA_HOJE.strftime("%d-%b-%Y")}')
        if status != "OK":
            print(f"⚠️ Erro ao buscar e-mails de {usuario}")
            return

        ids = mensagens[0].split()
        print(f"🔎 Total de e-mails hoje: {len(ids)}")

        for num in ids:
            status, dados = mail.fetch(num, '(BODY.PEEK[HEADER.FIELDS (DATE)])')
            if status != "OK":
                continue

            cabecalho = dados[0][1].decode(errors="ignore")
            msg_data = email.message_from_string(cabecalho)

            if not data_valida(msg_data):
                continue

            status, dados = mail.fetch(num, '(RFC822)')
            if status != "OK":
                continue

            raw_email = dados[0][1]
            msg = email.message_from_bytes(raw_email)

            for parte in msg.walk():
                if parte.get_content_maintype() == 'multipart':
                    continue
                if parte.get('Content-Disposition') is None:
                    continue

                nome_arquivo = parte.get_filename()
                if nome_arquivo and nome_arquivo.lower().endswith(".xml"):
                    os.makedirs(PASTA_DESTINO, exist_ok=True)
                    caminho_completo = os.path.join(PASTA_DESTINO, nome_arquivo)

                    if os.path.exists(caminho_completo):
                        print(f"⚠️ [{usuario}] Já existe: {nome_arquivo}")
                        continue

                    with open(caminho_completo, 'wb') as f:
                        f.write(parte.get_payload(decode=True))

                    print(f"✅ [{usuario}] Baixado: {nome_arquivo}")

        mail.logout()
    except Exception as e:
        print(f"❌ Erro com {usuario}: {e}")

# Execução principal
if __name__ == "__main__":
    contas = carregar_credenciais()
    if not contas:
        print("❌ Nenhuma conta encontrada no credenciais.env")
    for usuario, senha in contas:
        baixar_anexos_xml(usuario, senha)
