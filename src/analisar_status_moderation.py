import requests
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import logging
from datetime import datetime
import os
import shutil
import traceback

# ============================================================
# CONFIGURAÇÕES GERAIS
# ============================================================
API_URL = "https://opthub.layer.core.dcg.com.br/v1/Profile/API.svc/web/GetCustomer"
BASE_DIR = os.path.dirname(__file__)
LOG_FILE = os.path.join(BASE_DIR, "log_execucao_moderation.log")
OUTPUT_FILE = os.path.join(BASE_DIR, "Clientes_Pendentes.xlsx")

SENDER_EMAIL = "buyer.hb.opthub@gmail.com"
SENDER_PASSWORD = "app-password-ou-senha-aqui"
RECIPIENTS = "bruno@compreoculos.com.br,brunoera@gmail.com,comercial@opthub.com.br"

# ============================================================
# LOG
# ============================================================
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logging.info("==== Início da execução do script ====")

# ============================================================
# FUNÇÃO PARA BUSCAR DADOS DO CLIENTE
# ============================================================
def get_customer_data(customer_id):
    try:
        payload = {
            "Page": {"PageIndex": 0, "PageSize": 0},
            "SellerID": 0,
            "CustomerID": customer_id,
        }

        response = requests.post(API_URL, json=payload)
        if response.status_code != 200:
            logging.warning(f"Cliente {customer_id} retornou status {response.status_code}")
            return None

        data = response.json()

        cliente_info = {
            "CustomerID": data.get("CustomerID"),
            "Nome": data.get("Name"),
            "Email": data.get("Email"),
            "Celular": data.get("Contact", {}).get("CellPhone", ""),
            "Telefone": data.get("Contact", {}).get("Phone", ""),
            "CustomerStatusID": data.get("CustomerStatusID"),
            "CustomerType": data.get("CustomerType"),
        }

        logging.info(f"Cliente {customer_id} coletado com sucesso.")
        return cliente_info

    except Exception as e:
        logging.error(f"Erro ao buscar cliente {customer_id}: {e}")
        return None

# ============================================================
# CONSULTA DOS CLIENTES
# ============================================================
def processar_clientes(lista_clientes):
    resultados = []
    for customer_id in lista_clientes:
        info = get_customer_data(customer_id)
        if info:
            resultados.append(info)
    return resultados

# ============================================================
# GERAÇÃO DO EXCEL
# ============================================================
def salvar_excel(clientes):
    try:
        df = pd.DataFrame(clientes)
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Clientes")
        logging.info(f"Arquivo Excel gerado: {OUTPUT_FILE}")
    except Exception as e:
        logging.error(f"Erro ao salvar Excel: {e}")

# ============================================================
# ENVIO DE EMAIL
# ============================================================
def enviar_email():
    try:
        msg = MIMEMultipart()
        msg["From"] = SENDER_EMAIL
        msg["To"] = RECIPIENTS
        msg["Subject"] = f"Relatório de Clientes PENDENTES - {datetime.now().strftime('%d/%m/%Y')}"

        body = "Segue em anexo o relatório atualizado de clientes pendentes de aceite e o log de execução."
        msg.attach(MIMEText(body, "plain"))

        # anexar arquivos
        for file_path in [OUTPUT_FILE, LOG_FILE]:
            if os.path.exists(file_path):
                part = MIMEBase("application", "octet-stream")
                with open(file_path, "rb") as file:
                    part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
                msg.attach(part)
                logging.info(f"Anexo adicionado: {file_path}")
            else:
                logging.warning(f"Arquivo para anexo não encontrado: {file_path}")

        # Envio SMTP
        logging.info("Tentando enviar e-mail...")
        server = smtplib.SMTP("smtp.gmail.com", 587, timeout=30)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()

        logging.info("E-mail enviado com sucesso.")
        print("✅ E-mail enviado com sucesso.")

    except Exception as e:
        logging.error(f"Erro ao enviar e-mail: {e}")
        logging.error(traceback.format_exc())
        print("❌ Falha ao enviar e-mail:", e)

# ============================================================
# EXECUÇÃO PRINCIPAL
# ============================================================
if __name__ == "__main__":
    try:
        lista_clientes = [64, 68, 69]  # IDs de exemplo
        logging.info(f"Iniciando coleta de {len(lista_clientes)} clientes...")

        clientes = processar_clientes(lista_clientes)
        if clientes:
            salvar_excel(clientes)
            enviar_email()
        else:
            logging.warning("Nenhum cliente retornado da API.")

    except Exception as e:
        logging.error(f"Erro geral na execução: {e}")

    # Copiar log para o diretório raiz do repositório
    try:
        raiz_repo = os.path.abspath(os.path.join(BASE_DIR, ".."))
        shutil.copy(LOG_FILE, os.path.join(raiz_repo, "log_execucao_moderation.log"))
        logging.info("Log copiado para o diretório raiz do repositório.")
    except Exception as e:
        logging.error(f"Falha ao copiar log para o repositório: {e}")

    logging.info("==== Fim da execução do script ====")
