import os
import json
import requests
import smtplib
import ssl
import pandas as pd
from email.message import EmailMessage
from datetime import datetime
from openpyxl import load_workbook

# Configurações
OPTHUB_USER = os.getenv("OPTHUB_USER", "bruno.opthub")
OPTHUB_PASS = os.getenv("OPTHUB_PASS")
GMAIL_USER = os.getenv("GMAIL_USER", "bruno@compreoculos.com.br")
GMAIL_PASS = os.getenv("GMAIL_PASS")
RECIPIENTS = os.getenv("RECIPIENTS", "bruno@compreoculos.com.br,brunoera@gmail.com,comercial@opthub.com.br,maysa.gpanzuto@franqueadolinx.com.br")

STATUS_FILE = "StatusModeration.json"
TXT_FILE = "Clientes_Pendentes.txt"
XLSX_FILE = "Clientes_Pendentes.xlsx"
LOG_GETCUSTOMER = "log_getcustomer.txt"
LOG_EXECUCAO = "log_execucao.txt"
GETCUSTOMER_URL = "https://opthub.layer.core.dcg.com.br/v1/Profile/API.svc/web/GetCustomer"


def fetch_customer_email(customer_id, log_api):
    """Consulta o e-mail do cliente usando body numérico puro."""
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    email = None

    try:
        resp = requests.post(GETCUSTOMER_URL, headers=headers, json=customer_id,
                             auth=(OPTHUB_USER, OPTHUB_PASS), timeout=30)
        log_api.append(f"Consulta ID {customer_id} - Status {resp.status_code}")
        if resp.status_code == 200:
            data = resp.json()
            email = data.get("Email")
        else:
            log_api.append(f"⚠️ Resposta inesperada: {resp.text[:200]}")
    except Exception as e:
        log_api.append(f"❌ Erro ao consultar ID {customer_id}: {e}")

    return email


def autoajustar_colunas_excel(path):
    """Ajusta automaticamente a largura das colunas do Excel."""
    wb = load_workbook(path)
    ws = wb.active
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = max_len + 2
    wb.save(path)


def send_email_with_attachments(subject, body_text, attachments):
    """Envia o e-mail com texto e anexos (xlsx + logs)."""
    recipients = [r.strip() for r in RECIPIENTS.split(",") if r.strip()]
    msg = EmailMessage()
    msg["From"] = GMAIL_USER
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body_text)

    for file_path in attachments:
        if os.path.exists(file_path):
            with open(file_path, "rb") as f:
                maintype, subtype = ("application", "octet-stream")
                if file_path.endswith(".xlsx"):
                    subtype = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                msg.add_attachment(
                    f.read(),
                    maintype=maintype,
                    subtype=subtype,
                    filename=os.path.basename(file_path),
                )

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(GMAIL_USER, GMAIL_PASS)
        server.send_message(msg)


def main():
    log_exec = [f"==== LOG EXECUÇÃO - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} ====\n"]
    log_api = []

    try:
        if not os.path.exists(STATUS_FILE):
            log_exec.append("❌ Arquivo StatusModeration.json não encontrado.")
            raise FileNotFoundError("StatusModeration.json não encontrado.")

        log_exec.append("📥 Lendo StatusModeration.json...")
        with open(STATUS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)

        pendentes = []
        for customer in data.get("Result", []):
            customer_id = customer.get("CustomerID")
            customer_name = customer.get("CustomerName")

            for status in customer.get("ModerationStatus", []):
                if (
                    status.get("SellerAcceptanceStatus") == "approved"
                    and status.get("CustomerAcceptanceStatus") == "pending"
                ):
                    pendentes.append({"CustomerID": customer_id, "CustomerName": customer_name})
                    break

        log_exec.append(f"🧾 Clientes encontrados com Seller aprovado e Cliente pendente: {len(pendentes)}")

        # Consulta e-mails
        for c in pendentes:
            c["Email"] = fetch_customer_email(c["CustomerID"], log_api) or "Não encontrado"

        # Gera TXT e XLSX
        log_exec.append("💾 Gerando planilha Excel e arquivo TXT...")
        linhas = [
            "Clientes com Seller aprovado e Termo de Aceite pendente:\n",
            f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}\n",
            "-" * 60 + "\n",
        ]
        if not pendentes:
            linhas.append("Nenhum cliente pendente encontrado.\n")
        else:
            for c in pendentes:
                linhas.append(f"ID: {c['CustomerID']} | Nome: {c['CustomerName']} | Email: {c['Email']}\n")

        with open(TXT_FILE, "w", encoding="utf-8") as f:
            f.write("".join(linhas))

        df = pd.DataFrame(pendentes)
        df.to_excel(XLSX_FILE, index=False)
        autoajustar_colunas_excel(XLSX_FILE)
        log_exec.append(f"✅ Arquivos gerados: {TXT_FILE}, {XLSX_FILE}")

        # Gera logs
        with open(LOG_GETCUSTOMER, "w", encoding="utf-8") as lf:
            lf.write("\n".join(log_api))
        with open(LOG_EXECUCAO, "w", encoding="utf-8") as lf:
            lf.write("\n".join(log_exec))

        # Envia o e-mail
        subject = "[Opthub] Clientes com Termo de Aceite Pendente"
        send_email_with_attachments(subject, "".join(linhas),
                                    [XLSX_FILE, LOG_GETCUSTOMER, LOG_EXECUCAO])

        log_exec.append("📤 E-mail enviado com sucesso.")
        print(f"✅ E-mail enviado com {len(pendentes)} clientes pendentes.")
    except Exception as e:
        log_exec.append(f"❌ Erro inesperado: {e}")
    finally:
        with open(LOG_EXECUCAO, "w", encoding="utf-8") as lf:
            lf.write("\n".join(log_exec))


if __name__ == "__main__":
    main()
