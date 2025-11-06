import os
import json
import requests
import smtplib
import ssl
import pandas as pd
from email.message import EmailMessage
from datetime import datetime

# Configurações
OPTHUB_USER = os.getenv("OPTHUB_USER", "bruno.opthub")
OPTHUB_PASS = os.getenv("OPTHUB_PASS")
GMAIL_USER = os.getenv("GMAIL_USER", "bruno@compreoculos.com.br")
GMAIL_PASS = os.getenv("GMAIL_PASS")
RECIPIENTS = os.getenv("RECIPIENTS", "bruno@compreoculos.com.br,brunoera@gmail.com")

STATUS_FILE = "StatusModeration.json"
TXT_FILE = "Clientes_Pendentes.txt"
XLSX_FILE = "Clientes_Pendentes.xlsx"
LOG_FILE = "log_getcustomer.txt"
GETCUSTOMER_URL = "https://opthub.layer.core.dcg.com.br/v1/Profile/API.svc/web/GetCustomer"


def try_getcustomer(customer_id, log_lines):
    """Tenta diferentes formatos de body até encontrar o Email."""
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    bodies_to_try = [
        {"CustomerID": customer_id},             # formato 1
        {"model": {"CustomerID": customer_id}},  # formato 2
        customer_id,                             # formato 3 - número puro
        str(customer_id)                         # formato 4 - string pura
    ]

    for body in bodies_to_try:
        try:
            resp = requests.post(
                GETCUSTOMER_URL,
                headers=headers,
                json=body,
                auth=(OPTHUB_USER, OPTHUB_PASS),
                timeout=30
            )
            log_lines.append(f"\n🟢 Tentando body: {json.dumps(body)}")
            log_lines.append(f"Status: {resp.status_code}")
            log_lines.append(f"Resposta (primeiros 400 chars): {resp.text[:400]}\n")

            if resp.status_code == 200:
                try:
                    data = resp.json()
                    email = data.get("Email")
                    if email:
                        return email
                except Exception as e:
                    log_lines.append(f"❌ Erro ao ler JSON: {e}")
        except Exception as e:
            log_lines.append(f"❌ Falha ao tentar body {body}: {e}")

    return None


def send_email_with_attachments(subject, body_text, attachments):
    """Envia o e-mail com texto e anexos (xlsx + log)."""
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
    if not os.path.exists(STATUS_FILE):
        print("❌ StatusModeration.json não encontrado.")
        return

    with open(STATUS_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    pendentes = []
    log_lines = [f"==== LOG EXECUÇÃO {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} ====\n"]

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

    for c in pendentes:
        email = try_getcustomer(c["CustomerID"], log_lines)
        c["Email"] = email or "Não encontrado"
        if not email:
            log_lines.append(f"⚠️ Sem e-mail para {c['CustomerName']} (ID {c['CustomerID']})")

    # Salvar TXT e XLSX
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

    pd.DataFrame(pendentes).to_excel(XLSX_FILE, index=False)
    with open(LOG_FILE, "w", encoding="utf-8") as lf:
        lf.write("\n".join(log_lines))

    # Enviar e-mail com anexos
    subject = "[Opthub] Clientes com Termo de Aceite Pendente"
    send_email_with_attachments(subject, "".join(linhas), [XLSX_FILE, LOG_FILE])

    print(f"✅ E-mail enviado e arquivos salvos: {TXT_FILE}, {XLSX_FILE}, {LOG_FILE}")


if __name__ == "__main__":
    main()
