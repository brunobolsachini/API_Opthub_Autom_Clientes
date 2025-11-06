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
GETCUSTOMER_URL = "https://opthub.layer.core.dcg.com.br/v1/Profile/API.svc/web/GetCustomer"

def fetch_customer_email(customer_id):
    """Consulta o e-mail do cliente no GetCustomer."""
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    body = {"CustomerID": customer_id}
    try:
        resp = requests.post(GETCUSTOMER_URL, headers=headers, json=body, auth=(OPTHUB_USER, OPTHUB_PASS), timeout=30)
        resp.raise_for_status()
        data = resp.json()
        return data.get("Email")  # campo direto no JSON
    except Exception as e:
        print(f"⚠️ Erro ao buscar e-mail do CustomerID {customer_id}: {e}")
        return None

def send_email_with_attachment(subject, body_text, attachment_path):
    """Envia o e-mail com texto + anexo Excel."""
    recipients = [r.strip() for r in RECIPIENTS.split(",") if r.strip()]
    msg = EmailMessage()
    msg["From"] = GMAIL_USER
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body_text)

    with open(attachment_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(attachment_path),
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
    for customer in data.get("Result", []):
        customer_id = customer.get("CustomerID")
        customer_name = customer.get("CustomerName")
        for status in customer.get("ModerationStatus", []):
            if status.get("SellerAcceptanceStatus") == "approved" and status.get("CustomerAcceptanceStatus") == "pending":
                pendentes.append({"CustomerID": customer_id, "CustomerName": customer_name})
                break

    for c in pendentes:
        c["Email"] = fetch_customer_email(c["CustomerID"]) or "Não encontrado"

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

    # Enviar e-mail com anexo Excel
    subject = "[Opthub] Clientes com Termo de Aceite Pendente"
    send_email_with_attachment(subject, "".join(linhas), XLSX_FILE)

    print(f"✅ E-mail enviado e arquivos salvos: {TXT_FILE}, {XLSX_FILE}")

if __name__ == "__main__":
    main()
