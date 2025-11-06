import os
import json
import requests
import smtplib
import ssl
from email.message import EmailMessage
from datetime import datetime

# --- CONFIGURAÇÕES GERAIS ---

OPTHUB_USER = os.getenv("OPTHUB_USER", "bruno.opthub")
OPTHUB_PASS = os.getenv("OPTHUB_PASS")
GMAIL_USER = os.getenv("GMAIL_USER", "bruno@compreoculos.com.br")
GMAIL_PASS = os.getenv("GMAIL_PASS")
RECIPIENTS = os.getenv("RECIPIENTS", "bruno@compreoculos.com.br,brunoera@gmail.com")

STATUS_FILE = "StatusModeration.json"
OUT_TXT = "Clientes_Pendentes.txt"
GETCUSTOMER_URL = "https://opthub.layer.core.dcg.com.br/v1/Profile/API.svc/web/GetCustomer"


# --- FUNÇÕES AUXILIARES ---

def fetch_customer_email(customer_id):
    """Consulta o e-mail de um cliente pelo ID."""
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    body = {"CustomerID": customer_id}

    try:
        resp = requests.post(GETCUSTOMER_URL, headers=headers, json=body, auth=(OPTHUB_USER, OPTHUB_PASS), timeout=30)
        resp.raise_for_status()
        data = resp.json()
        return data.get("Email") or data.get("Result", {}).get("Email")
    except Exception as e:
        print(f"⚠️ Erro ao buscar e-mail do CustomerID {customer_id}: {e}")
        return None


def send_email(subject, body_text):
    """Envia o e-mail com a lista de clientes pendentes."""
    recipients = [r.strip() for r in RECIPIENTS.split(",") if r.strip()]

    msg = EmailMessage()
    msg["From"] = GMAIL_USER
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body_text)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(GMAIL_USER, GMAIL_PASS)
        server.send_message(msg)


# --- LÓGICA PRINCIPAL ---

def main():
    if not os.path.exists(STATUS_FILE):
        print(f"❌ Arquivo {STATUS_FILE} não encontrado.")
        return

    with open(STATUS_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Lista de clientes com seller aprovado e cliente pendente
    pendentes = []
    for customer in data.get("Result", []):
        customer_id = customer.get("CustomerID")
        customer_name = customer.get("CustomerName")
        moderations = customer.get("ModerationStatus", [])

        for status in moderations:
            if status.get("SellerAcceptanceStatus") == "approved" and status.get("CustomerAcceptanceStatus") == "pending":
                pendentes.append({"CustomerID": customer_id, "CustomerName": customer_name})
                break  # adiciona só uma vez por cliente

    # Buscar e-mails
    for c in pendentes:
        email = fetch_customer_email(c["CustomerID"])
        c["Email"] = email or "Não encontrado"

    # Gerar texto
    linhas = []
    linhas.append("Clientes com Seller aprovado e Termo de Aceite pendente:\n")
    linhas.append(f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}\n")
    linhas.append("-" * 60 + "\n")

    if not pendentes:
        linhas.append("Nenhum cliente pendente encontrado.\n")
    else:
        for c in pendentes:
            linhas.append(f"ID: {c['CustomerID']} | Nome: {c['CustomerName']} | Email: {c['Email']}\n")

    texto_final = "".join(linhas)

    # Salvar TXT local
    with open(OUT_TXT, "w", encoding="utf-8") as f:
        f.write(texto_final)

    # Enviar por e-mail
    subject = "[Opthub] Clientes com Termo de Aceite Pendente"
    send_email(subject, texto_final)

    print(f"✅ Lista gerada e e-mail enviado com {len(pendentes)} clientes pendentes.")


if __name__ == "__main__":
    main()
