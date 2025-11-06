import os
import json
import smtplib
import ssl
from email.message import EmailMessage
from datetime import datetime, timezone
import requests

OPTHUB_URL = "https://opthub.layer.core.dcg.com.br/v1/Profile/API.svc/web/GetStatusModerationCustomerMarketplace"

def ensure_out_dir(path: str):
    os.makedirs(path, exist_ok=True)

def fetch_status_moderation(username: str, password: str) -> dict:
    body = {
        "Page": {
            "PageIndex": 0,
            "PageSize": 10000
        }
    }
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json"
    }

    resp = requests.post(
        OPTHUB_URL,
        headers=headers,
        json=body,
        auth=(username, password),  # Basic Auth – ajustamos se a Linx exigir outro método
        timeout=120
    )
    resp.raise_for_status()
    try:
        return resp.json()
    except json.JSONDecodeError:
        return {"raw_text": resp.text}

def save_json(payload: dict, out_dir: str, base_name: str) -> str:
    now = datetime.now(timezone.utc).astimezone()
    stamp = now.strftime("%Y%m%d_%H%M%S")
    filepath = os.path.join(out_dir, f"{base_name}_{stamp}.json")
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return filepath

def summarize_payload(payload: dict) -> str:
    total_items = None
    candidates = []
    if isinstance(payload, dict):
        for k in ("items", "data", "results", "model", "value"):
            if k in payload:
                candidates.append((k, payload[k]))
    if candidates:
        for _, v in candidates:
            if isinstance(v, list):
                total_items = len(v)
                break
    if total_items is None:
        for v in payload.values() if isinstance(payload, dict) else []:
            if isinstance(v, list):
                total_items = len(v)
                break
    if total_items is None:
        total_items = 1 if payload else 0
    return f"Registros encontrados (estimativa): {total_items}"

def send_email(gmail_user: str, gmail_pass: str, recipients_csv: str, subject: str, body: str, attachment_path: str):
    recipients = [r.strip() for r in recipients_csv.split(",") if r.strip()]
    msg = EmailMessage()
    msg["From"] = gmail_user
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body)

    with open(attachment_path, "rb") as f:
        data = f.read()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="json",
        filename=os.path.basename(attachment_path)
    )

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(gmail_user, gmail_pass)
        server.send_message(msg)

def main():
    out_dir = os.getenv("OUTPUT_DIR", "out")
    base_name = os.getenv("STATUS_FILE_BASENAME", "StatusModeration")

    opthub_user = os.getenv("OPTHUB_USER")
    opthub_pass = os.getenv("OPTHUB_PASS")

    gmail_user = os.getenv("GMAIL_USER")
    gmail_pass = os.getenv("GMAIL_PASS")
    recipients = os.getenv("RECIPIENTS")

    if not all([opthub_user, opthub_pass, gmail_user, gmail_pass, recipients]):
        raise RuntimeError("Faltam variáveis de ambiente obrigatórias (OPTHUB_*, GMAIL_*, RECIPIENTS).")

    ensure_out_dir(out_dir)
    payload = fetch_status_moderation(opthub_user, opthub_pass)
    json_path = save_json(payload, out_dir, base_name)
    resumo = summarize_payload(payload)

    agora_brt = datetime.now().strftime("%d/%m/%Y %H:%M")
    subject = f"[Opthub] Status Moderation - {agora_brt}"
    body = (
        "Olá, Bruno!\n\n"
        "Segue em anexo o retorno bruto do endpoint GetStatusModerationCustomerMarketplace.\n"
        "Quando você me passar os campos exatos do 'Termo de Aceite pendente', "
        "eu filtro e mando somente nome + e-mail dos pendentes no corpo do e-mail.\n\n"
        f"{resumo}\n\n"
        f"Arquivo: {os.path.basename(json_path)}\n"
        "— Automação API_Opthub_Autom_Clientes"
    )

    send_email(gmail_user, gmail_pass, recipients, subject, body, json_path)
    print(f"OK - Arquivo salvo em {json_path} e e-mail enviado para {recipients}")

if __name__ == "__main__":
    main()
