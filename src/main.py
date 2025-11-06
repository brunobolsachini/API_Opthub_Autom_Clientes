import os
import json
import smtplib
import ssl
from email.message import EmailMessage
from datetime import datetime, timezone
import requests
import pandas as pd

OPTHUB_URL = "https://opthub.layer.core.dcg.com.br/v1/Profile/API.svc/web/GetStatusModerationCustomerMarketplace"

def ensure_out_dir(path: str):
    os.makedirs(path, exist_ok=True)

def fetch_status_moderation(username: str, password: str) -> dict:
    """Consulta o endpoint e retorna o JSON bruto."""
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
        auth=(username, password),  # Basic Auth – ajustamos depois se mudar
        timeout=120
    )
    resp.raise_for_status()
    try:
        return resp.json()
    except json.JSONDecodeError:
        return {"raw_text": resp.text}

def normalize_payload_to_dataframe(payload: dict) -> pd.DataFrame:
    """
    Tenta encontrar a lista de registros dentro do JSON retornado.
    Adapta automaticamente se vier em 'model', 'data', 'items', etc.
    """
    if not isinstance(payload, dict):
        return pd.DataFrame([payload])

    for key in ["model", "data", "items", "results", "value"]:
        if key in payload and isinstance(payload[key], list):
            return pd.DataFrame(payload[key])

    # Se não achar lista, converte tudo
    return pd.json_normalize(payload)

def save_excel(payload: dict, out_dir: str, base_name: str) -> str:
    """Salva o resultado em formato Excel."""
    df = normalize_payload_to_dataframe(payload)

    now = datetime.now(timezone.utc).astimezone()
    stamp = now.strftime("%Y%m%d_%H%M%S")
    filepath = os.path.join(out_dir, f"{base_name}_{stamp}.xlsx")

    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="StatusModeration")

    return filepath, len(df)

def send_email(gmail_user: str, gmail_pass: str, recipients_csv: str, subject: str, body: str, attachment_path: str):
    recipients = [r.strip() for r in recipients_csv.split(",") if r.strip()]
    msg = EmailMessage()
    msg["From"] = gmail_user
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body)

    # Anexar arquivo XLSX
    with open(attachment_path, "rb") as f:
        data = f.read()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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

    # 1️⃣ Buscar dados
    payload = fetch_status_moderation(opthub_user, opthub_pass)

    # 2️⃣ Salvar Excel
    excel_path, qtd = save_excel(payload, out_dir, base_name)

    # 3️⃣ Montar e enviar e-mail
    agora_brt = datetime.now().strftime("%d/%m/%Y %H:%M")
    subject = f"[Opthub] Status Moderation - {agora_brt}"
    body = (
        "Olá, Bruno!\n\n"
        "Segue em anexo o retorno do endpoint GetStatusModerationCustomerMarketplace "
        "em formato Excel (.xlsx).\n\n"
        f"Foram encontrados {qtd} registros.\n\n"
        "Assim que você me passar os campos correspondentes a 'Aprovação do Seller' "
        "e 'Aprovação do Cliente', aplico o filtro e envio apenas os pendentes.\n\n"
        f"Arquivo: {os.path.basename(excel_path)}\n"
        "— Automação API_Opthub_Autom_Clientes"
    )

    send_email(gmail_user, gmail_pass, recipients, subject, body, excel_path)
    print(f"OK - Arquivo salvo em {excel_path} e e-mail enviado para {recipients}")

if __name__ == "__main__":
    main()
