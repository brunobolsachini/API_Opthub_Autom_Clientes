import os
import json
import smtplib
import ssl
from email.message import EmailMessage
from datetime import datetime, timezone
import requests
import pandas as pd
from textwrap import shorten

OPTHUB_URL = "https://opthub.layer.core.dcg.com.br/v1/Profile/API.svc/web/GetStatusModerationCustomerMarketplace"

def ensure_out_dir(path: str):
    os.makedirs(path, exist_ok=True)

def fetch_status_moderation(username: str, password: str) -> dict:
    body = {"Page": {"PageIndex": 0, "PageSize": 10000}}
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    resp = requests.post(OPTHUB_URL, headers=headers, json=body, auth=(username, password), timeout=120)
    resp.raise_for_status()
    try:
        return resp.json()
    except json.JSONDecodeError:
        return {"raw_text": resp.text}

def normalize_payload_to_dataframe(payload: dict) -> pd.DataFrame:
    if not isinstance(payload, dict):
        return pd.DataFrame([payload])
    for key in ["model", "data", "items", "results", "value"]:
        if key in payload and isinstance(payload[key], list):
            return pd.DataFrame(payload[key])
    return pd.json_normalize(payload)

def save_files(payload: dict, out_dir: str, base_name: str):
    df = normalize_payload_to_dataframe(payload)
    now = datetime.now(timezone.utc).astimezone()
    stamp = now.strftime("%Y%m%d_%H%M%S")

    # Caminhos
    excel_path = os.path.join(out_dir, f"{base_name}_{stamp}.xlsx")
    json_path  = os.path.join(out_dir, f"{base_name}.json")
    txt_path   = os.path.join(out_dir, f"{base_name}.txt")

    # Excel (histórico)
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="StatusModeration")

    # JSON (usado por outro endpoint futuramente)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    # TXT (resumo para leitura humana)
    resumo = [f"Atualizado em: {now.strftime('%d/%m/%Y %H:%M:%S')}"]
    resumo.append(f"Total de registros: {len(df)}\n")
    preview = df.head(10).to_string(index=False)
    resumo.append("Prévia dos primeiros registros:\n" + preview)
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(resumo))

    return excel_path, json_path, txt_path, len(df)

def send_email(gmail_user, gmail_pass, recipients_csv, subject, body, attachment_path):
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

    ensure_out_dir(out_dir)

    payload = fetch_status_moderation(opthub_user, opthub_pass)
    excel_path, json_path, txt_path, qtd = save_files(payload, ".", base_name)

    subject = f"[Opthub] Status Moderation - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    body = (
        "Olá, Bruno!\n\n"
        "Segue em anexo o arquivo Excel com os dados do endpoint "
        "GetStatusModerationCustomerMarketplace.\n\n"
        f"Foram encontrados {qtd} registros.\n\n"
        "Além disso, os arquivos 'StatusModeration.json' e 'StatusModeration.txt' "
        "foram atualizados na raiz do repositório para controle e uso interno.\n\n"
        "— Automação
