import os
import json
import requests
import smtplib
import ssl
import pandas as pd
from email.message import EmailMessage
from datetime import datetime
from openpyxl import load_workbook
import time

# Configurações
OPTHUB_USER = os.getenv("OPTHUB_USER", "bruno.opthub")
OPTHUB_PASS = os.getenv("OPTHUB_PASS")
GMAIL_USER = os.getenv("GMAIL_USER", "bruno@compreoculos.com.br")
GMAIL_PASS = os.getenv("GMAIL_PASS")
RECIPIENTS = os.getenv("RECIPIENTS", "bruno@compreoculos.com.br,brunoera@gmail.com,comercial@opthub.com.br")

STATUS_FILE = "StatusModeration.json"
TXT_FILE = "Clientes_Pendentes.txt"
XLSX_FILE = "Clientes_Pendentes.xlsx"
LOG_GETCUSTOMER = "log_getcustomer.txt"
LOG_EXECUCAO = "log_execucao.txt"
GETCUSTOMER_URL = "https://opthub.layer.core.dcg.com.br/v1/Profile/API.svc/web/GetCustomer"


def log_step(log_exec, mensagem):
    """Adiciona uma linha ao log com timestamp."""
    hora = datetime.now().strftime("%H:%M:%S")
    log_exec.append(f"[{hora}] {mensagem}")


def fetch_customer_info(customer_id, log_api, log_exec):
    """Consulta o e-mail, telefone e celular do cliente."""
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    email = None
    phone = None
    cellphone = None
    inicio = time.time()

    try:
        resp = requests.post(GETCUSTOMER_URL, headers=headers, json=customer_id,
                             auth=(OPTHUB_USER, OPTHUB_PASS), timeout=30)
        duracao = time.time() - inicio
        log_api.append(f"Consulta ID {customer_id} - HTTP {resp.status_code} - {duracao:.2f}s")

        if resp.status_code == 200:
            data = resp.json()
            email = data.get("Email")
            phone = data.get("Phone")
            cellphone = data.get("CellPhone")
        else:
            log_api.append(f"⚠️ Resposta inesperada: {resp.text[:200]}")
    except Exception as e:
        log_api.append(f"❌ Erro ao consultar ID {customer_id}: {e}")

    log_step(log_exec, f"Consultado CustomerID {customer_id} -> Email: {email or 'N/A'} | Phone: {phone or 'N/A'} | Cell: {cellphone or 'N/A'}")
    return email, phone, cellphone


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


def send_email_with_attachments(subject, body_text, attachments, log_exec):
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
    inicio = time.time()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(GMAIL_USER, GMAIL_PASS)
        server.send_message(msg)
    duracao = time.time() - inicio
    log_step(log_exec, f"E-mail enviado para: {', '.join(recipients)} ({duracao:.2f}s)")


def main():
    inicio_total = time.time()
    log_exec = [f"==== LOG EXECUÇÃO - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} ===="]
    log_api = []

    try:
        if not os.path.exists(STATUS_FILE):
            log_step(log_exec, "❌ Arquivo StatusModeration.json não encontrado.")
            raise FileNotFoundError("StatusModeration.json não encontrado.")

        log_step(log_exec, "📥 Lendo StatusModeration.json...")
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
                    pendentes.append({
                        "CustomerID": customer_id,
                        "CustomerName": customer_name
                    })
                    break

        log_step(log_exec, f"🧾 Clientes identificados com pendência: {len(pendentes)}")

        # Consulta dados de contato
        for i, c in enumerate(pendentes, start=1):
            log_step(log_exec, f"➡️ ({i}/{len(pendentes)}) Consultando dados para {c['CustomerName']} (ID {c['CustomerID']})")
            email, phone, cellphone = fetch_customer_info(c["CustomerID"], log_api, log_exec)
            c["Email"] = email or "Não encontrado"
            c["Phone"] = phone or "Não informado"
            c["CellPhone"] = cellphone or "Não informado"

        # Gera TXT e XLSX
        log_step(log_exec, "💾 Gerando planilha Excel e arquivo TXT...")
        linhas = [
            "Clientes com Seller aprovado e Termo de Aceite pendente:\n",
            f"Atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}\n",
            "-" * 80 + "\n",
        ]
        if not pendentes:
            linhas.append("Nenhum cliente pendente encontrado.\n")
        else:
            for c in pendentes:
                linhas.append(f"ID: {c['CustomerID']} | Nome: {c['CustomerName']} | Email: {c['Email']} | Tel: {c['Phone']} | Cel: {c['CellPhone']}\n")

        with open(TXT_FILE, "w", encoding="utf-8") as f:
            f.write("".join(linhas))

        df = pd.DataFrame(pendentes)
        df.to_excel(XLSX_FILE, index=False)
        autoajustar_colunas_excel(XLSX_FILE)
        log_step(log_exec, "✅ Planilha Excel gerada e colunas autoajustadas.")

        # Gera logs
        with open(LOG_GETCUSTOMER, "w", encoding="utf-8") as lf:
            lf.write("\n".join(log_api))
        with open(LOG_EXECUCAO, "w", encoding="utf-8") as lf:
            lf.write("\n".join(log_exec))

        # Envia o e-mail com anexos
        subject = "[Opthub] Clientes com Termo de Aceite Pendente"
        send_email_with_attachments(subject, "".join(linhas),
                                    [XLSX_FILE, LOG_GETCUSTOMER, LOG_EXECUCAO], log_exec)

        log_step(log_exec, "📤 E-mail enviado com sucesso.")
        duracao_total = time.time() - inicio_total
        log_step(log_exec, f"✅ Execução finalizada em {duracao_total:.2f}s.")

    except Exception as e:
        log_step(log_exec, f"❌ Erro inesperado: {e}")
    finally:
        with open(LOG_EXECUCAO, "w", encoding="utf-8") as lf:
            lf.write("\n".join(log_exec))
        print("\n".join(log_exec))


if __name__ == "__main__":
    main()
