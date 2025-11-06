import os
import json
from datetime import datetime, timezone
import requests
import pandas as pd

OPTHUB_URL = "https://opthub.layer.core.dcg.com.br/v1/Profile/API.svc/web/GetStatusModerationCustomerMarketplace"

def ensure_out_dir(path: str):
    os.makedirs(path, exist_ok=True)

def fetch_status_moderation(username: str, password: str) -> dict:
    """Consulta o endpoint e retorna o JSON bruto."""
    body = {"Page": {"PageIndex": 0, "PageSize": 10000}}
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    resp = requests.post(OPTHUB_URL, headers=headers, json=body, auth=(username, password), timeout=120)
    resp.raise_for_status()
    return resp.json()

def normalize_payload_to_dataframe(payload: dict) -> pd.DataFrame:
    if not isinstance(payload, dict):
        return pd.DataFrame([payload])
    for key in ["model", "data", "items", "results", "value", "Result"]:
        if key in payload and isinstance(payload[key], list):
            return pd.DataFrame(payload[key])
    return pd.json_normalize(payload)

def save_excel(payload: dict, out_dir: str, base_name: str) -> tuple[str, int]:
    df = normalize_payload_to_dataframe(payload)
    now = datetime.now(timezone.utc).astimezone()
    stamp = now.strftime("%Y%m%d_%H%M%S")
    filepath = os.path.join(out_dir, f"{base_name}_{stamp}.xlsx")
    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="StatusModeration")
    return filepath, len(df)

def main():
    out_dir = os.getenv("OUTPUT_DIR", "out")
    base_name = os.getenv("STATUS_FILE_BASENAME", "StatusModeration")
    opthub_user = os.getenv("OPTHUB_USER")
    opthub_pass = os.getenv("OPTHUB_PASS")

    ensure_out_dir(out_dir)
    payload = fetch_status_moderation(opthub_user, opthub_pass)
    excel_path, qtd = save_excel(payload, out_dir, base_name)

    json_path_repo = "StatusModeration.json"
    txt_path_repo = "StatusModeration.txt"
    with open(json_path_repo, "w", encoding="utf-8") as jf:
        json.dump(payload, jf, ensure_ascii=False, indent=2)
    with open(txt_path_repo, "w", encoding="utf-8") as tf:
        tf.write(f"Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}\n")
        tf.write(f"Registros encontrados: {qtd}\n\n")
        tf.write(json.dumps(payload, indent=2, ensure_ascii=False))

    print(f"✅ StatusModeration atualizado e salvo em {json_path_repo} ({qtd} registros).")

if __name__ == "__main__":
    main()
