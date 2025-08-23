#!/usr/bin/env python3
"""
mega_script.py  (обновлён 24-Aug-2025)
=====================================
– Берёт invoice_number  →  скачивает новейшие PDF «Invoice» и «Grant Agreement»
  из папки-кейса на Google Drive, создаёт/находит локальную папку случая и
  дописывает строку в Google-Sheets (лист “Help Global”).

Требуется:
    • client_secret.json — рядом со скриптом
    • token.json         — создаётся автоматически
    • parser_Invoicev2.py (лежит в …\Functions)
"""

from __future__ import annotations
import argparse
import io
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional
from importlib import util as _import_util

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError

# ─────────────── GOOGLE CONSTANTS ───────────────
ROOT_FOLDER_ID = "1KNvnzuBL_froKQs-JVd8TVoGDMtL4-wx"
SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]
SHEET_ID = "1Pr0rb89ZIsy2qiBkZAySPuEf9B_zdn58CDurhtnQm0U"
SHEET_RANGE = "Help Global!A:L"

# ─────────────── PATHS TO HELPERS ───────────────
FUNCTIONS_DIR = Path(
    r"F:\Служебная\Волонтерство 4UA\ChatGPT\Автоматизация\Functions"
)
PARSER_FILE = FUNCTIONS_DIR / "parser_Invoicev2.py"

# ─────────────── DYNAMIC IMPORT ───────────────
if not PARSER_FILE.exists():
    sys.exit(f"❌ parser_Invoicev2.py не найден ({PARSER_FILE})")
spec = _import_util.spec_from_file_location("parser_Invoicev2", PARSER_FILE)
parser_Invoice = _import_util.module_from_spec(spec)          # type: ignore
spec.loader.exec_module(parser_Invoice)                       # type: ignore

# ─────────────── HELPERS FOR FOLDER NAME ───────────────
CONSTANT_TEXT = "XXX"
FORBIDDEN = r'<>:"/\\|?*'


def sanitize(name: str) -> str:
    return re.sub(f"[{re.escape(FORBIDDEN)}]", "_", name).strip(" .")


def create_case_dir(date_iso: str, amount: float, invoice_number: int) -> Path:
    yy_mm = datetime.fromisoformat(date_iso).strftime("%y-%m")
    dir_name = f"Нова {yy_mm} {CONSTANT_TEXT} {int(round(amount))} №{invoice_number} Хелп"
    path = Path.cwd() / sanitize(dir_name)
    path.mkdir(exist_ok=True)
    return path


# ─────────────── OAUTH ───────────────
def get_credentials() -> Credentials:
    token_path = Path("token.json")
    creds: Optional[Credentials] = None

    if token_path.exists():
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                creds = None

    if not creds or not creds.valid:
        if not Path("client_secret.json").exists():
            sys.exit("❌ client_secret.json не найден.")
        flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", SCOPES)
        creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return creds


# ─────────────── DRIVE FUNCTIONS ───────────────
def find_case_folder(service, folder_name: str) -> Optional[str]:
    query = (
        f"'{ROOT_FOLDER_ID}' in parents and "
        "mimeType='application/vnd.google-apps.folder' and "
        f"name='{folder_name}' and trashed=false"
    )
    resp = service.files().list(
        q=query,
        fields="files(id)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora="allDrives",
        pageSize=1,
    ).execute()
    files = resp.get("files", [])
    return files[0]["id"] if files else None


def find_latest_pdf(service, parent_id: str, prefix: str) -> Optional[Dict]:
    """Возвращает самый свежий PDF в папке parent_id, начинающийся с prefix."""
    resp = service.files().list(
        q=f"'{parent_id}' in parents and mimeType='application/pdf' and trashed=false",
        orderBy="modifiedTime desc",
        fields="files(id, name, modifiedTime)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora="allDrives",
        pageSize=100,
    ).execute()

    for f in resp.get("files", []):
        if f["name"].lower().startswith(prefix.lower()):
            return f
    return None


def clean_filename(name: str) -> str:
    name = re.sub(f"[{re.escape(FORBIDDEN)}]", "_", name)
    return name.strip(" .")


def download_pdf(service, file_meta: Dict, desired_name: str) -> Path:
    if not desired_name.lower().endswith(".pdf"):
        desired_name += ".pdf"
    desired_name = clean_filename(desired_name)

    request = service.files().get_media(fileId=file_meta["id"])
    with io.FileIO(desired_name, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return Path(desired_name).resolve()


# ─────────────── SHEETS APPEND ───────────────
def append_row(service, row):
    service.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range=SHEET_RANGE,
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [row]},
    ).execute()


# ─────────────── MAIN ───────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("invoice_number", type=int, help="Номер инвойса / кейса")
    args = parser.parse_args()

    invoice_num = args.invoice_number
    padded = f"{invoice_num:08d}"

    creds = get_credentials()
    drive = build("drive", "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)

    # 1. Папка-кейс
    folder_id = find_case_folder(drive, padded)
    if not folder_id:
        sys.exit("❌ Папка-кейс не найдена.")

    # 2. PDF файлы (самые свежие)
    inv_meta = find_latest_pdf(drive, folder_id, "Invoice")
    ga_meta = find_latest_pdf(drive, folder_id, "Grant Agreement")
    if not inv_meta:
        sys.exit("❌ Invoice*.pdf не найден.")

    inv_local = download_pdf(drive, inv_meta, f"Invoice {invoice_num}.pdf")
    print(f"⬇️  Invoice  → {inv_local.name}")

    ga_local = None
    if ga_meta:
        ga_local = download_pdf(drive, ga_meta, f"Grant Agreement {invoice_num}.pdf")
        print(f"⬇️  Grant Agreement → {ga_local.name}")
    else:
        print("⚠️  Grant Agreement не найден, продолжаем без него.")

    # 3. Парсим Invoice
    try:
        info: Dict = parser_Invoice.parse_invoice(str(inv_local))
        date_iso = info["date"]
        amount = info["amount"]
        case_descr = info.get("case_descr", "")
    except Exception as e:
        sys.exit(f"❌ Ошибка парсинга Invoice: {e}")

    # 4. Папка на диске
    target = next((p for p in Path.cwd().iterdir()
                   if p.is_dir() and str(invoice_num) in p.name), None)
    if not target:
        target = create_case_dir(date_iso, amount, invoice_num)
        print(f"📁 Создана папка: {target.name}")
    else:
        print(f"📁 Папка найдена: {target.name}")

    # 5. Перемещаем файлы
    shutil.move(str(inv_local), target / inv_local.name)
    if ga_local:
        shutil.move(str(ga_local), target / ga_local.name)

    # 6. Запись в Sheets
    append_row(sheets, [
        date_iso, "", "", invoice_num, case_descr, amount,
        "", "", "", "", "", "хер"
    ])
    print("✅ Строка добавлена в «Help Global».")


if __name__ == "__main__":
    try:
        main()
    except HttpError as e:
        sys.exit(f"Google API error: {e}")
    except KeyboardInterrupt:
        sys.exit("\n⏹️  Прервано пользователем.")
