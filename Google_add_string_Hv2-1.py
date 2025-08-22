#!/usr/bin/env python3
"""
mega_script.py
==============
End-to-end помощник для кейса Help Global.

Алгоритм
1.  Принимает invoice_number (целое).
2.  Формирует строку-папку «000XXXXX».
3.  На Google Drive (CO ICF HELP GLOBAL) ищет подпапку-кейс.
4.  Скачивает из неё:
      • Grant Agreement*.pdf
      • Invoice*.pdf → сохраняет локально как «Invoice <num>.pdf».
5.  Создаёт/находит локальную папку вида
         «Нова YY-MM XXX <целая сумма> №<num> Хелп»
    (встроенная функция create_case_dir), и
    перемещает в неё оба PDF.
6.  Пропускает Invoice через parser_Invoicev2.py, получает
      {date, amount, case_descr}.
7.  Добавляет строку в Google-таблицу
      «Учёт грантов 4UA» / лист «Help Global».

Требуется
    client_secret.json  – рядом со скриптом
    token.json          – создаётся автоматически
    parser_Invoicev2.py – F:\Служебная\Волонтерство 4UA\ChatGPT\Автоматизация\Functions
"""

from __future__ import annotations
import argparse
import io
import os
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

# ────────── GOOGLE CONSTANTS ──────────
ROOT_FOLDER_ID = "1KNvnzuBL_froKQs-JVd8TVoGDMtL4-wx"
SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]
SHEET_ID = "1Pr0rb89ZIsy2qiBkZAySPuEf9B_zdn58CDurhtnQm0U"
SHEET_RANGE = "Help Global!A:L"
PAGE_SIZE = 1000

# ────────── ПУТИ К ВНУТРЕННИМ СКРИПТАМ ──────────
FUNCTIONS_DIR = Path(
    r"F:\Служебная\Волонтерство 4UA\ChatGPT\Автоматизация\Functions"
)
PARSER_FILE = FUNCTIONS_DIR / "parser_Invoicev2.py"

# ────────── ДИНАМИЧЕСКИЙ ИМПОРТ parser_Invoicev2 ──────────
if not PARSER_FILE.exists():
    sys.exit(f"❌ parser_Invoicev2.py не найден по пути {PARSER_FILE}")
spec = _import_util.spec_from_file_location("parser_Invoicev2", PARSER_FILE)
parser_Invoice = _import_util.module_from_spec(spec)  # type: ignore
spec.loader.exec_module(parser_Invoice)               # type: ignore

# ────────── SIMPLE FOLDER-CREATOR (адаптировано) ──────────
CONSTANT_TEXT = "XXX"
FORBIDDEN = r'<>:"/\\|?*'


def sanitize(name: str) -> str:
    """Удаляет запрещённые символы из имени папки."""
    name = re.sub(f"[{re.escape(FORBIDDEN)}]", "", name)
    return name.strip(" .")


def create_case_dir(date_iso: str, amount: float, invoice_number: int) -> Path:
    """Создаёт папку «Нова YY-MM XXX <amount_int> №<num> Хелп» и возвращает путь."""
    dt = datetime.fromisoformat(date_iso).date()
    yy_mm = dt.strftime("%y-%m")
    amount_int = int(round(amount))
    dir_name = f"Нова {yy_mm} {CONSTANT_TEXT} {amount_int} №{invoice_number} Хелп"
    dir_path = Path.cwd() / sanitize(dir_name)
    dir_path.mkdir(exist_ok=True)
    return dir_path


# ────────── OAUTH ──────────
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


# ────────── DRIVE HELPERS ──────────
def find_case_folder(service, case_folder_name: str) -> Optional[str]:
    query = (
        f"'{ROOT_FOLDER_ID}' in parents and "
        "mimeType='application/vnd.google-apps.folder' and "
        f"name='{case_folder_name}' and trashed=false"
    )
    resp = (
        service.files()
        .list(
            q=query,
            fields="files(id)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            corpora="allDrives",
            pageSize=1,
        )
        .execute()
    )
    files = resp.get("files", [])
    return files[0]["id"] if files else None


def find_first_pdf(service, parent_id: str, prefix: str) -> Optional[Dict]:
    query = (
        f"'{parent_id}' in parents and "
        "mimeType='application/pdf' and trashed=false"
    )
    resp = (
        service.files()
        .list(
            q=query,
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            corpora="allDrives",
            pageSize=PAGE_SIZE,
        )
        .execute()
    )
    for f in resp.get("files", []):
        if f["name"].lower().startswith(prefix.lower()):
            return f
    return None


FORBIDDEN = r'<>:"/\\|?*'

def _clean_filename(name: str) -> str:
    """Убирает запрещённые символы Windows и лишние пробелы/точки."""
    import re
    name = re.sub(f"[{re.escape(FORBIDDEN)}]", "_", name)
    return name.strip(" .")

def download_pdf(service, file_meta: Dict, local_name: str) -> Path:
    # 1. дописываем .pdf, если нужно
    if not local_name.lower().endswith(".pdf"):
        local_name += ".pdf"
    # 2. чистим имя
    local_name = _clean_filename(local_name)

    request = service.files().get_media(fileId=file_meta["id"])
    with io.FileIO(local_name, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            downloader.next_chunk()
            done = downloader._done  # type: ignore
    return Path(local_name).resolve()


# ────────── SHEETS APPEND ──────────
def append_row(sheets_service, row):
    sheets_service.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range=SHEET_RANGE,
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [row]},
    ).execute()


# ────────── MAIN ──────────
def main():
    parser = argparse.ArgumentParser(description="Скачивает файлы, создаёт папку и дописывает строку.")
    parser.add_argument("invoice_number", type=int)
    args = parser.parse_args()

    invoice_num = args.invoice_number
    invoice_str_padded = f"{invoice_num:08d}"

    creds = get_credentials()
    drive = build("drive", "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)

    # 1. Папка-кейс на Drive
    folder_id = find_case_folder(drive, invoice_str_padded)
    if not folder_id:
        sys.exit("❌ Папка-кейс на Drive не найдена.")

    # 2. Скачиваем Invoice + Grant Agreement
    inv_meta = find_first_pdf(drive, folder_id, "Invoice")
    if not inv_meta:
        sys.exit("❌ Invoice*.pdf не найден.")
    inv_local = download_pdf(drive, inv_meta, f"Invoice {invoice_num}.pdf")
    print(f"⬇️  Invoice → {inv_local.name}")

    ga_meta = find_first_pdf(drive, folder_id, "Grant Agreement")
    ga_local = None
    if ga_meta:
        ga_local = download_pdf(drive, ga_meta, ga_meta["name"])
        print(f"⬇️  Grant Agreement → {ga_local.name}")
    else:
        print("⚠️  Grant Agreement не найден, продолжаем без него.")

    # 3. Парсим Invoice
    try:
        info: Dict = parser_Invoice.parse_invoice(str(inv_local))
        date_iso = info["date"]          # YYYY-MM-DD
        amount = info["amount"]          # float / decimal
        case_descr = info.get("case_descr", "")
    except Exception as e:
        sys.exit(f"❌ Ошибка парсинга Invoice: {e}")

    # 4. Создаём / находим локальную папку
    target_dir = None
    for p in Path.cwd().iterdir():
        if p.is_dir() and str(invoice_num) in p.name:
            target_dir = p
            break
    if not target_dir:
        target_dir = create_case_dir(date_iso, amount, invoice_num)
        print(f"📁 Создана новая папка: {target_dir.name}")
    else:
        print(f"📁 Используем существующую папку: {target_dir.name}")

    # 5. Перемещаем PDF
    shutil.move(str(inv_local), target_dir / inv_local.name)
    if ga_local:
        shutil.move(str(ga_local), target_dir / ga_local.name)

    # 6. Append to Google Sheet
    row = [
        date_iso,               # A
        "",                     # B
        "",                     # C
        invoice_num,            # D
        case_descr,             # E
        amount,                 # F
        "", "", "", "", "",     # G-K
        "хер",                  # L
    ]
    append_row(sheets, row)
    print("✅ Строка успешно добавлена в «Help Global».")


if __name__ == "__main__":
    try:
        main()
    except HttpError as e:
        sys.exit(f"Google API error: {e}")
    except KeyboardInterrupt:
        sys.exit("\n⏹️  Прервано пользователем.")
