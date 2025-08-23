# -*- coding: utf-8 -*-
"""Mega_Help Script
====================
Automation script that scans Gmail for completed DocuSign cases, keeps track of
case statuses in a local Excel file, downloads case documents from Google Drive,
parses Invoice PDFs, creates local case folders and appends rows to a Google
Sheet. After execution a single Telegram message with a log of performed actions
is sent.

The script is designed to be idempotent – repeated runs will not duplicate
entries in Excel or send repeated Telegram messages for already processed
cases.
"""

from __future__ import annotations

import base64
import io
import json
import re
import shutil
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional

import argparse
import logging

import pandas as pd
import requests
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Paths and constants
ROOT_DIR = Path.cwd()
CASES_XLSX = ROOT_DIR / "cases_status.xlsx"
SEEN_JSON = ROOT_DIR / "seen_cases.json"
FUNCTIONS_DIR = Path(
    r"F:\Служебная\Волонтерство 4UA\ChatGPT\Автоматизация\Functions"
)
PARSER_FILE = FUNCTIONS_DIR / "parser_Invoicev2.py"

# Telegram
TELEGRAM_TOKEN = "<YOUR_TOKEN>"
TELEGRAM_CHAT_ID = 0  # replace with real chat id
TELEGRAM_URL = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"

# Gmail and Drive
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
DRIVE_SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]
ROOT_FOLDER_ID = "1KNvnzuBL_froKQs-JVd8TVoGDMtL4-wx"
SHEET_ID = "1Pr0rb89ZIsy2qiBkZAySPuEf9B_zdn58CDurhtnQm0U"
SHEET_RANGE = "Help Global!A:L"

# Regex patterns
RE_APPROVED = re.compile(r"Approved case\s*(\d{8})")
RE_CASE_FOLDER = "{0:08d}"
FORBIDDEN = r'<>:"/\\|?*'

# ---------------------------------------------------------------------------
# Data models
@dataclass
class CaseFlags:
    invoice_downloaded: bool = False
    grant_downloaded: bool = False
    parsed: bool = False
    error: Optional[str] = None


# ---------------------------------------------------------------------------
# Utility functions

def load_seen() -> Dict:
    if SEEN_JSON.exists():
        with open(SEEN_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
            logger.debug("Loaded seen data from %s", SEEN_JSON)
            return data
    logger.debug("No seen data found, starting fresh")
    return {"messages": [], "cases": {}}


def save_seen(data: Dict) -> None:
    with open(SEEN_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    logger.debug("Saved seen data to %s", SEEN_JSON)


def ensure_cases_excel() -> pd.DataFrame:
    if not CASES_XLSX.exists():
        df = pd.DataFrame(
            columns=["YY-MM", "case_descr", "amount", "invoice_number", "Статус"]
        )
        df.to_excel(CASES_XLSX, index=False)
        logger.debug("Created new Excel file %s", CASES_XLSX)
    else:
        df = pd.read_excel(CASES_XLSX)
        logger.debug("Loaded existing Excel file %s", CASES_XLSX)
    return df


def save_cases_excel(df: pd.DataFrame) -> None:
    df.to_excel(CASES_XLSX, index=False)
    logger.debug("Saved cases dataframe to %s", CASES_XLSX)


def get_gmail_service() -> object:
    logger.debug("Authorizing Gmail service")
    creds = None
    token_path = ROOT_DIR / "token.json"
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), GMAIL_SCOPES)
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            str(ROOT_DIR / "client_secret.json"), GMAIL_SCOPES
        )
        creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")
    service = build("gmail", "v1", credentials=creds, cache_discovery=False)
    logger.debug("Gmail service initialized")
    return service


def search_new_messages(service, seen_ids: set[str]) -> List[tuple[str, int]]:
    kyiv_now = datetime.now()
    after = (kyiv_now - timedelta(days=3)).strftime("%Y/%m/%d")
    query = (
        "from:docusign.net "
        "subject:(Завершен OR Завершён) "
        f"after:{after} "
        '"Approved case"'
    )
    logger.debug("Gmail search query: %s", query)
    resp = service.users().messages().list(userId="me", q=query).execute()
    results = []
    subjects: List[str] = []
    for msg in resp.get("messages", []):
        msg_id = msg["id"]
        if msg_id in seen_ids:
            continue
        full = service.users().messages().get(
            userId="me", id=msg_id, format="full"
        ).execute()
        headers = full.get("payload", {}).get("headers", [])
        subject = next((h["value"] for h in headers if h["name"] == "Subject"), "")
        subjects.append(subject)
        logger.debug("Checking message %s with subject: %s", msg_id, subject)
        body_data = full.get("payload", {}).get("body", {}).get("data")
        body = ""
        if body_data:
            body = base64.urlsafe_b64decode(body_data).decode("utf-8", errors="ignore")
        full_text = subject + "\n" + body
        m = RE_APPROVED.search(full_text)
        if not m:
            continue
        results.append((msg_id, int(m.group(1))))
    logger.info("Email subjects checked: %s", subjects)
    return results


def get_drive_service() -> tuple[object, Credentials]:
    logger.debug("Authorizing Drive service")
    token_path = ROOT_DIR / "token.json"
    creds = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), DRIVE_SCOPES)
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            str(ROOT_DIR / "client_secret.json"), DRIVE_SCOPES
        )
        creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")
    service = build("drive", "v3", credentials=creds)
    logger.debug("Drive service initialized")
    return service, creds


def get_sheets_service(creds) -> object:
    return build("sheets", "v4", credentials=creds)


def find_case_folder(service, padded: str) -> Optional[str]:
    query = (
        f"'{ROOT_FOLDER_ID}' in parents and "
        "mimeType='application/vnd.google-apps.folder' and "
        f"name='{padded}' and trashed=false"
    )
    logger.debug("Searching for case folder with query: %s", query)
    resp = service.files().list(
        q=query,
        fields="files(id)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora="allDrives",
        pageSize=1,
    ).execute()
    files = resp.get("files", [])
    folder_id = files[0]["id"] if files else None
    logger.debug("Case folder found: %s", folder_id)
    return folder_id


def find_latest_pdf(service, parent_id: str, prefix: str) -> Optional[Dict]:
    logger.debug("Searching for PDF starting with '%s' in folder %s", prefix, parent_id)
    resp = service.files().list(
        q=f"'{parent_id}' in parents and mimeType='application/pdf' and trashed=false",
        orderBy="modifiedTime desc",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        corpora="allDrives",
        pageSize=100,
    ).execute()
    for f in resp.get("files", []):
        if f["name"].lower().startswith(prefix.lower()):
            logger.debug("Found PDF: %s", f)
            return f
    logger.debug("No PDF found for prefix %s", prefix)
    return None


def clean_filename(name: str) -> str:
    return re.sub(f"[{re.escape(FORBIDDEN)}]", "_", name).strip(" .")


def download_pdf(service, file_meta: Dict, desired_name: str) -> Path:
    if not desired_name.lower().endswith(".pdf"):
        desired_name += ".pdf"
    desired_name = clean_filename(desired_name)
    logger.debug("Downloading %s to %s", file_meta.get("name"), desired_name)
    request = service.files().get_media(fileId=file_meta["id"])
    fh = io.FileIO(desired_name, "wb")
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.close()
    logger.debug("Downloaded file saved as %s", desired_name)
    return Path(desired_name)


def load_parser():
    from importlib import util as _import_util

    logger.debug("Loading parser module from %s", PARSER_FILE)
    if not PARSER_FILE.exists():
        raise FileNotFoundError(f"parser_Invoicev2.py not found at {PARSER_FILE}")
    spec = _import_util.spec_from_file_location("parser_Invoicev2", PARSER_FILE)
    module = _import_util.module_from_spec(spec)
    spec.loader.exec_module(module)  # type: ignore
    logger.debug("Parser module loaded")
    return module


def create_case_dir(date_iso: str, amount: float, invoice_number: int) -> Path:
    yy_mm = datetime.fromisoformat(date_iso).strftime("%y-%m")
    name = f"Нова {yy_mm} XXX {int(round(amount))} №{invoice_number} Хелп"
    name = clean_filename(name)
    path = ROOT_DIR / name
    path.mkdir(exist_ok=True)
    logger.debug("Case directory ready: %s", path)
    return path


def append_row(sheets, row: List) -> None:
    logger.debug("Appending row to sheet: %s", row)
    sheets.spreadsheets().values().append(
        spreadsheetId=SHEET_ID,
        range=SHEET_RANGE,
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [row]},
    ).execute()


def send_telegram(log: List[str]) -> None:
    if not log:
        return
    text = "\n".join(log)
    payload = {"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "Markdown"}
    try:
        logger.debug("Sending Telegram message")
        requests.post(TELEGRAM_URL, data=payload, timeout=30)
    except Exception as e:
        logger.debug("Telegram send failed: %s", e)


# ---------------------------------------------------------------------------
# Main processing

def process():
    logger.info("Starting processing")
    telegram_log: List[str] = []
    parser_module = load_parser()
    logger.debug("Parser module ready")

    cases_df = ensure_cases_excel()
    logger.debug("Cases dataframe loaded with %d rows", len(cases_df))
    seen_data = load_seen()
    logger.debug("Seen data: %s", seen_data)
    seen_msgs = set(seen_data.get("messages", []))

    # 1. Gmail
    try:
        logger.info("Connecting to Gmail and searching for new messages")
        gmail = get_gmail_service()
        new_msgs = search_new_messages(gmail, seen_msgs)
        logger.info("New messages found: %d", len(new_msgs))
    except Exception as e:
        new_msgs = []
        telegram_log.append(f"❌ Gmail error: {e}")
        logger.error("Gmail error: %s", e)

    for msg_id, inv_no in new_msgs:
        logger.debug("Processing message %s for invoice %s", msg_id, inv_no)
        if str(inv_no) not in cases_df["invoice_number"].astype(str).tolist():
            new_row = {
                "YY-MM": "",
                "case_descr": "",
                "amount": "",
                "invoice_number": int(inv_no),
                "Статус": "Ожидает Invoice",
            }
            cases_df = pd.concat([cases_df, pd.DataFrame([new_row])], ignore_index=True)
            logger.debug("Added new row for invoice %s", inv_no)
        telegram_log.append(f"📬 Найдено письмо: №{int(inv_no)}")
        seen_msgs.add(msg_id)
        seen_data.setdefault("cases", {}).setdefault(str(inv_no), asdict(CaseFlags()))

    # Save initial data after Gmail stage
    logger.debug("Saving state after Gmail stage")
    save_cases_excel(cases_df)
    seen_data["messages"] = list(seen_msgs)
    save_seen(seen_data)

    # 2. Process cases
    logger.info("Processing cases")
    drive, creds = get_drive_service()
    sheets = build("sheets", "v4", credentials=creds)

    for idx, row in cases_df.iterrows():
        status = str(row.get("Статус", ""))
        if status == "Готово":
            logger.debug("Skipping invoice %s - already done", row["invoice_number"])
            continue
        invoice_number = int(row["invoice_number"])
        logger.info("Processing invoice %s", invoice_number)
        case_flags = seen_data.get("cases", {}).get(str(invoice_number), asdict(CaseFlags()))
        padded = RE_CASE_FOLDER.format(invoice_number)
        try:
            folder_id = find_case_folder(drive, padded)
            if not folder_id:
                raise RuntimeError("Папка-кейс не найдена")
            invoice_meta = find_latest_pdf(drive, folder_id, "Invoice")
            grant_meta = find_latest_pdf(drive, folder_id, "Grant Agreement")
            inv_path = grant_path = None
            if invoice_meta:
                inv_path = download_pdf(
                    drive, invoice_meta, f"Invoice {invoice_number}.pdf"
                )
                telegram_log.append(f"📥 скачан файл: {inv_path.name}")
                case_flags["invoice_downloaded"] = True
            if grant_meta:
                grant_path = download_pdf(
                    drive, grant_meta, f"Grant Agreement {invoice_number}.pdf"
                )
                telegram_log.append(f"📥 скачан файл: {grant_path.name}")
                case_flags["grant_downloaded"] = True

            if inv_path:
                info: Dict = parser_module.parse_invoice(str(inv_path))
                logger.debug("Parsed invoice info: %s", info)
                yy_mm = datetime.fromisoformat(info["date"]).strftime("%y-%m")
                cases_df.loc[idx, ["YY-MM", "case_descr", "amount"]] = [
                    yy_mm,
                    info.get("case_descr", ""),
                    info.get("amount", ""),
                ]
                telegram_log.append("📊 Invoice распарсен")
                case_flags["parsed"] = True
                target_dir = create_case_dir(info["date"], info["amount"], invoice_number)
                shutil.move(str(inv_path), target_dir / inv_path.name)
                if grant_path:
                    shutil.move(str(grant_path), target_dir / grant_path.name)
                telegram_log.append("📂 Папка сформирована")
                append_row(
                    sheets,
                    [
                        info["date"],
                        "",
                        "",
                        invoice_number,
                        info.get("case_descr", ""),
                        info.get("amount", ""),
                        "",
                        "",
                        "",
                        "",
                        "",
                        "хер",
                    ],
                )
                telegram_log.append("📊 Добавлено в таблицу")
                cases_df.loc[idx, "Статус"] = "Готово"
                telegram_log.append(f"✅ Кейс №{invoice_number} обработан")
            else:
                cases_df.loc[idx, "Статус"] = "Ошибка: Invoice не найден"
                telegram_log.append(
                    f"❌ Ошибка кейса №{invoice_number}: Invoice не найден"
                )
                case_flags["error"] = "Invoice not found"
        except Exception as e:
            cases_df.loc[idx, "Статус"] = f"Ошибка: {e}"
            telegram_log.append(f"❌ Ошибка кейса №{invoice_number}: {e}")
            case_flags["error"] = str(e)
            logger.error("Error processing invoice %s: %s", invoice_number, e)
        finally:
            seen_data.setdefault("cases", {})[str(invoice_number)] = case_flags
            save_cases_excel(cases_df)
            save_seen(seen_data)

    send_telegram(telegram_log)
    logger.info("Processing finished")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Mega_Help Script")
    parser.add_argument("--debug", action="store_true", help="Enable debug output")
    args = parser.parse_args()
    logging.basicConfig(
        level=logging.DEBUG if args.debug else logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    try:
        process()
    except HttpError as e:
        logger.error("Google API error: %s", e)
