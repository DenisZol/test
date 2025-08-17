#!/usr/bin/env python3
"""
DriveScript v1.1 (Verbose + Shared Drive aware)
-----------------------------------------------
Скачивает самый свежий «Grant Agreement*.pdf» для указанного кейса.

🔑  Требования:
    • client_secret.json   – рядом со скриптом
    • pip install google-api-python-client google-auth google-auth-oauthlib

🔧  Запуск:
    python DriveScript.py 13297            # «тихий» режим
    python DriveScript.py 13297 --verbose  # подробный лог
"""

from __future__ import annotations
import argparse
import io
import os
import sys
from datetime import datetime, timezone
from typing import Dict, List, Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError

# ────────────────────────── КОНСТАНТЫ ──────────────────────────
ROOT_FOLDER_ID = "1KNvnzuBL_froKQs-JVd8TVoGDMtL4-wx"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
PAGE_SIZE = 1000      # сколько элементов за один запрос list()
# ────────────────────────────────────────────────────────────────


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Download latest Grant-Agreement PDF for a case.",
    )
    parser.add_argument("case_id", help="Номер кейса / название подпапки")
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Показывать подробный пошаговый лог"
    )
    return parser.parse_args()


def log(msg: str, *, verbose: bool):
    """Выводит сообщение, если включён verbose-режим."""
    if verbose:
        print(msg)


# ─────────────────────────── OAuth 2.0 ─────────────────────────
def get_credentials(verbose: bool) -> Credentials:
    creds: Optional[Credentials] = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        log("🔑 token.json найден, пробую использовать его…", verbose=verbose)

    # обновляем при необходимости
    if creds and creds.expired and creds.refresh_token:
        log("🔄 Токен просрочен, пробую refresh…", verbose=verbose)
        creds.refresh(Request())

    # если токена нет или обновить не вышло → полный OAuth-flow
    if not creds or not creds.valid:
        log("🌐 Запускаю браузер для авторизации…", verbose=verbose)
        if not os.path.exists("client_secret.json"):
            sys.exit("❌ client_secret.json не найден рядом со скриптом.")
        flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", SCOPES)
        creds = flow.run_local_server(port=0)
        with open("token.json", "w", encoding="utf-8") as f:
            f.write(creds.to_json())
        log("✅ token.json сохранён.", verbose=verbose)

    return creds


# ─────────────────────── Работа с Drive API ─────────────────────
def build_service(creds: Credentials, verbose: bool):
    log("⚙️  Строю service object Drive v3…", verbose=verbose)
    return build("drive", "v3", credentials=creds)


def list_subfolders(service, parent_id: str, verbose: bool) -> List[Dict]:
    """Возвращает список всех подпапок в parent_id (1-й уровень)."""
    log(f"📑 Читаю содержимое папки {parent_id}…", verbose=verbose)
    items: List[Dict] = []
    page_token = None
    while True:
        resp = (
            service.files()
            .list(
                q=f"'{parent_id}' in parents "
                  "and mimeType='application/vnd.google-apps.folder' "
                  "and trashed=false",
                fields="nextPageToken, files(id, name)",
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
                corpora="allDrives",
                pageSize=PAGE_SIZE,
                pageToken=page_token,
            )
            .execute()
        )
        items.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return items


def find_case_folder(service, case_id: str, verbose: bool) -> Optional[str]:
    """Ищет подпапку case_id внутри ROOT_FOLDER_ID."""
    # 1) Прямая попытка (быстро)
    log(f"🔎 Ищу подпапку «{case_id}» напрямую…", verbose=verbose)
    query = (
        f"'{ROOT_FOLDER_ID}' in parents and "
        "mimeType='application/vnd.google-apps.folder' and "
        f"name='{case_id}' and trashed=false"
    )
    resp = (
        service.files()
        .list(
            q=query,
            fields="files(id, name)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            corpora="allDrives",
            pageSize=1,
        )
        .execute()
    )
    if resp.get("files"):
        return resp["files"][0]["id"]

    # 2) Если не нашли — покажем,
    #    что вообще лежит в корне (поможет глазами убедиться)
    log("⚠️  Точная подпапка не найдена. "
        "Список подпапок в корневой:", verbose=verbose)
    children = list_subfolders(service, ROOT_FOLDER_ID, verbose)
    for f in children[:20]:
        log(f"   • {f['name']}  (ID: {f['id']})", verbose=verbose)
    if len(children) > 20:
        log(f"   …ещё {len(children) - 20} папок скрыто…", verbose=verbose)
    return None


# ─────────────────── поиск PDF Grant Agreement ──────────────
def find_latest_grant(service, parent_id: str, verbose: bool) -> Optional[Dict]:
    """
    Ищет ВСЕ PDF напрямую внутри parent_id, выбирает тот,
    у которого name.lower().startswith("grant agreement").
    Берёт самый свежий modifiedTime.
    """
    log("🔎 Ищу файлы «Grant Agreement*.pdf» в самой папке-кейсе…", verbose=verbose)

    query = (
        f"'{parent_id}' in parents and "
        "mimeType='application/pdf' and trashed=false"
    )

    resp = (
        service.files()
        .list(
            q=query,
            fields="files(id, name, modifiedTime)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            corpora="allDrives",
            pageSize=PAGE_SIZE,
        )
        .execute()
    )

    # фильтруем в питоне: имя ДОЛЖНО начинаться с «Grant Agreement» (регистр игнорируем)
    candidates = [
        f for f in resp.get("files", [])
        if f["name"].lower().startswith("grant agreement")
    ]

    if not candidates:
        log("⚠️  Подходящих PDF не найдено. Ниже список всех PDF в папке:", verbose=verbose)
        if verbose:
            for f in resp.get("files", [])[:30]:
                print("   •", f["name"])
        return None

    latest = max(candidates, key=lambda f: f["modifiedTime"])
    log(f"📄 Найден: {latest['name']} (modified {latest['modifiedTime']})", verbose=verbose)
    return latest


def download(service, file_id: str, gdrive_name: str, verbose: bool):
    """
    Скачивает файл file_id из Google Drive и сохраняет его в текущей
    директории под тем же именем, гарантируя расширение «.pdf».
    """
    # ── 1. Формируем локальное имя ──────────────────────────────
    local_name = gdrive_name
    if not local_name.lower().endswith(".pdf"):
        local_name += ".pdf"          # добавляем, если Drive-имя было без расширения

    local_path = os.path.join(os.getcwd(), local_name)
    log(f"⬇️  Скачиваю «{local_name}»…", verbose=verbose)

    # ── 2. Качаем через MediaIoBaseDownload ─────────────────────
    request = service.files().get_media(fileId=file_id)
    with io.FileIO(local_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status and verbose:
                print(f"   {int(status.progress() * 100)} %")

    print(f"✅ Файл сохранён: {local_name}")



# ───────────────────────────── main() ───────────────────────────
def main():
    args = parse_args()
    verbose = args.verbose

    print(f"🚀 Запуск DriveScript для кейса {args.case_id}")
    try:
        creds = get_credentials(verbose)
        service = build_service(creds, verbose)

        case_folder_id = find_case_folder(service, args.case_id, verbose)
        if not case_folder_id:
            sys.exit("❌ Папка кейса не найдена. "
                     "Проверьте, правильно ли указан номер кейса "
                     "и существует ли она внутри CO ICF HELP GLOBAL.")

        log(f"📂 case_folder_id = {case_folder_id}", verbose=verbose)

        file_meta = find_latest_grant(service, case_folder_id, verbose)
        if not file_meta:
            sys.exit("❌ В папке нет файлов «Grant Agreement*.pdf».")

        download(
            service,
            file_meta["id"],
            file_meta["name"],
            verbose=verbose,
        )

    except HttpError as e:
        sys.exit(f"Google API Error: {e}")
    except KeyboardInterrupt:
        sys.exit("\n⏹️  Прервано пользователем.")


if __name__ == "__main__":
    main()
