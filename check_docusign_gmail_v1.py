#!/usr/bin/env python3
"""
check_docusign_gmail.py
Запускать из Task Scheduler каждые 2 часа.

Проверяет Gmail (API), ищет письма от @docusign.net,
чья тема начинается словом «Завершен» / «Завершён» и
содержит фразу 'for Approved case <8-digit>'.
Извлекает число и печатает, не дублируя уже обработанные письма.
"""

import json
import pathlib
import re
import sys
from datetime import datetime, timedelta
from typing import List

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from zoneinfo import ZoneInfo   # Python 3.9+

# ---------- настраиваемые файлы ----------
CREDENTIALS_FILE = 'client_secret.json'          # OAuth client_secret
TOKEN_FILE       = 'token.json'                # refresh-token
SEEN_FILE        = 'seen_ids.json'             # обработанные письма
# -----------------------------------------

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']  # или gmail.modify

# ─── Gmail search query ─────────────────────────────────────────────────────────
# «after:YYYY/MM/DD» берёт письма, полученные ПОСЛЕ указанной даты 00:00,
# поэтому yesterday = today - 1 day → получаем сегодня + вчера.
kyiv_now      = datetime.now(ZoneInfo('Europe/Kyiv'))
yesterday_str = (kyiv_now - timedelta(days=1)).strftime('%Y/%m/%d')
GMAIL_QUERY   = (
    f'from:docusign.net '
    f'subject:Завершен OR subject:Завершён '
    f'after:{yesterday_str}'
)
# ────────────────────────────────────────────────────────────────────────────────

SUBJECT_RE = re.compile(r'for Approved case\s+(\d{8})', re.IGNORECASE)
TZ         = ZoneInfo('Europe/Kyiv')           # для читаемых тайм-штампов


def load_seen() -> List[str]:
    if pathlib.Path(SEEN_FILE).exists():
        with open(SEEN_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []


def save_seen(seen: List[str]) -> None:
    with open(SEEN_FILE, 'w', encoding='utf-8') as f:
        json.dump(seen, f)


def get_service():
    """Авторизация Gmail API c единым списком SCOPES."""
    creds = None
    if pathlib.Path(TOKEN_FILE).exists():
        # НЕ передаём scopes -> будут использованы те, что зашиты в token.json
        creds = Credentials.from_authorized_user_file(TOKEN_FILE)
    # Если файла нет ИЛИ токен недействителен/истёк
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f'⚠️  Не удалось обновить токен: {e}\nПолучаем новый…')
                creds = None  # упадём в блок «получить новый»
        if not creds or not creds.valid:
            # Полный поток авторизации
            from google_auth_oauthlib.flow import InstalledAppFlow
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
            with open(TOKEN_FILE, 'w', encoding='utf-8') as token:
                token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds, cache_discovery=False)


def extract_case_number(subject: str) -> str | None:
    """Возвращает 8-значный номер из темы или None."""
    m = SUBJECT_RE.search(subject)
    return m.group(1) if m else None


def main():
    seen_ids = set(load_seen())
    service  = get_service()

    results  = service.users().messages().list(userId='me', q=GMAIL_QUERY).execute()
    for msg in results.get('messages', []):
        msg_id = msg['id']
        if msg_id in seen_ids:         # уже обработан
            continue

        # Получаем только заголовки (быстро)
        meta = service.users().messages().get(
            userId='me', id=msg_id,
            format='metadata',
            metadataHeaders=['Subject']
        ).execute()

        subject = next(
            (h['value'] for h in meta['payload']['headers'] if h['name'] == 'Subject'),
            ''
        )

        number = extract_case_number(subject)
        if number:
            ts = datetime.now(TZ).strftime('%Y-%m-%d %H:%M:%S')
            print(f'[{ts}] Approved case → {number}')
        else:
            # не кричим исключением, просто лог в stderr
            print(f'⚠️  Не смог извлечь номер из темы: «{subject}»', file=sys.stderr)

        seen_ids.add(msg_id)

    save_seen(list(seen_ids))


if __name__ == '__main__':
    main()
