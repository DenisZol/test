"""DriveScript
================
A command-line utility that downloads the latest "Grant Agreement" PDF for a
specified case from a shared Google Drive folder.

Usage:
    python DriveScript.py <case_number>

The script expects `client_secret.json` to be located in the same directory.
Upon first run it will create `token.json` for OAuth credentials.
"""

import io
import os
import sys
from typing import List, Dict

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Scope providing read-only access to the user's Drive.
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
# ID of the shared folder "CO ICF HELP GLOBAL".
ROOT_FOLDER_ID = "1KNvnzuBL_froKQs-JVd8TVoGDMtL4-wx"


def authenticate() -> Credentials:
    """Authenticate the user and return valid credentials.

    If a saved token exists it will be used; otherwise an OAuth flow will be
    initiated and `token.json` will be created.
    """

    creds: Credentials | None = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "client_secret.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open("token.json", "w", encoding="utf-8") as token:
            token.write(creds.to_json())
    return creds


def find_case_folder(service, case_number: str) -> str | None:
    """Return the folder ID matching the given case number, if it exists."""
    query = (
        f"'{ROOT_FOLDER_ID}' in parents and "
        "mimeType='application/vnd.google-apps.folder' and "
        f"name='{case_number}' and trashed=false"
    )
    results = service.files().list(q=query, fields="files(id, name)").execute()
    folders: List[Dict[str, str]] = results.get("files", [])
    return folders[0]["id"] if folders else None


def find_latest_grant_file(service, folder_id: str) -> Dict[str, str] | None:
    """Return metadata of the newest 'Grant Agreement*.pdf' file in the folder."""
    query = (
        f"'{folder_id}' in parents and mimeType!='application/vnd.google-apps.folder' "
        "and name contains 'Grant Agreement' and name contains '.pdf' and trashed=false"
    )
    results = service.files().list(
        q=query, fields="files(id, name, createdTime)").execute()
    files: List[Dict[str, str]] = results.get("files", [])
    # Filter to match prefix/suffix precisely and select the most recent by createdTime.
    grant_files = [
        f for f in files if f["name"].startswith("Grant Agreement") and f["name"].endswith(".pdf")
    ]
    if not grant_files:
        return None
    return max(grant_files, key=lambda f: f["createdTime"])


def download_file(service, file_id: str, name: str) -> None:
    """Download the specified file to the current working directory."""
    request = service.files().get_media(fileId=file_id)
    filepath = os.path.join(os.getcwd(), name)
    with open(filepath, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                progress = int(status.progress() * 100)
                print(f"Download {progress}%.")
    print(f"Downloaded '{name}' to '{filepath}'.")


def main() -> None:
    if len(sys.argv) != 2:
        print("Usage: python DriveScript.py <case_number>")
        sys.exit(1)

    case_number = sys.argv[1]

    try:
        creds = authenticate()
        service = build("drive", "v3", credentials=creds)

        folder_id = find_case_folder(service, case_number)
        if not folder_id:
            print(f"Case folder '{case_number}' not found.")
            return

        latest_file = find_latest_grant_file(service, folder_id)
        if not latest_file:
            print("No matching 'Grant Agreement' PDF files found.")
            return

        download_file(service, latest_file["id"], latest_file["name"])

    except FileNotFoundError as err:
        print(f"Credential file missing: {err}")
    except HttpError as error:
        print(f"An error occurred: {error}")


if __name__ == "__main__":
    main()
