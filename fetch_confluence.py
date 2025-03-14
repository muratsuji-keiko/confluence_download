import requests
import io
import os
import re
import pickle
import json
import base64
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials  # ✅ 追加（`Refresh Token` 用）
from google_auth_oauthlib.flow import InstalledAppFlow
from reportlab.platypus import Paragraph, SimpleDocTemplate
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT
from reportlab.lib import colors
from bs4 import BeautifulSoup

# ✅ Confluence 認証（環境変数から取得）
CONFLUENCE_PAT = os.getenv("CONFLUENCE_PAT")
CONFLUENCE_URL = "https://confl.arms.dmm.com"
headers = {
    "Authorization": f"Bearer {CONFLUENCE_PAT}",
    "Accept": "application/json"
}

# ✅ Google Drive API 設定
SCOPES = ["https://www.googleapis.com/auth/drive.file"]
PARENT_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID")

PARENT_PAGES = [
    {"id": "2191300228", "name": "01.事業監査会_2025年度"},
    {"id": "694536054", "name": "02.事業別損益（BI）"},
    {"id": "469803397", "name": "03.部門コード"},
    {"id": "694536236", "name": "90.各種手続き"},
    {"id": "469843465", "name": "99.FAQ"}
]

EXCLUDED_TITLES = ["廃止", "2020予算", "2020年度", "2021予算", "2021年度",
                    "2022予算", "2022年度", "2023予算", "2023年度", "2024予算", "2024年度"]

# ✅ 環境変数から `credentials.json` を復元
def restore_google_credentials():
    credentials_b64 = os.getenv("GOOGLE_CREDENTIALS")
    if credentials_b64:
        with open("credentials.json", "wb") as f:
            f.write(base64.b64decode(credentials_b64))
        print("✅ credentials.json を復元しました")
    else:
        print("⚠️ 環境変数 GOOGLE_CREDENTIALS が設定されていません")
        exit(1)

restore_google_credentials()

# ✅ Google Drive 認証（`GOOGLE_REFRESH_TOKEN` を使用）
def authenticate_google_drive():
    creds = None

    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    # ✅ `Refresh Token` を使って `Access Token` を取得
    if not creds or not creds.valid:
        refresh_token = os.getenv("GOOGLE_REFRESH_TOKEN")
        client_id = os.getenv("GOOGLE_CLIENT_ID")  # `credentials.json` から取得
        client_secret = os.getenv("GOOGLE_CLIENT_SECRET")  # `credentials.json` から取得

        if not refresh_token or not client_id or not client_secret:
            print("⚠️ 環境変数 GOOGLE_REFRESH_TOKEN, GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET が必要です。")
            exit(1)

        creds = Credentials(
            None,
            refresh_token=refresh_token,
            token_uri="https://oauth2.googleapis.com/token",
            client_id=client_id,
            client_secret=client_secret
        )
        creds.refresh(Request())  # ✅ 自動的に `Access Token` を更新

        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    return build("drive", "v3", credentials=creds)

drive_service = authenticate_google_drive()

def fetch_page_content(page_id):
    """Confluence からページ内容を取得"""
    url = f"{CONFLUENCE_URL}/rest/api/content/{page_id}?expand=body.export_view,title"
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        print(f"❌ Error fetching page content: {response.status_code} - {response.text}")
        return None, None

    data = response.json()
    title = data.get("title", "Untitled")
    raw_content = data.get("body", {}).get("export_view", {}).get("value", "")

    if not raw_content:
        return title, "(本文なし)"

    soup = BeautifulSoup(raw_content, "html.parser")
    text_content = soup.get_text("\n").strip()

    # ✅ ゼロ幅スペースや制御文字を削除
    text_content = re.sub(r'[\u200B-\u200D\uFEFF]', '', text_content)

    return title, f"# {title}\n\n{text_content}"

def fetch_and_upload_recursive(page_id, parent_folder_id):
    """再帰的に Confluence ページを取得し、Google Drive に保存"""
    title, content = fetch_page_content(page_id)

    if title and (title.lower().startswith("wip") or any(keyword in title for keyword in EXCLUDED_TITLES)):
        print(f"⏭ Skipping excluded page: {title}")
        return

    if content:
        upload_to_google_drive(title, content, parent_folder_id, page_id)

    for child in fetch_child_pages(page_id):
        fetch_and_upload_recursive(child["id"], parent_folder_id)

def main():
    for parent_page in PARENT_PAGES:
        fetch_and_upload_recursive(parent_page["id"], PARENT_FOLDER_ID)

if __name__ == "__main__":
    main()
