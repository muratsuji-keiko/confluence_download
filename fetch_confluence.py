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
from google_auth_oauthlib.flow import InstalledAppFlow
from reportlab.platypus import Paragraph, SimpleDocTemplate
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
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

# ✅ Google 認証
def authenticate_google_drive():
    creds = None

    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())  # トークンをリフレッシュ
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)

            # ✅ Render環境では `run_local_server()` ではなく `authorization_url` を使う
            auth_url, _ = flow.authorization_url(prompt="consent")
            print(f"🔗 認証用のURL: {auth_url}")

            # ✅ ユーザーが手動でコードを入力する
            auth_code = input("🔑 上記のリンクを開き、認証コードを入力してください: ").strip()
            flow.fetch_token(code=auth_code)
            creds = flow.credentials

        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    return build("drive", "v3", credentials=creds)

drive_service = authenticate_google_drive()

def get_or_create_drive_folder(folder_name, parent_folder_id):
    """Google Drive上で親ページフォルダを取得 or 作成（子フォルダは作らない）"""
    query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}' and '{parent_folder_id}' in parents and trashed=false"
    results = drive_service.files().list(
        q=query, fields="files(id)", supportsAllDrives=True, includeItemsFromAllDrives=True
    ).execute().get("files", [])

    if results:
        return results[0]["id"]

    folder_metadata = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_folder_id]
    }
    folder = drive_service.files().create(
        body=folder_metadata, fields="id", supportsAllDrives=True
    ).execute()
    return folder.get("id")

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

def fetch_child_pages(page_id):
    """指定した Confluence ページの子ページ一覧を取得"""
    url = f"{CONFLUENCE_URL}/rest/api/content/{page_id}/child/page?expand=title"
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        print(f"❌ Error fetching child pages: {response.status_code} - {response.text}")
        return []
    return response.json().get("results", [])

def upload_to_google_drive(file_name, content, parent_folder_id, page_id):
    """ページIDを使ってGoogle Drive内のファイルを上書き判定し、PDFを保存"""

    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4)

    styles = getSampleStyleSheet()
    styles["Normal"].fontSize = 10
    styles["Normal"].leading = 12
    styles["Normal"].alignment = TA_LEFT
    styles["Normal"].textColor = colors.black

    title_paragraph = Paragraph(f"<b>{file_name}</b>", styles["Normal"])
    content = content.replace("\n", "<br/>")
    body_paragraph = Paragraph(content, styles["Normal"])

    elements = [title_paragraph, body_paragraph]
    doc.build(elements)

    pdf_buffer.seek(0)
    media = MediaIoBaseUpload(pdf_buffer, mimetype="application/pdf", resumable=True)

    formatted_file_name = f"{file_name}.pdf"
    query = f"name = '{formatted_file_name}' and '{parent_folder_id}' in parents and trashed=false"
    results = drive_service.files().list(
        q=query, fields="files(id, name, appProperties)", supportsAllDrives=True, includeItemsFromAllDrives=True
    ).execute()
    existing_files = results.get("files", [])

    matching_file = next((file for file in existing_files if file.get("appProperties", {}).get("page_id") == str(page_id)), None)

    if matching_file:
        drive_service.files().update(fileId=matching_file["id"], media_body=media, supportsAllDrives=True).execute()
        print(f"♻️ Updated: {formatted_file_name}")
    else:
        file_metadata = {"name": formatted_file_name, "parents": [parent_folder_id], "appProperties": {"page_id": str(page_id)}}
        drive_service.files().create(body=file_metadata, media_body=media, fields="id", supportsAllDrives=True).execute()
        print(f"✅ Uploaded: {formatted_file_name}")

def main():
    for parent_page in PARENT_PAGES:
        parent_folder_id = get_or_create_drive_folder(parent_page["name"], PARENT_FOLDER_ID)
        fetch_and_upload_recursive(parent_page["id"], parent_folder_id)

if __name__ == "__main__":
    main()
