import requests
import io
import os
import re
import pickle
import base64
import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from reportlab.platypus import Paragraph, SimpleDocTemplate
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT
from reportlab.lib import colors
from bs4 import BeautifulSoup

# ✅ Confluence認証（環境変数から取得）
CONFLUENCE_PAT = os.getenv("CONFLUENCE_PAT")
CONFLUENCE_URL = "https://confl.arms.dmm.com"
headers = {
    "Authorization": f"Bearer {CONFLUENCE_PAT}",
    "Accept": "application/json"
}

# ✅ Google Drive API設定
SCOPES = ["https://www.googleapis.com/auth/drive.file"]
PARENT_FOLDER_ID = "1043WWogLOtmo63cUYKkYoBk5ub0sgShM"

PARENT_PAGES = [
    {"id": "2191300228", "name": "01.事業監査会_2025年度"},
    {"id": "694536054", "name": "02.事業別損益（BI）"},
    {"id": "469803397", "name": "03.部門コード"},
    {"id": "694536236", "name": "90.各種手続き"},
    {"id": "469843465", "name": "99.FAQ"}
]

# 🔽 除外対象のタイトルリスト
EXCLUDED_TITLES = ["廃止", "2020予算", "2020年度", "2021予算", "2021年度",
                    "2022予算", "2022年度", "2023予算", "2023年度", "2024予算", "2024年度"]

SCOPES = ["https://www.googleapis.com/auth/drive.file"]

def authenticate_google_drive():
    creds = None

    # Render では環境変数 `GOOGLE_CREDENTIALS` を使用
    google_credentials_base64 = os.getenv("GOOGLE_CREDENTIALS")
    if google_credentials_base64:
        google_credentials_json = base64.b64decode(google_credentials_base64).decode("utf-8")
        with open("credentials_temp.json", "w") as temp_file:
            temp_file.write(google_credentials_json)
        creds = InstalledAppFlow.from_client_secrets_file("credentials_temp.json", SCOPES).run_local_server(port=0)
        os.remove("credentials_temp.json")  # 一時ファイルを削除

    # `token.pickle` のキャッシュを利用
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            creds = InstalledAppFlow.from_client_secrets_file("credentials_temp.json", SCOPES).run_local_server(port=0)

        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    return build("drive", "v3", credentials=creds)

drive_service = authenticate_google_drive()


def fetch_page_content(page_id):
    url = f"{CONFLUENCE_URL}/rest/api/content/{page_id}?expand=body.export_view,title"
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        print(f"❌ Error fetching page content: {response.status_code} - {response.text}")
        return None, None

    data = response.json()
    title = data.get("title", "Untitled")

    # `export_view` の生HTMLを取得
    raw_content = data.get("body", {}).get("export_view", {}).get("value", "")

    if not raw_content:
        return title, "(本文なし)"

    # 🎯 余計なタグや特殊文字を削除
    soup = BeautifulSoup(raw_content, "html.parser")
    text_content = soup.get_text("\n").strip()

    # 🎯 Unicode制御文字（ゼロ幅スペースなど）を削除
    text_content = re.sub(r'[\u200B-\u200D\uFEFF]', '', text_content)

    return title, f"# {title}\n\n{text_content}"

def upload_to_google_drive(page_id, file_name, content, parent_folder_id):
    """コンフルエンスのページ内容をメモリ上でPDF化し、Google Driveに直接アップロード"""

    # **日本語フォントを登録**
    font_path = "C:/Windows/Fonts/msgothic.ttc"
    pdfmetrics.registerFont(TTFont("MSGothic", font_path))

    # **メモリ上でPDFを生成**
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4)

    # **スタイル設定**
    styles = getSampleStyleSheet()
    styles["Normal"].fontName = "MSGothic"
    styles["Normal"].fontSize = 10
    styles["Normal"].leading = 12
    styles["Normal"].textColor = colors.black

    # **タイトルを設定**
    title_paragraph = Paragraph(f"<b>{file_name}</b>", styles["Normal"])

    # **本文を改行処理してセット**
    content = content.replace("\n", "<br/>")
    body_paragraph = Paragraph(content, styles["Normal"])

    # **PDF に書き込み**
    elements = [title_paragraph, body_paragraph]
    doc.build(elements)

    # **PDFデータをバッファに保存**
    pdf_buffer.seek(0)

    # **Google Drive 内で「ページIDを含むファイル名」を検索**
    formatted_file_name = f"{page_id}_{file_name}.pdf"
    query = f"name = '{formatted_file_name}' and '{parent_folder_id}' in parents and trashed=false"
    results = drive_service.files().list(
        q=query, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True
    ).execute()
    existing_files = results.get("files", [])

    media = MediaIoBaseUpload(pdf_buffer, mimetype="application/pdf", resumable=True)

    if existing_files:
        file_id = existing_files[0]["id"]
        drive_service.files().update(
            fileId=file_id,
            media_body=media,
            supportsAllDrives=True
        ).execute()
        print(f"♻️ Updated: {formatted_file_name} (File ID: {file_id})")
    else:
        file_metadata = {
            "name": formatted_file_name,
            "parents": [parent_folder_id]
        }
        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id",
            supportsAllDrives=True
        ).execute()
        print(f"✅ Uploaded: {formatted_file_name} (File ID: {file.get('id')})")

def fetch_and_upload_recursive(page_id, parent_folder_id):
    title, content = fetch_page_content(page_id)

    if title and (title.lower().startswith("wip") or any(keyword in title for keyword in EXCLUDED_TITLES)):
        print(f"⏭ Skipping excluded page: {title}")
        return

    if content:
        formatted_file_name = f"{page_id}_{title}"
        upload_to_google_drive(page_id, formatted_file_name, content, parent_folder_id)

    child_pages = fetch_child_pages(page_id)
    for child in child_pages:
        fetch_and_upload_recursive(child["id"], parent_folder_id)

def main():
    for parent_page in PARENT_PAGES:
        parent_folder_id = get_or_create_drive_folder(parent_page["name"], PARENT_FOLDER_ID)
        fetch_and_upload_recursive(parent_page["id"], parent_folder_id)

if __name__ == "__main__":
    drive_service = authenticate_google_drive()
    main()
