import os
import pickle
import logging
import datetime
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials
import base64
from email.mime.text import MIMEText

from lineworks_api import LineWorksAPI

# ===== 設定項目（実際の値に書き換えてください） =====
LW_CONFIG = {
    "CLIENT_ID": "YOUR_CLIENT_ID",
    "CLIENT_SECRET": "YOUR_CLIENT_SECRET",
    "SERVICE_ACCOUNT": "YOUR_SERVICE_ACCOUNT_ID",
    "PRIVATE_KEY": """-----BEGIN PRIVATE KEY-----
YOUR_PRIVATE_KEY_HERE
-----END PRIVATE KEY-----""",
    "USER_ID": "s.abe@runbird-biz"
}

GS_CONFIG = {
    "SPREADSHEET_ID": "1Qrd2HTrWSeYp2f5vrMg3rV7ey72GS2nN463LdzgrrQo",
    "SHEET_GID": "1447447650",
    "EMAIL_COLUMN": 17,  # Q列
    "DATA_START_ROW": 2,
    "CREDENTIALS_FILE": "credentials.json"  # Google Service Account JSON
}

GMAIL_CONFIG = {
    "FROM_EMAIL": "creative.yt@runbird.net",
    "SUBJECT": "【定例会日程調整】月次報告のご案内（株式会社ランバード）",
    "BODY_TMPL": """お客様各位

平素より大変お世話になっております。
株式会社ランバード　広告運用サポート窓口です。

本メールは皆様へのお知らせのためBCCで送信しております。

月次の報告をさせていただきたく思い、日程調整のご連絡をいたしました。
下記日程より、3つほどご都合の良いお日にちを挙げて頂ければ幸いでございます。

【候補日時】
{CANDIDATE_DATES}
※所要時間は30分ほどを予定しております。

ご確認のほどよろしくお願いいたします。
＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
株式会社ランバード
広告運用サポート窓口
東京都中央区晴海1丁目8－10
晴海アイランド トリトンスクエア オフィスタワーX棟40階
TEL　0120-949-851　FAX　03-5843-6978
Mail   creative.yt@runbird.net
HP　 https://runbird.net/
＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝"""
}

# Gmail API用スコープ
GMAIL_SCOPES = ['https://www.googleapis.com/auth/gmail.compose']
# Google Sheets用スコープ
GS_SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

def get_gmail_service():
    """Gmail APIサービスを取得する（初回のOAuth認証を伴う）"""
    creds = None
    if os.path.exists('token_gmail.pickle'):
        with open('token_gmail.pickle', 'rb') as token:
            creds = pickle.load(token)
            
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # 連携するGoogleアカウントのクライアントシークレットが必要
            if not os.path.exists('client_secret.json'):
                print("Error: 'client_secret.json' (OAuth Client Secret) is required for Gmail API.")
                return None
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
            
        with open('token_gmail.pickle', 'wb') as token:
            pickle.dump(creds, token)
            
    return build('gmail', 'v1', credentials=creds)

def get_client_emails():
    """スプレッドシートからメールアドレスを取得する"""
    try:
        creds = Credentials.from_service_account_file(GS_CONFIG["CREDENTIALS_FILE"], scopes=GS_SCOPES)
        client = gspread.authorize(creds)
        ss = client.open_by_key(GS_CONFIG["SPREADSHEET_ID"])
        
        # GIDでシートを特定
        sheets = ss.worksheets()
        target_ws = None
        for s in sheets:
            if str(s.id) == GS_CONFIG["SHEET_GID"]:
                target_ws = s
                break
        
        if not target_ws:
            logger.warning(f"Sheet with GID {GS_CONFIG['SHEET_GID']} not found. Using the first sheet.")
            target_ws = sheets[0]
            
        all_values = target_ws.get_all_values()
        emails = []
        for row in all_values[GS_CONFIG["DATA_START_ROW"]-1:]:
            if len(row) >= GS_CONFIG["EMAIL_COLUMN"]:
                email = row[GS_CONFIG["EMAIL_COLUMN"]-1].strip()
                if email and "@" in email:
                    emails.append(email)
        
        return emails
    except Exception as e:
        logger.error(f"Error fetching emails from Sheet: {e}")
        return []

def create_draft(service, bcc_list, body):
    """Gmail APIで下書きを作成する"""
    message = MIMEText(body)
    message['to'] = GMAIL_CONFIG["FROM_EMAIL"]  # 自分宛（BCCメインのため）
    message['subject'] = GMAIL_CONFIG["SUBJECT"]
    message['Bcc'] = ", ".join(bcc_list)
    
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    draft = {
        'message': {
            'raw': raw
        }
    }
    
    try:
        res = service.users().drafts().create(userId='me', body=draft).execute()
        logger.info(f"Draft created successfully. Draft ID: {res['id']}")
        return res
    except Exception as e:
        logger.error(f"Error creating draft: {e}")
        return None

def main():
    logger.info("Starting meeting scheduler process...")
    
    # 1. LINE WORKSから候補予定を取得
    try:
        lw = LineWorksAPI(
            LW_CONFIG["CLIENT_ID"],
            LW_CONFIG["CLIENT_SECRET"],
            LW_CONFIG["SERVICE_ACCOUNT"],
            LW_CONFIG["PRIVATE_KEY"]
        )
        logger.info("Fetching LINE WORKS events...")
        events = lw.get_candidate_events(LW_CONFIG["USER_ID"])
        if not events:
            logger.warning("No candidate events found in LINE WORKS.")
            return

        date_list_text = lw.format_events_as_list(events)
        logger.info(f"Formatted candidate dates:\n{date_list_text}")
    except Exception as e:
        logger.error(f"LINE WORKS process failed: {e}")
        return

    # 2. スプレッドシートからBCCリストを取得
    logger.info("Fetching client emails from Spreadsheet...")
    client_emails = get_client_emails()
    if not client_emails:
        logger.warning("No client emails found in Spreadsheet.")
        return
    logger.info(f"Found {len(client_emails)} client emails.")

    # 3. Gmail下書き作成
    logger.info("Connecting to Gmail API...")
    gmail_service = get_gmail_service()
    if not gmail_service:
        logger.error("Could not initialize Gmail service.")
        return

    email_body = GMAIL_CONFIG["BODY_TMPL"].format(CANDIDATE_DATES=date_list_text)
    create_draft(gmail_service, client_emails, email_body)
    
    logger.info("Process completed successfully.")

if __name__ == "__main__":
    main()
