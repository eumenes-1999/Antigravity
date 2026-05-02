import os
import sys
import base64
from email.message import EmailMessage
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import re

# スプレッドシート読み取り権限と、Gmailの下書き作成権限
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/gmail.compose'
]

SPREADSHEET_ID = '1Qrd2HTrWSeYp2f5vrMg3rV7ey72GS2nN463LdzgrrQo'
RANGE_NAME = 'R2:S'  # R列（メール）とS列（送信対象チェック）を取得

def main(candidate_dates):
    creds = None
    # 以前に保存されたトークンがあれば読み込む
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # 有効な認証情報がない場合はログインさせる
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # 次回のためにトークンを保存
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # 1. スプレッドシートからBCCリストを取得
    try:
        service_sheets = build('sheets', 'v4', credentials=creds)
        sheet = service_sheets.spreadsheets()
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
        values = result.get('values', [])

        bcc_emails = []
        for row in values:
            # S列（index 1）が存在するか確認
            if row and len(row) >= 2:
                email_cell = str(row[0])
                checkbox_cell = str(row[1]).strip().upper()
                
                # S列が TRUE の場合のみ取得
                if checkbox_cell == 'TRUE':
                    # 正規表現でセル内のすべての有効なメールアドレスを抽出
                    emails = re.findall(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', email_cell)
                    bcc_emails.extend(emails)

        # 重複を削除
        bcc_emails = list(set(bcc_emails))

        if not bcc_emails:
            print("❌ R列に有効なメールアドレスが見つかりませんでした。")
            return

        print(f"✅ スプレッドシートから {len(bcc_emails)} 件のアドレスを取得しました。")
    except Exception as e:
        print(f"❌ スプレッドシート読み込みエラー: {e}")
        return

    # 2. Gmailに下書きを作成
    try:
        service_gmail = build('gmail', 'v1', credentials=creds)
        message = EmailMessage()
        
        email_body = f"""お客様各位

平素より大変お世話になっております。
株式会社ランバード　広告運用サポート窓口です。

本メールは皆様へのお知らせのためBCCで送信しております。

月次の報告をさせていただきたく思い、日程調整のご連絡をいたしました。
下記日程より、3つほどご都合の良いお日にちを挙げて頂ければ幸いでございます。

【候補日時】
{candidate_dates}
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

        message.set_content(email_body)
        message['Subject'] = '【定例会日程調整】月次報告のご案内（株式会社ランバード）'
        message['To'] = 'creative.yt@runbird.net'
        message['Bcc'] = ', '.join(bcc_emails)

        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        create_message = {'message': {'raw': encoded_message}}
        
        draft = service_gmail.users().drafts().create(userId="me", body=create_message).execute()
        print(f"✅ Gmailの「下書き」フォルダにメールを作成しました！ (Draft ID: {draft['id']})")
    except Exception as e:
        print(f"❌ Gmail下書き作成エラー: {e}")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("使い方: python3 create_gmail_draft.py '・候補1\n・候補2'")
        sys.exit(1)
    main(sys.argv[1])
