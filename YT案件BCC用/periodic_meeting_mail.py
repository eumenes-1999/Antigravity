import os
import subprocess
import time
import datetime
import jwt  # pyjwt
import httpx
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ==========================================
# 設定エリア
# ==========================================
CONFIG = {
    # スプレッドシート設定
    "SPREADSHEET_ID": "1vxcDGBt0K58i7FQ7p-5KwabVBgCTIMYooswVglB2LK0",
    "SHEET_RANGE": "A2:R",       # R列まで取得
    "EMAIL_COL_INDEX": 16,       # Q列 (0始まり) → メールアドレス
    "CHECKBOX_COL_INDEX": 17,    # R列 (0始まり) → TRUE/FALSEのチェックボックス
    
    # 認証用設定 (環境変数または直接入力)
    "GOOGLE_SERVICE_ACCOUNT_JSON": "google_credentials.json",
    
    # LINE WORKS 設定
    "LW_CLIENT_ID": os.environ.get("LW_CLIENT_ID", "YOUR_CLIENT_ID"),
    "LW_CLIENT_SECRET": os.environ.get("LW_CLIENT_SECRET", "YOUR_CLIENT_SECRET"),
    "LW_SERVICE_ACCOUNT": os.environ.get("LW_SERVICE_ACCOUNT", "YOUR_SERVICE_ACCOUNT"),
    "LW_PRIVATE_KEY": os.environ.get("LW_PRIVATE_KEY", "").replace("\\n", "\n"), # \nを改行コードに変換
    "LW_USER_ID": os.environ.get("LW_USER_ID", "YOUR_USER_ID"),

    # 空き時間検索設定
    "SEARCH_DAYS": 14,
    "WORK_START": 10,
    "WORK_END": 18,
    "LUNCH_START": 13,
    "LUNCH_END": 14,
}

# ==========================================
# LINE WORKS連携 (空き時間抽出)
# ==========================================

def get_lw_access_token():
    now = int(time.time())
    payload = {
        "iss": CONFIG["LW_CLIENT_ID"],
        "sub": CONFIG["LW_SERVICE_ACCOUNT"],
        "iat": now,
        "exp": now + 3600
    }
    # JWTの生成 (RS256)
    token = jwt.encode(payload, CONFIG["LW_PRIVATE_KEY"], algorithm="RS256")
    
    url = "https://auth.worksmobile.com/oauth2/v2.0/token"
    data = {
        "assertion": token,
        "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
        "client_id": CONFIG["LW_CLIENT_ID"],
        "client_secret": CONFIG["LW_CLIENT_SECRET"],
        "scope": "calendar.read"
    }
    res = httpx.post(url, data=data)
    res.raise_for_status()
    return res.json()["access_token"]

def fetch_free_slots():
    token = get_lw_access_token()
    user_id = CONFIG["LW_USER_ID"]
    
    start_dt = datetime.datetime.now()
    end_dt = start_dt + datetime.timedelta(days=CONFIG["SEARCH_DAYS"])
    
    # LINE WORKS API用のISO8601形式
    range_start = start_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    range_end = end_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    
    url = f"https://www.worksapis.com/v1.0/users/{user_id}/calendar/events"
    params = {
        "rangeStartTime": range_start,
        "rangeEndTime": range_end,
        "limit": 100
    }
    headers = {"Authorization": f"Bearer {token}"}
    res = httpx.get(url, params=params, headers=headers)
    res.raise_for_status()
    events = res.json().get("events", [])

    free_slots = []
    # 抽出ロジック
    for i in range(CONFIG["SEARCH_DAYS"] + 1):
        target_date = (start_dt + datetime.timedelta(days=i)).date()
        if target_date.weekday() >= 5: continue # 土日スキップ
        
        # その日の予定を取得・JSTに変換
        day_events = []
        for ev in events:
            s_str = ev["start"].get("dateTime") or ev["start"].get("date")
            e_str = ev["end"].get("dateTime") or ev["end"].get("date")
            # 簡易的にJST固定で処理
            ev_start = datetime.datetime.fromisoformat(s_str.replace("Z", "+00:00")).astimezone(datetime.timezone(datetime.timedelta(hours=9)))
            ev_end = datetime.datetime.fromisoformat(e_str.replace("Z", "+00:00")).astimezone(datetime.timezone(datetime.timedelta(hours=9)))
            
            if ev_start.date() == target_date:
                day_events.append((ev_start.replace(tzinfo=None), ev_end.replace(tzinfo=None)))
        
        day_events.sort()

        # 勤務時間のブロック定義
        windows = [
            (datetime.datetime.combine(target_date, datetime.time(CONFIG["WORK_START"])), 
             datetime.datetime.combine(target_date, datetime.time(CONFIG["LUNCH_START"]))),
            (datetime.datetime.combine(target_date, datetime.time(CONFIG["LUNCH_END"])), 
             datetime.datetime.combine(target_date, datetime.time(CONFIG["WORK_END"])))
        ]

        for w_start, w_end in windows:
            curr = w_start
            for ev_s, ev_e in day_events:
                if ev_e <= curr or ev_s >= w_end: continue
                if ev_s > curr:
                    if (ev_s - curr).total_seconds() >= 1800:
                        free_slots.append(f"・{target_date.strftime('%m/%d')}({['月','火','水','木','金','土','日'][target_date.weekday()]}) {curr.strftime('%H:%M')}~{ev_s.strftime('%H:%M')}")
                curr = max(curr, ev_e)
            if curr < w_end:
                if (w_end - curr).total_seconds() >= 1800:
                    free_slots.append(f"・{target_date.strftime('%m/%d')}({['月','火','水','木','金','土','日'][target_date.weekday()]}) {curr.strftime('%H:%M')}~{w_end.strftime('%H:%M')}")

    return "\n".join(free_slots)

# ==========================================
# Google Sheets連携 (BCCリスト抽出)
# ==========================================

def get_bcc_list():
    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    if not os.path.exists(CONFIG["GOOGLE_SERVICE_ACCOUNT_JSON"]):
        raise FileNotFoundError(f"Google認証ファイル '{CONFIG['GOOGLE_SERVICE_ACCOUNT_JSON']}' が見つかりません。")
        
    creds = service_account.Credentials.from_service_account_file(CONFIG["GOOGLE_SERVICE_ACCOUNT_JSON"], scopes=scopes)
    service = build('sheets', 'v4', credentials=creds)
    
    result = service.spreadsheets().values().get(
        spreadsheetId=CONFIG["SPREADSHEET_ID"], 
        range=CONFIG["SHEET_RANGE"]
    ).execute()
    rows = result.get('values', [])
    
    emails = []
    for row in rows:
        # R列(インデックス17)がTRUEの行のみ対象
        checkbox_val = row[CONFIG["CHECKBOX_COL_INDEX"]] if len(row) > CONFIG["CHECKBOX_COL_INDEX"] else False
        is_checked = str(checkbox_val).upper() == "TRUE"
        
        if not is_checked:
            continue
        
        # Q列(インデックス16)のメールアドレスを取得
        if len(row) > CONFIG["EMAIL_COL_INDEX"]:
            email = row[CONFIG["EMAIL_COL_INDEX"]].strip()
            if email and "@" in email:
                emails.append(email)
    
    print(f"BCC対象者: {len(emails)}件")
    return emails

# ==========================================
# AppleScript 連携 (Mail.app)
# ==========================================

def create_mail_in_mac_mail(to_email, bcc_list, subject, body):
    # AppleScript用にエスケープ処理
    safe_subject = subject.replace('"', '\\"')
    safe_body = body.replace('"', '\\"').replace('\n', '\\n')
    
    # BCC全件をループで追加するコマンドを生成
    bcc_commands = "\n".join([
        f'make new bcc recipient at end of bcc recipients with properties {{address:"{addr}"}}'
        for addr in bcc_list
    ])
    
    applescript = f'''
tell application "Mail"
    set newMessage to make new outgoing message with properties {{subject:"{safe_subject}", content:"{safe_body}", visible:true}}
    tell newMessage
        make new to recipient at end of to recipients with properties {{address:"{to_email}"}}
        {bcc_commands}
    end tell
    activate
end tell
'''
    
    result = subprocess.run(["osascript", "-e", applescript], capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"AppleScriptエラー: {result.stderr}")

# ==========================================
# メイン実行
# ==========================================

if __name__ == "__main__":
    print("LINE WORKSとスプレッドシートから情報を取得しています...")
    try:
        candidate_text = fetch_free_slots()
        bcc_list = get_bcc_list()
        
        subject = "【定例会日程調整】月次報告のご案内（株式会社ランバード）"
        to_email = "frontnumber@icloud.com"  # 宛名欄（Toアドレス）
        body = f"""お客様各位

平素より大変お世話になっております。
株式会社ランバード　広告運用サポート窓口です。

本メールは皆様へのお知らせのためBCCで送信しております。

月次の報告をさせていただきたく思い、日程調整のご連絡をいたしました。
下記日程より、3つほどご都合の良いお日にちを挙げて頂ければ幸いでございます。

【候補日時】
{candidate_text}
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

        print(f"Mail.appを起動してメッセージを作成します... (宛先: {len(bcc_list)}件)")
        create_mail_in_mac_mail(to_email, bcc_list, subject, body)
        print("完了しました。メールアプリを確認してください。")

    except Exception as e:
        print(f"\nエラーが発生しました: {e}")
        print("\n[確認事項]")
        print("1. google_credentials.json がこのファイルと同じ場所にありますか？")
        print("2. 環境変数（LW_CLIENT_ID等）が設定されているか、CONFIGを直接書き換えていますか？")
        print("3. pip install pyjwt cryptography httpx google-api-python-client google-auth を実行しましたか？")
