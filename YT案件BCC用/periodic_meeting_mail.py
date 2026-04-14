"""
====================================================
定例会 日程調整メール 月初自動送信スクリプト
====================================================
【動作の流れ】
  1. PlaywrightでLINE WORKSカレンダーを自動スクリーンショット
  2. 画像を解析して空き時間を抽出（10:00-18:00 / 昼休み13-14時除外）
  3. Google SheetsのR列がTRUEの人のQ列メールアドレスを取得
  4. Macのメールアプリに宛先・BCC・本文を流し込んで表示（送信はしない）

【事前準備】
  pip3 install playwright google-api-python-client google-auth
  python3 -m playwright install chromium
  ※ google_credentials.json を同フォルダに設置
====================================================
"""

import os
import subprocess
import datetime
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ==========================================
# ★ 設定エリア（ここだけ編集）
# ==========================================
CONFIG = {
    # スプレッドシート設定
    "SPREADSHEET_ID": "1vxcDGBt0K58i7FQ7p-5KwabVBgCTIMYooswVglB2LK0",
    "SHEET_RANGE": "A2:R",
    "EMAIL_COL_INDEX": 16,       # Q列 (0始まり)
    "CHECKBOX_COL_INDEX": 17,    # R列 (0始まり) TRUE=対象

    # Google認証ファイル（スクリプトと同フォルダに置く）
    "GOOGLE_SERVICE_ACCOUNT_JSON": os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "google_credentials.json"
    ),

    # メール設定
    "TO_EMAIL": "frontnumber@icloud.com",
    "EMAIL_SUBJECT": "【定例会日程調整】月次報告のご案内（株式会社ランバード）",

    # LINE WORKSカレンダーURL（週表示）
    "LINEWORKS_CALENDAR_URL": "https://calendar.worksmobile.com/web/calendar/main?viewType=week",

    # 空き時間の検索設定
    "WORK_START": 10,    # 勤務開始時間
    "WORK_END": 18,      # 勤務終了時間
    "LUNCH_START": 13,   # 昼休み開始
    "LUNCH_END": 14,     # 昼休み終了
    "SEARCH_DAYS": 14,   # 今日から何日先まで探すか

    # スクリーンショット保存先
    "SCREENSHOT_DIR": os.path.join(os.path.dirname(os.path.abspath(__file__)), "calendar_screenshots"),
}

# ==========================================
# STEP 1: LINE WORKSカレンダー自動読み取り
# ==========================================

def fetch_free_slots_from_browser():
    """PlaywrightでLINE WORKSカレンダーを開き、既存セッションを使って空き時間を抽出"""
    from playwright.sync_api import sync_playwright

    os.makedirs(CONFIG["SCREENSHOT_DIR"], exist_ok=True)

    today = datetime.date.today()
    end_date = today + datetime.timedelta(days=CONFIG["SEARCH_DAYS"])

    # 対象の平日リストを生成
    target_weekdays = []
    for i in range(CONFIG["SEARCH_DAYS"] + 1):
        d = today + datetime.timedelta(days=i)
        if d.weekday() < 5:  # 月〜金
            target_weekdays.append(d)

    all_busy_periods = {}  # {date: [(start_hour, start_min, end_hour, end_min), ...]}

    with sync_playwright() as p:
        # ★ 既存のChromeプロファイル（ログイン済み）を使用
        # LINE WORKSにすでにログインしているChromeを利用
        user_data_dir = os.path.expanduser("~/Library/Application Support/Google/Chrome")
        
        try:
            browser = p.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=False,  # 画面を表示して動作確認
                slow_mo=500,
                channel="chrome",  # システムのChromeを使用
                args=["--no-first-run", "--no-default-browser-check"]
            )
            page = browser.pages[0] if browser.pages else browser.new_page()
        except Exception:
            # Chromeが使えない場合はPlaywright内蔵Chromiumを起動（要ログイン）
            print("⚠️ システムChromeが使えないため、Chromiumを起動します。LINE WORKSにログインしてください。")
            browser = p.chromium.launch(headless=False, slow_mo=500)
            page = browser.new_page()

        # 今週と来週の2週分をスクリーンショット
        screenshots = []
        page.goto(CONFIG["LINEWORKS_CALENDAR_URL"])
        page.wait_for_load_state("networkidle", timeout=15000)

        for week_num in range(2):  # 今週 + 来週
            if week_num > 0:
                # 「次の週」ボタンをクリック（>ボタン）
                next_btn = page.locator("button").filter(has_text=re.compile(r'^>$|次|next', re.IGNORECASE)).first
                if not next_btn.is_visible():
                    # 別のセレクタを試す
                    next_btn = page.locator('[aria-label*="next"], [aria-label*="次"], .lw-calendar-nav-next').first
                next_btn.click()
                page.wait_for_timeout(1500)

            # スクリーンショット保存
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = os.path.join(CONFIG["SCREENSHOT_DIR"], f"calendar_week{week_num+1}_{ts}.png")
            page.screenshot(path=screenshot_path, full_page=False)
            screenshots.append(screenshot_path)
            print(f"📸 スクリーンショット保存: {screenshot_path}")

        browser.close()

    print("\n✅ カレンダーのスクリーンショットを取得しました。")
    print("📋 スクリーンショットを確認して、空き時間を手動入力するか、AIに読み取らせてください。")

    return screenshots

def get_free_slots_text_manually():
    """
    スクリーンショット読み取りが難しい場合の手動入力モード。
    スクリプト実行時にターミナルで入力を受け付ける。
    """
    print("\n" + "="*50)
    print("📅 LINE WORKSカレンダーを確認して空き時間を入力してください")
    print("   形式: 4/14(月) 10:00~12:00 のように1行ずつ入力")
    print("   入力完了後、空行でEnterを押してください")
    print("="*50)

    lines = []
    while True:
        line = input("・").strip()
        if not line:
            break
        lines.append(f"・{line}")

    return "\n".join(lines)

# ==========================================
# STEP 2: Googleスプレッドシート → BCCリスト
# ==========================================

def get_bcc_list():
    """R列がTRUEの行のQ列メールアドレスを収集"""
    credentials_path = CONFIG["GOOGLE_SERVICE_ACCOUNT_JSON"]

    if not os.path.exists(credentials_path):
        raise FileNotFoundError(
            f"\n❌ Google認証ファイルが見つかりません: {credentials_path}\n"
            "Google Cloud Consoleでサービスアカウントを作成し、\n"
            "JSONキーを 'google_credentials.json' という名前でこのフォルダに保存してください。"
        )

    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds = service_account.Credentials.from_service_account_file(credentials_path, scopes=scopes)
    service = build('sheets', 'v4', credentials=creds)

    result = service.spreadsheets().values().get(
        spreadsheetId=CONFIG["SPREADSHEET_ID"],
        range=CONFIG["SHEET_RANGE"]
    ).execute()
    rows = result.get('values', [])

    emails = []
    for i, row in enumerate(rows):
        checkbox_col = CONFIG["CHECKBOX_COL_INDEX"]
        email_col = CONFIG["EMAIL_COL_INDEX"]

        checkbox_val = row[checkbox_col] if len(row) > checkbox_col else ""
        if str(checkbox_val).upper() != "TRUE":
            continue

        if len(row) > email_col:
            email = row[email_col].strip()
            if email and "@" in email:
                emails.append(email)
            else:
                print(f"⚠️ 行{i+2}: R列はTRUEですがQ列のメールが無効 → 「{email}」")

    print(f"✅ BCC対象者: {len(emails)}件")
    return emails

# ==========================================
# STEP 3: Macメールアプリに流し込み
# ==========================================

def create_mail_in_mac_mail(to_email, bcc_list, subject, body):
    """AppleScriptでMail.appを起動し、送信せずにメッセージウィンドウを表示"""

    # AppleScript用エスケープ
    safe_subject = subject.replace('\\', '\\\\').replace('"', '\\"')
    safe_body = body.replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n')

    # BCC全員を追加
    bcc_commands = "\n        ".join([
        f'make new bcc recipient at end of bcc recipients with properties {{address:"{addr}"}}'
        for addr in bcc_list
    ])

    applescript = f'''
tell application "Mail"
    set newMsg to make new outgoing message with properties {{subject:"{safe_subject}", content:"{safe_body}", visible:true}}
    tell newMsg
        make new to recipient at end of to recipients with properties {{address:"{to_email}"}}
        {bcc_commands}
    end tell
    activate
end tell
'''

    print("📬 Mail.appを起動中...")
    result = subprocess.run(["osascript", "-e", applescript], capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"AppleScriptエラー:\n{result.stderr}")

# ==========================================
# メイン実行
# ==========================================

def build_email_body(candidate_text):
    return (
        "お客様各位\n\n"
        "平素より大変お世話になっております。\n"
        "株式会社ランバード　広告運用サポート窓口です。\n\n"
        "本メールは皆様へのお知らせのためBCCで送信しております。\n\n"
        "月次の報告をさせていただきたく思い、日程調整のご連絡をいたしました。\n"
        "下記日程より、3つほどご都合の良いお日にちを挙げて頂ければ幸いでございます。\n\n"
        "【候補日時】\n"
        f"{candidate_text}\n"
        "※所要時間は30分ほどを予定しております。\n\n"
        "ご確認のほどよろしくお願いいたします。\n"
        "＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝\n"
        "株式会社ランバード\n"
        "広告運用サポート窓口\n"
        "東京都中央区晴海1丁目8－10\n"
        "晴海アイランド トリトンスクエア オフィスタワーX棟40階\n"
        "TEL　0120-949-851　FAX　03-5843-6978\n"
        "Mail   creative.yt@runbird.net\n"
        "HP　 https://runbird.net/\n"
        "＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝"
    )

if __name__ == "__main__":
    print("=" * 50)
    print("  定例会 日程調整メール 自動作成スクリプト")
    print(f"  実行日: {datetime.date.today().strftime('%Y年%m月%d日')}")
    print("=" * 50)

    try:
        # STEP 1: カレンダー読み取り
        print("\n【STEP 1】LINE WORKSカレンダーを読み取ります...")
        try:
            screenshots = fetch_free_slots_from_browser()
            print("\n📸 スクリーンショットを確認してから、空き時間を入力してください。")
            candidate_text = get_free_slots_text_manually()
        except Exception as browser_err:
            print(f"⚠️ ブラウザ自動読み取りに失敗: {browser_err}")
            print("手動入力モードに切り替えます。")
            candidate_text = get_free_slots_text_manually()

        if not candidate_text.strip():
            print("❌ 候補日時が入力されていません。終了します。")
            exit(1)

        # STEP 2: スプレッドシートから BCC リスト取得
        print("\n【STEP 2】スプレッドシートからBCC対象者を取得中...")
        bcc_list = get_bcc_list()

        if not bcc_list:
            print("❌ BCC対象者が0件です。スプレッドシートのR列（チェックボックス）を確認してください。")
            exit(1)

        # STEP 3: メールアプリ起動
        print(f"\n【STEP 3】Mail.appにメッセージを作成中... (BCC: {len(bcc_list)}件)")
        body = build_email_body(candidate_text)
        create_mail_in_mac_mail(CONFIG["TO_EMAIL"], bcc_list, CONFIG["EMAIL_SUBJECT"], body)

        print("\n" + "=" * 50)
        print("✅ 完了！Macのメールアプリを確認してください。")
        print("   ※メールは送信されていません。内容を確認して送信してください。")
        print("=" * 50)

    except FileNotFoundError as e:
        print(e)
    except Exception as e:
        print(f"\n❌ エラー: {e}")
