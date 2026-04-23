#!/usr/bin/env python3
"""
管理スプレッドシートに meo_automation.js の全文を「GASソース_meo_automation」シートへ書き込む（バックアップ用）。

前提:
  - pip install gspread google-auth
  - サービスアカウント JSON を用意し、スプレッドシートをそのメールに「編集者」で共有

認証ファイル:
  環境変数 MEO_GOOGLE_CREDENTIALS に JSON のパス、または同フォルダの google_credentials.json
"""
import os
import re
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
JS_PATH = os.path.join(SCRIPT_DIR, "meo_automation.js")
BACKUP_SHEET_TITLE = "GASソース_meo_automation"


def read_spreadsheet_id():
    with open(JS_PATH, encoding="utf-8") as f:
        m = re.search(r"SPREADSHEET_ID\s*=\s*'([^']+)'", f.read())
    if not m:
        sys.exit("SPREADSHEET_ID を meo_automation.js から読み取れませんでした")
    return m.group(1)


def main():
    cred_path = os.environ.get(
        "MEO_GOOGLE_CREDENTIALS",
        os.path.join(SCRIPT_DIR, "google_credentials.json"),
    )
    if not os.path.isfile(cred_path):
        print(
            "認証用 JSON が見つかりません。次を実施してください:\n"
            f"  1) Google Cloud でサービスアカウントを作成し JSON をダウンロード\n"
            f"  2) その JSON を次のいずれかに置く:\n"
            f"     - {cred_path}\n"
            f"     - または環境変数 MEO_GOOGLE_CREDENTIALS にパスを指定\n"
            f"  3) スプレッドシートをサービスアカウントのメール（client_email）に「編集者」で共有\n"
            f"  4) pip install gspread google-auth\n"
            f"  5) 再度このスクリプトを実行\n"
        )
        sys.exit(1)

    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        print("pip install gspread google-auth を実行してください。")
        sys.exit(1)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(cred_path, scopes=scopes)
    gc = gspread.authorize(creds)

    ss_id = read_spreadsheet_id()
    sh = gc.open_by_key(ss_id)

    with open(JS_PATH, encoding="utf-8") as f:
        lines = f.read().splitlines()
    rows = [[line] for line in lines]

    try:
        ws = sh.worksheet(BACKUP_SHEET_TITLE)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(
            title=BACKUP_SHEET_TITLE,
            rows=max(len(rows) + 20, 200),
            cols=2,
        )

    ws.clear()
    if rows:
        ws.update("A1", rows, value_input_option="RAW")

    print(f"完了: {len(rows)} 行をシート「{BACKUP_SHEET_TITLE}」の A 列に書き込みました。")
    print(f"https://docs.google.com/spreadsheets/d/{ss_id}/edit")


if __name__ == "__main__":
    main()
