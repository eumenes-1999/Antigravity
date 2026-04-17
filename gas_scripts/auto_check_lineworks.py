import os
import re
import time
import requests
from playwright.sync_api import sync_playwright

GAS_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbyLOinVRFQ8YJ5HxtsIt5ISsdgsgOsXNEnyQ4dRRuPDIxoW3N3DPZCzvzVuEuk_rg/exec"
TARGET_GROUP = "MEO×運用"
AUTH_FILE = "lineworks_auth.json"

def main():
    print("🚀 LINE WORKS 自動巡回エージェントを起動します...")
    print("--------------------------------------------------")
    
    with sync_playwright() as p:
        # ブラウザを画面表示あり(headless=False)で起動します。
        browser = p.chromium.launch(headless=False, args=['--window-size=1200,800'])
        context_args = {"viewport": {'width': 1200, 'height': 800}}
        
        # 以前ログインしたセッション情報があれば読み込んで「自動ログイン」します
        if os.path.exists(AUTH_FILE):
            print("🔑 保存されたログイン情報（Cookie）を読み込みます...")
            context_args["storage_state"] = AUTH_FILE
            
        context = browser.new_context(**context_args)
        page = context.new_page()
        
        # 1. LINE WORKSへアクセス
        print("🌐 LINE WORKS (https://talk.worksmobile.com/) にアクセス中...")
        page.goto("https://talk.worksmobile.com/")
        
        # ログイン画面か、既にログイン済みかを判定
        try:
            # 「トーク」等の文字が画面に出るまで最大10秒待機
            page.wait_for_function('document.body.innerText.includes("トーク") || document.body.innerText.includes("Contacts")', timeout=10000)
            print("✅ ログイン済みのメイン画面を確認しました！")
        except:
            print("⚠️ 【アクション要求】自動ログインできませんでした。")
            print("👉 開いたブラウザ画面で LINE WORKS にログインしてください。")
            print("⏳ ログインが完了するまで待機しています... (なんと最大3分待ちます！)")
            try:
                page.wait_for_function('document.body.innerText.includes("トーク") || document.body.innerText.includes("Contacts")', timeout=180000)
                # ログインに成功したら、次から自動ログインできるように認証情報を保存！
                context.storage_state(path=AUTH_FILE)
                print("💾 ログイン情報を自動保存しました！★次回からは自動ログインになります！★")
            except:
                print("❌ ログインタイムアウトしました。プログラムを終了します。再度実行してください。")
                browser.close()
                return
        
        # 2. 対象グループを選択
        print(f"\n🔍 グループ「{TARGET_GROUP}」を探して開きます...")
        time.sleep(3) # ロード待ち
        try:
            # グループ名を持つ要素をクリック
            group_elements = page.get_by_text(TARGET_GROUP).all()
            if group_elements:
                group_elements[0].click()
                print(f"✅ 「{TARGET_GROUP}」グループを開きました！")
                time.sleep(4) # メッセージがロードされるのを少し待つ
        except Exception:
            pass
            
        print("⏳ （※もし画面上で対象のグループが開かれていない場合は、ご自身でクリックして開いてください。10秒間待機します...）")
        time.sleep(10)
        
        # 3. 画面に映っている情報から特定メッセージ構造だけを賢く抽出 (DOM非依存ハック)
        print("\n📥 表示されているチャットから「MEO運用依頼」のテキストを抽出しています...")
        
        # "MEO運用依頼" いう文字を持つ要素を全て取得
        elements = page.get_by_text("MEO運用依頼").all()
        if not elements:
            print("❌ 画面に「MEO運用依頼」の文字が見当たりません！エラー終了します。")
            browser.close()
            return
            
        # 画面に映っている一番下（最新）の依頼をターゲットにする
        target_node = elements[-1]
        
        target_message = None
        # 対象の文字を含む「チャット吹き出し（親要素）」を特定するために、DOMツリーを下に少しずつ遡るAI的アプローチ
        for i in range(1, 15):
            try:
                # i階層上の親要素を取得
                container = target_node.evaluate_handle(f"(el) => {{ let parent = el; for(let j=0; j<{i}; j++) {{ if(parent.parentElement) parent = parent.parentElement; }} return parent; }}")
                text = container.evaluate("(el) => el.innerText")
                if not text: continue
                    
                # その親要素の箱の中に「希望キーワード」まで含まれていれば、それは1つのチャット吹き出しであると判断
                if "希望キーワード" in text:
                    # 階層を遡りすぎて「画面全体のテキスト」になっていないか（"MEO運用依頼"が1つだけ含まれるか）をチェック
                    if text.count("MEO運用依頼") == 1:
                        target_message = text.strip()
                        break
            except:
                continue
                    
        if target_message:
            print("✨ 抽出大成功！以下の情報をシートへ転送します：\n")
            lines = target_message.split("\n")
            for line in lines[:5]: print(f"  {line}")
            print("  ... (中略) ...\n")
            
            # 4. GASのWebhook URLへPOST送信（通信ケーブルの役割）
            print("🚀 GASへデータを全自動送信中...")
            payload = {"text": target_message}
            response = requests.post(GAS_WEBHOOK_URL, json=payload, headers={'Content-Type': 'application/json'})
            
            if response.status_code == 200:
                print(f"🎉 送信完了・スプレッドシートへの書き込み成功！！")
                print(f"   (GASからの応答: {response.text})")
            else:
                print(f"❌ 通信エラーが発生しました: {response.status_code} - {response.text}")
        else:
            print("⚠️ メッセージのフォーマットが変則的なため、取得領域をうまく特定できませんでした。")
            
        print("--------------------------------------------------")
        print("✅ 全ての自動処理が完了しました！ブラウザを閉じて終了します。")
        time.sleep(2)
        browser.close()

if __name__ == "__main__":
    main()
