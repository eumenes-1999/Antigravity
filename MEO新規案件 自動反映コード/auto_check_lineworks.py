import os
import re
import time
import requests
from playwright.sync_api import sync_playwright

GAS_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbyLOinVRFQ8YJ5HxtsIt5ISsdgsgOsXNEnyQ4dRRuPDIxoW3N3DPZCzvzVuEuk_rg/exec"
TARGET_GROUP = "MEO×運用"
AUTH_FILE = "lineworks_auth.json"
TEMPLATE_CONFIRM_TEXT = "内容を確認する"
TEMPLATE_HOST = "template.worksmobile.com"


def _body_has_request_fields(body_text: str) -> bool:
    if "法人名" not in body_text:
        return False
    return "危険度" in body_text or "キーワード" in body_text


def _extract_target_message(page):
    """指定ページから依頼本文を抽出する（トーク／テンプレ詳細の両方で共通）。"""
    body_text = page.evaluate("() => document.body.innerText")
    target_message = page.evaluate(r'''() => {
            const hasFields = (t) =>
                t.includes("法人名") &&
                (t.includes("危険度") || t.includes("キーワード"));
            let elements = document.querySelectorAll('*');
            let candidates = [];
            for (let el of elements) {
                let text = el.innerText || "";
                if (hasFields(text) && text.length < 5000) {
                    candidates.push(el);
                }
            }
            if (candidates.length === 0) return null;

            let lastEl = candidates[candidates.length - 1];
            let innermost = lastEl;
            let foundDeeper = true;
            while (foundDeeper) {
                foundDeeper = false;
                for (let child of innermost.children) {
                    let text = child.innerText || "";
                    if (hasFields(text)) {
                        innermost = child;
                        foundDeeper = true;
                        break;
                    }
                }
            }
            return innermost.innerText.trim();
        }''')

    if not target_message:
        matches = re.findall(
            r'(法人名[\s\S]*?(?:危険度|キーワード)[^\n]*(?:\n[^\n]+){0,10})',
            body_text,
        )
        if matches:
            target_message = matches[-1].strip()
    return target_message, body_text


def _open_template_popup(talk_page):
    """トーク上の「内容を確認する」からテンプレ詳細ウィンドウを開き、その Page を返す。"""
    with talk_page.expect_popup(timeout=30000) as popup_info:
        talk_page.get_by_text(TEMPLATE_CONFIRM_TEXT).last.click()
    tpl = popup_info.value
    tpl.wait_for_load_state("domcontentloaded")
    tpl.wait_for_url(f"**/{TEMPLATE_HOST}/**", timeout=60000)
    return tpl


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
            
        print("⏳ （※もし画面上で対象のグループが開かれていない場合は開いてください。）")
        print("⏳ 【重要】目的の「MEO運用依頼」のメッセージが昔のもので画面外にある場合は、")
        print("          ブラウザ上で『上にスクロール』して画面内に表示させてください！")
        print("          (最大60秒間、「内容を確認する」または依頼本文の表示を監視します...)")

        source_page = page
        ready = False
        for _ in range(30):
            body_text = page.evaluate("() => document.body.innerText")
            if _body_has_request_fields(body_text):
                print("✅ トーク画面上に依頼本文（法人名＋危険度/キーワード）を確認しました！")
                time.sleep(1)
                ready = True
                break
            if TEMPLATE_CONFIRM_TEXT in body_text:
                print(f"✅ 「{TEMPLATE_CONFIRM_TEXT}」を検出。テンプレ詳細ウィンドウを開きます…")
                try:
                    source_page = _open_template_popup(page)
                    print(f"✅ テンプレ詳細を読み込みました ({TEMPLATE_HOST})")
                    time.sleep(1)
                    ready = True
                    break
                except Exception as e:
                    print(f"⚠️ テンプレ詳細ウィンドウの取得に失敗しました: {e}")
                    print("   手動で「内容を確認する」を押し、ウィンドウが開くまで待ってから再実行してください。")
            time.sleep(2)
        else:
            print(
                "⚠️ 60秒待機しましたが、「内容を確認する」も、"
                "「法人名」と「危険度/キーワード」を揃えた本文もトーク上で確認できませんでした。"
            )

        # 3. トークまたはテンプレ詳細ページから本文を抽出
        print(
            "\n📥 表示中のページから「法人名」と「危険度/キーワード」を含むテキストを抽出しています..."
        )

        target_message, body_text = _extract_target_message(source_page)

        if not target_message:
            print(f"【DEBUG情報】画面全体の取得文字数: {len(body_text)}文字")
            print(f"  -> 「法人名」: {'法人名' in body_text}")
            print(f"  -> 「危険度」: {'危険度' in body_text}")
            print(f"  -> 「キーワード」: {'キーワード' in body_text}")
            if not ready:
                print("【DEBUG情報】正規表現での抽出も失敗しました。画面情報の最初と最後を少し表示します：")
                print("---------------------------------")
                stripped = body_text.strip()
                if len(stripped) > 500:
                    print(stripped[:250] + "\n...(省略)...\n" + stripped[-250:])
                else:
                    print(stripped)
                print("---------------------------------")
            else:
                print("【DEBUG情報】正規表現での抽出も失敗しました（テンプレ側の DOM 構造の可能性あり）。")

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
