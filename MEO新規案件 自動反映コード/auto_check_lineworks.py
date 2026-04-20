import os
import re
import time
import requests
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

GAS_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbyLOinVRFQ8YJ5HxtsIt5ISsdgsgOsXNEnyQ4dRRuPDIxoW3N3DPZCzvzVuEuk_rg/exec"
TARGET_GROUP = "MEO×運用"
AUTH_FILE = "lineworks_auth.json"
TEMPLATE_CONFIRM_TEXT = "内容を確認する"
TEMPLATE_HOST = "template.worksmobile.com"
# トーク／テンプレで対象とするテンプレート名（表記ゆれを許容）
TEMPLATE_TITLE_FULL = "MEO運用依頼 (営業→運用)"
TEMPLATE_TITLE_FULLWIDTH_PAREN = "MEO運用依頼（営業→運用）"


def _body_has_exact_target_template_title(body_text: str) -> bool:
    """自動クリックの前提: トーク上に対象テンプレ名そのものが見えているか（別テンプレ誤クリック防止）。"""
    if TEMPLATE_TITLE_FULL in body_text or TEMPLATE_TITLE_FULLWIDTH_PAREN in body_text:
        return True
    compact = body_text.replace(" ", "").replace("\u3000", "")
    return "MEO運用依頼(営業→運用)" in compact


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


def _click_confirm_inside_exact_meo_card(talk_page):
    """
    「MEO運用依頼 (営業→運用)」という表記が同一メッセージ塊に含まれる「内容を確認する」だけをクリックする。
    「MEO運用依頼」と「営業→運用」だけの緩い一致は使わない（別テンプレ誤爆防止）。
    """
    return talk_page.evaluate(
        """([full, fullw, label]) => {
            const exactTitle = (t) =>
                t.includes(full) ||
                t.includes(fullw) ||
                t.includes("MEO運用依頼(営業→運用)");
            const candidates = [];
            const nodes = document.querySelectorAll(
                "a, button, [role='button'], span, div, p"
            );
            for (const el of nodes) {
                const raw = (el.innerText || "").replace(/\\s+/g, " ").trim();
                if (!raw.includes(label)) continue;
                if (raw.length > 80) continue;

                let bestLen = Infinity;
                let p = el;
                for (let d = 0; d < 45 && p; d++) {
                    const block = p.innerText || "";
                    if (
                        block.includes(label) &&
                        exactTitle(block) &&
                        block.length < 12000
                    )
                        bestLen = Math.min(bestLen, block.length);
                    p = p.parentElement;
                }
                if (bestLen < Infinity)
                    candidates.push({
                        el,
                        bestLen,
                        top: el.getBoundingClientRect().top,
                    });
            }
            if (!candidates.length) return { ok: false, reason: "no_exact_title_candidates" };

            candidates.sort(
                (a, b) => a.bestLen - b.bestLen || b.top - a.top
            );
            const chosen = candidates[0];
            chosen.el.scrollIntoView({ block: "center", inline: "nearest" });
            chosen.el.click();
            return {
                ok: true,
                picked_block_len: chosen.bestLen,
                candidate_count: candidates.length,
            };
        }""",
        [TEMPLATE_TITLE_FULL, TEMPLATE_TITLE_FULLWIDTH_PAREN, TEMPLATE_CONFIRM_TEXT],
    )


def _open_template_page(context, talk_page):
    """
    対象テンプレのみを開く想定でクリックし、template.worksmobile.com の Page を返す。
    popup / 新タブ / 同一タブ遷移のいずれにも対応するため、クリック前後の context.pages を比較する。
    """
    pages_before = list(context.pages)

    def _do_click():
        picked = _click_confirm_inside_exact_meo_card(talk_page)
        if not picked.get("ok"):
            print(
                f"⚠️ 厳密タイトル付きカード内の「{TEMPLATE_CONFIRM_TEXT}」を特定できませんでした（{picked.get('reason')}）。"
                "Playwright の locator で同一タイトル＋ボタンに絞り込みます。"
            )
            talk_page.locator("div, article, section, li").filter(
                has_text=TEMPLATE_TITLE_FULL
            ).filter(has_text=TEMPLATE_CONFIRM_TEXT).last.get_by_text(
                TEMPLATE_CONFIRM_TEXT
            ).click(timeout=15000)
        else:
            print(
                f"   → 「{TEMPLATE_TITLE_FULL}」付近の「{TEMPLATE_CONFIRM_TEXT}」をクリック "
                f"(候補数={picked.get('candidate_count')}, ブロック長≈{picked.get('picked_block_len')})"
            )

    tpl = None
    try:
        with context.expect_page(timeout=45000) as new_page_info:
            _do_click()
        tpl = new_page_info.value
    except PlaywrightTimeoutError:
        # クリックは1回のみ。同一タブ遷移などで新 Page イベントが来ない場合は下のポーリングへ。
        tpl = None

    deadline = time.time() + 45
    while time.time() < deadline and tpl is None:
        for p in context.pages:
            if p not in pages_before:
                try:
                    u = p.url or ""
                except Exception:
                    continue
                if TEMPLATE_HOST in u:
                    tpl = p
                    break
        if tpl is None:
            try:
                if TEMPLATE_HOST in (talk_page.url or ""):
                    tpl = talk_page
                    break
            except Exception:
                pass
        time.sleep(0.2)

    if tpl is None:
        raise RuntimeError(
            f"{TEMPLATE_HOST} のページが検出できませんでした（ポップアップ／同一タブどちらも不明）。"
        )

    tpl.wait_for_load_state("domcontentloaded")
    tpl.wait_for_url(f"**/{TEMPLATE_HOST}/**", timeout=90000)
    return tpl


def _wait_meo_template_page_ready(tpl_page, timeout_ms: int = 90000) -> None:
    """template.worksmobile.com 上で「MEO運用依頼 (営業→運用)」と依頼フィールドが描画されるまで待つ。"""
    tpl_page.wait_for_function(
        r"""() => {
            const t = document.body.innerText || "";
            const title =
                t.includes("MEO運用依頼 (営業→運用)") ||
                t.includes("MEO運用依頼（営業→運用）") ||
                t.includes("MEO運用依頼(営業→運用)");
            const fields =
                t.includes("法人名") &&
                (t.includes("危険度") || t.includes("キーワード"));
            return title && fields;
        }""",
        timeout=timeout_ms,
    )


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
        print("⏳ 【重要】対象はテンプレート「MEO運用依頼 (営業→運用)」です。")
        print("          トーク上でそのタイトルと「内容を確認する」が両方見える位置までスクロールしてください。")
        print("          (最大60秒間、トーク上の本文または上記テンプレ用ボタンの表示を監視します...)")

        source_page = page
        ready = False
        hinted_no_title = False
        for _ in range(30):
            body_text = page.evaluate("() => document.body.innerText")
            if _body_has_request_fields(body_text):
                print("✅ トーク画面上に依頼本文（法人名＋危険度/キーワード）を確認しました！")
                time.sleep(1)
                ready = True
                break
            if TEMPLATE_CONFIRM_TEXT in body_text and _body_has_exact_target_template_title(
                body_text
            ):
                print(
                    f"✅ 「{TEMPLATE_TITLE_FULL}」と「{TEMPLATE_CONFIRM_TEXT}」を検出。"
                    "テンプレ詳細を開きます…"
                )
                try:
                    source_page = _open_template_page(context, page)
                    print(f"✅ テンプレ詳細を開きました ({TEMPLATE_HOST})")
                    print(
                        f"⏳ 「{TEMPLATE_TITLE_FULL}」の本文（法人名・危険度/キーワード）が表示されるまで待機します…"
                    )
                    _wait_meo_template_page_ready(source_page)
                    print("✅ テンプレ本文の表示を確認しました。")
                    time.sleep(1)
                    ready = True
                    break
                except Exception as e:
                    print(f"⚠️ テンプレ詳細の取得または読み込みに失敗しました: {e}")
                    print(
                        "   トークで「MEO運用依頼 (営業→運用)」が見えるカードの「内容を確認する」を"
                        "手動で押し、開いたウィンドウに本文が出るまで待ってから再実行してください。"
                    )
            elif TEMPLATE_CONFIRM_TEXT in body_text and not _body_has_exact_target_template_title(
                body_text
            ):
                if not hinted_no_title:
                    print(
                        f"⏳ 「{TEMPLATE_CONFIRM_TEXT}」は見えていますが、"
                        f"「{TEMPLATE_TITLE_FULL}」の表記がトーク上に見えません（略称だけのカードは対象外）。"
                        "該当テンプレ名が全文表示される位置までスクロールしてください。"
                    )
                    hinted_no_title = True
            time.sleep(2)
        else:
            print(
                "⚠️ 60秒待機しましたが、トーク上で「MEO運用依頼 (営業→運用)」の全文表記と「内容を確認する」の組み合わせ、"
                "または「法人名」と「危険度/キーワード」を揃えた本文を確認できませんでした。"
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
