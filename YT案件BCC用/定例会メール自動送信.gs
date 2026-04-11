/**
 * =====================================================
 * 定例会 日程調整メール BCC一括送信スクリプト
 * 株式会社ランバード 広告運用サポート窓口
 * =====================================================
 *
 * 【処理の流れ】
 *   1. LINE WORKSカレンダーから「候補」タグのイベントを取得
 *   2. スプレッドシートのQ列からクライアントのメールアドレスを収集
 *   3. 全クライアント宛にBCCで日程調整メールを一括送信
 *
 * 【LINE WORKSカレンダーの使い方】
 *   送信したい候補日時をLINE WORKSカレンダーに登録してください。
 *   イベントタイトルに「候補」という文字を含めてください。
 *   例: 「定例_候補」「候補日時」など
 *
 * 【スクリプトプロパティに登録する設定値】
 *   GASエディタ → プロジェクトの設定 → スクリプトプロパティ に追加
 *
 *   LW_CLIENT_ID       : LINE WORKS Developer Console の Client ID
 *   LW_CLIENT_SECRET   : LINE WORKS Developer Console の Client Secret
 *   LW_SERVICE_ACCOUNT : LINE WORKS サービスアカウント (xxx@works-bot.com)
 *   LW_PRIVATE_KEY     : 秘密鍵の全文 (-----BEGIN PRIVATE KEY----- から末尾まで)
 *   LW_USER_ID         : カレンダーを持つLINE WORKSユーザーID (メールアドレス形式)
 * =====================================================
 */

// =====================================================
// ★ ここだけ編集してください（設定エリア）
// =====================================================
const CONFIG = {

  // スプレッドシートID（URLの /d/〇〇〇/ の部分）
  SPREADSHEET_ID: '1vxcDGBt0K58i7FQ7p-5KwabVBgCTIMYooswVglB2LK0',

  // メールアドレスが入っているシートのGID
  SHEET_GID: '0',

  // メールアドレスの列番号（Q列 = 17）
  EMAIL_COLUMN: 17,

  // データ開始行（1行目がヘッダーなら 2 を指定）
  DATA_START_ROW: 2,

  // --- 空き時間検索の設定 ---
  SEARCH_DAYS_AHEAD: 14,      // 今日から何日間先まで探すか
  WORK_START_HOUR: 10,       // 開始時間 (10時)
  WORK_END_HOUR: 18,         // 終了時間 (18時)
  LUNCH_START_HOUR: 13,      // 休憩開始 (13時)
  LUNCH_END_HOUR: 14,        // 休憩終了 (14時)
  MIN_SLOT_MINUTES: 30,      // 最低限必要な空き時間（分）

  // 送信モード
  //   true  → Gmailの下書きに保存（送信前に確認したい場合）
  //   false → 即時送信
  DRAFT_ONLY: true,

  // 差出人のメールアドレス
  FROM_EMAIL: 'creative.yt@runbird.net',

  // メール件名
  EMAIL_SUBJECT: '【定例会日程調整】月次報告のご案内（株式会社ランバード）',

  // メール本文テンプレート
  EMAIL_BODY: `お客様各位

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
＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝`,
};

// =====================================================
// メイン処理（この関数を実行してください）
// =====================================================
function sendSchedulingEmail() {
  try {
    Logger.log('========================================');
    Logger.log('定例会 日程調整メール送信処理 開始');
    Logger.log('========================================');

    // --- STEP 1: LINE WORKSから空き時間（候補日時）を取得 ---
    Logger.log('\n[STEP 1] LINE WORKSカレンダーから空き時間を抽出中...');
    const freeSlots = getFreeSlotsFromLineWorks();

    if (freeSlots.length === 0) {
      const msg = `指定された期間内に空き時間が見つかりませんでした。`;
      Logger.log('⚠️ ' + msg);
      Browser.msgBox('空き時間なし', msg, Browser.Buttons.OK);
      return;
    }

    Logger.log(`✅ 空き枠抽出完了: ${freeSlots.length}件`);
    const candidateDateText = formatCandidateDates(freeSlots);
    Logger.log('\n--- 生成された候補日時テキスト ---');
    Logger.log(candidateDateText);

    // --- STEP 2: スプレッドシートからメールアドレスを取得 ---
    Logger.log('\n[STEP 2] スプレッドシートからクライアントメールを取得中...');
    const clientEmails = getClientEmailsFromSheet();

    if (clientEmails.length === 0) {
      const msg = 'スプレッドシートのQ列にメールアドレスが見つかりませんでした。\nシートとGIDの設定を確認してください。';
      Logger.log('⚠️ ' + msg);
      Browser.msgBox('メールアドレスが見つかりません', msg, Browser.Buttons.OK);
      return;
    }

    Logger.log(`✅ クライアントメール取得完了: ${clientEmails.length}件`);
    Logger.log('送信先: ' + clientEmails.join(', '));

    // --- STEP 3: メール作成・送信 ---
    Logger.log('\n[STEP 3] メール作成・送信中...');
    const emailBody = CONFIG.EMAIL_BODY.replace('{CANDIDATE_DATES}', candidateDateText);
    const subject = CONFIG.EMAIL_SUBJECT;
    const bccList = clientEmails.join(',');

    if (CONFIG.DRAFT_ONLY) {
      // 下書き保存（GmailApp.createDraft の bcc オプション）
      GmailApp.createDraft(
        CONFIG.FROM_EMAIL,
        subject,
        emailBody,
        { bcc: bccList }
      );
      const msg = `下書きを作成しました！\nGmailの下書きフォルダをご確認ください。\n\nBCC送信先: ${clientEmails.length}件`;
      Logger.log('✅ ' + msg);
      Browser.msgBox('下書き作成完了', msg, Browser.Buttons.OK);
    } else {
      // 即時送信
      GmailApp.sendEmail(
        CONFIG.FROM_EMAIL,
        subject,
        emailBody,
        { bcc: bccList }
      );
      const msg = `送信完了しました！\n\nBCC送信先: ${clientEmails.length}件`;
      Logger.log('✅ ' + msg);
      Browser.msgBox('送信完了', msg, Browser.Buttons.OK);
    }

    Logger.log('\n========================================');
    Logger.log('定例会 日程調整メール送信処理 完了');
    Logger.log('========================================');

  } catch (e) {
    Logger.log(`\n❌ エラーが発生しました:\n${e.message}\n${e.stack}`);
    Browser.msgBox('エラー', `処理中にエラーが発生しました:\n\n${e.message}`, Browser.Buttons.OK);
  }
}

// =====================================================
// LINE WORKS カレンダー連携
// =====================================================

/**
 * LINE WORKSから予定を取得し、空き時間を算出する
 */
function getFreeSlotsFromLineWorks() {
  const accessToken = getLineWorksAccessToken();
  const userId = encodeURIComponent(getProp('LW_USER_ID'));

  // 検索範囲: 今日から SEARCH_DAYS_AHEAD 日間
  const startSearch = new Date();
  const endSearch = new Date();
  endSearch.setDate(startSearch.getDate() + CONFIG.SEARCH_DAYS_AHEAD);

  const rangeStart = toIso8601Utc(startSearch);
  const rangeEnd = toIso8601Utc(endSearch);

  const url = `https://www.worksapis.com/v1.0/users/${userId}/calendar/events?rangeStartTime=${rangeStart}&rangeEndTime=${rangeEnd}&limit=100`;

  const response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: { 'Authorization': `Bearer ${accessToken}` },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`LINE WORKS APIエラー: ${response.getContentText()}`);
  }

  const events = JSON.parse(response.getContentText()).events || [];
  const freeSlots = [];

  // 1日ずつループして空き時間を計算
  for (let i = 0; i <= CONFIG.SEARCH_DAYS_AHEAD; i++) {
    const targetDate = new Date(startSearch.getTime());
    targetDate.setDate(targetDate.getDate() + i);

    // 土日はスキップ
    const dayOfWeek = targetDate.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) continue;

    // その日の予定をフィルタリング
    const dailyEvents = events.filter(ev => {
      const s = ev.start.dateTime || ev.start.date;
      const d = new Date(s);
      return d.getFullYear() === targetDate.getFullYear() &&
             d.getMonth() === targetDate.getMonth() &&
             d.getDate() === targetDate.getDate();
    });

    // 空き時間を計算
    const slots = calculateDailyFreeSlots(targetDate, dailyEvents);
    freeSlots.push(...slots);
  }

  return freeSlots;
}

/**
 * 1日の空き時間を計算する
 */
function calculateDailyFreeSlots(date, events) {
  const slots = [];
  
  // 判定用の基準時間を生成
  const createTime = (hour, min) => {
    const d = new Date(date.getTime());
    d.setHours(hour, min, 0, 0);
    return d;
  };

  // 10:00 - 13:00 と 14:00 - 18:00 の2つのブロックで判定
  const windows = [
    { start: createTime(CONFIG.WORK_START_HOUR, 0), end: createTime(CONFIG.LUNCH_START_HOUR, 0) },
    { start: createTime(CONFIG.LUNCH_END_HOUR, 0), end: createTime(CONFIG.WORK_END_HOUR, 0) }
  ];

  // 予定を [start, end] のミリ秒ペアに変換
  const busyPeriods = events.map(ev => {
    return {
      start: new Date(ev.start.dateTime || ev.start.date).getTime(),
      end: new Date(ev.end.dateTime || ev.end.date).getTime()
    };
  }).sort((a, b) => a.start - b.start);

  windows.forEach(window => {
    let currentStart = window.start.getTime();
    const windowEnd = window.end.getTime();

    busyPeriods.forEach(busy => {
      // 予定がウィンドウ外なら無視
      if (busy.end <= currentStart || busy.start >= windowEnd) return;

      // 前の予定との間に隙間があれば追加
      if (busy.start > currentStart) {
        const diffMin = (busy.start - currentStart) / (1000 * 60);
        if (diffMin >= CONFIG.MIN_SLOT_MINUTES) {
          slots.push({
            start: { dateTime: new Date(currentStart).toISOString() },
            end: { dateTime: new Date(busy.start).toISOString() }
          });
        }
      }
      // 次の開始地点を更新
      currentStart = Math.max(currentStart, busy.end);
    });

    // 最後の予定の後に隙間があれば追加
    if (currentStart < windowEnd) {
      const diffMin = (windowEnd - currentStart) / (1000 * 60);
      if (diffMin >= CONFIG.MIN_SLOT_MINUTES) {
        slots.push({
          start: { dateTime: new Date(currentStart).toISOString() },
          end: { dateTime: new Date(windowEnd).toISOString() }
        });
      }
    }
  });

  return slots;
}

/**
 * イベントリストを「・日付 時間~時間」形式に整形する
 * 同じ日付に複数イベントがある場合はまとめる
 *
 * 出力例:
 * ・4月6日(月) 16:00~17:00
 * ・4月7日(火) 14:00~15:00 17:00~18:00
 */
function formatCandidateDates(events) {
  const WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];

  // 各イベントを {dateKey, dateLabel, timeRange} に変換
  const parsed = events.map(ev => {
    const start = ev.start || {};
    const end = ev.end || {};
    const startStr = start.dateTime || start.date || '';
    const endStr = end.dateTime || end.date || '';

    const startDt = startStr ? new Date(startStr) : null;
    const endDt = endStr ? new Date(endStr) : null;

    if (!startDt) return null;

    // JSTに変換
    const jstStart = toJst(startDt);
    const jstEnd = endDt ? toJst(endDt) : null;

    const month = jstStart.getMonth() + 1;
    const day = jstStart.getDate();
    const wday = WEEKDAYS[jstStart.getDay()];

    // ソート・グルーピング用のキー（年月日）
    const dateKey = `${jstStart.getFullYear()}${String(month).padStart(2, '0')}${String(day).padStart(2, '0')}`;
    const dateLabel = `${month}月${day}日(${wday})`;

    // 時刻文字列
    const startTime = formatTime(jstStart);
    const endTime = jstEnd ? formatTime(jstEnd) : '';
    const timeRange = endTime ? `${startTime}~${endTime}` : startTime;

    return { dateKey, dateLabel, timeRange, sortTime: jstStart.getTime() };
  }).filter(Boolean);

  // 日付でソート
  parsed.sort((a, b) => a.sortTime - b.sortTime);

  // 日付ごとにグループ化
  const grouped = {};
  const dateOrder = [];

  parsed.forEach(item => {
    if (!grouped[item.dateKey]) {
      grouped[item.dateKey] = { dateLabel: item.dateLabel, times: [] };
      dateOrder.push(item.dateKey);
    }
    grouped[item.dateKey].times.push(item.timeRange);
  });

  // テキスト生成
  const lines = dateOrder.map(key => {
    const g = grouped[key];
    return `・${g.dateLabel} ${g.times.join(' ')}`;
  });

  return lines.join('\n');
}

// =====================================================
// LINE WORKS 認証（JWT + OAuth2）
// =====================================================

/**
 * LINE WORKS API 2.0 のアクセストークンを取得する
 */
function getLineWorksAccessToken() {
  const clientId = getProp('LW_CLIENT_ID');
  const clientSecret = getProp('LW_CLIENT_SECRET');
  const serviceAccount = getProp('LW_SERVICE_ACCOUNT');
  const privateKey = getProp('LW_PRIVATE_KEY');

  // 未設定チェック
  const missing = [];
  if (!clientId) missing.push('LW_CLIENT_ID');
  if (!clientSecret) missing.push('LW_CLIENT_SECRET');
  if (!serviceAccount) missing.push('LW_SERVICE_ACCOUNT');
  if (!privateKey) missing.push('LW_PRIVATE_KEY');

  if (missing.length > 0) {
    throw new Error(`スクリプトプロパティが未設定です:\n${missing.join(', ')}\n\nGASエディタ → プロジェクトの設定 → スクリプトプロパティ で登録してください。`);
  }

  // JWT 生成
  const jwt = createJwt(clientId, serviceAccount, privateKey);
  Logger.log('JWT生成完了');

  // アクセストークン取得
  const response = UrlFetchApp.fetch('https://auth.worksmobile.com/oauth2/v2.0/token', {
    method: 'POST',
    payload: {
      assertion: jwt,
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      client_id: clientId,
      client_secret: clientSecret,
      scope: 'calendar.read',
    },
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  if (status !== 200) {
    Logger.log(`トークン取得失敗: HTTP ${status}\n${response.getContentText()}`);
    throw new Error(`LINE WORKSトークン取得失敗 (HTTP ${status})\n${response.getContentText()}`);
  }

  const tokenData = JSON.parse(response.getContentText());
  Logger.log('✅ アクセストークン取得成功');
  return tokenData.access_token;
}

/**
 * RS256署名のJWTを作成する
 */
function createJwt(clientId, serviceAccount, privateKey) {
  const now = Math.floor(Date.now() / 1000);

  const header = { alg: 'RS256', typ: 'JWT' };
  const payload = {
    iss: clientId,
    sub: serviceAccount,
    iat: now,
    exp: now + 3600,
  };

  const encodeBase64 = (obj) =>
    Utilities.base64EncodeWebSafe(JSON.stringify(obj)).replace(/=+$/, '');

  const headerB64 = encodeBase64(header);
  const payloadB64 = encodeBase64(payload);
  const signingInput = `${headerB64}.${payloadB64}`;

  const sig = Utilities.computeRsaSha256Signature(signingInput, privateKey);
  const sigB64 = Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '');

  return `${signingInput}.${sigB64}`;
}

// =====================================================
// スプレッドシート連携
// =====================================================

/**
 * スプレッドシートのQ列（17列目）からメールアドレスを収集する
 */
function getClientEmailsFromSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // GIDでシートを特定
  const targetGid = parseInt(CONFIG.SHEET_GID, 10);
  const sheets = ss.getSheets();
  let sheet = null;

  for (const s of sheets) {
    if (s.getSheetId() === targetGid) {
      sheet = s;
      break;
    }
  }

  if (!sheet) {
    // GIDで見つからない場合はアクティブシートを使用
    Logger.log(`⚠️ GID ${CONFIG.SHEET_GID} のシートが見つかりません。最初のシートを使用します。`);
    sheet = ss.getSheets()[0];
  }

  Logger.log(`対象シート: 「${sheet.getName()}」`);

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    Logger.log('データが見つかりません');
    return [];
  }

  // Q列のデータを取得
  const range = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.EMAIL_COLUMN, lastRow - CONFIG.DATA_START_ROW + 1, 1);
  const values = range.getValues();

  const emails = [];
  values.forEach((row, i) => {
    const val = String(row[0] || '').trim();
    if (val && isValidEmail(val)) {
      emails.push(val);
    } else if (val) {
      Logger.log(`⚠️ 行${CONFIG.DATA_START_ROW + i}: 無効なメールアドレス → 「${val}」（スキップ）`);
    }
  });

  return emails;
}

// =====================================================
// ユーティリティ
// =====================================================

/** スクリプトプロパティを取得する */
function getProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

/** メールアドレスの簡易バリデーション */
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

/** DateオブジェクトをISO8601 UTC文字列に変換 */
function toIso8601Utc(date) {
  return Utilities.formatDate(date, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

/** DateオブジェクトをJST（UTC+9）に変換したDateを返す */
function toJst(date) {
  return new Date(date.getTime() + 9 * 60 * 60 * 1000);
}

/** JSTのDateから HH:MM 形式を返す */
function formatTime(jstDate) {
  const h = String(jstDate.getUTCHours()).padStart(2, '0');
  const m = String(jstDate.getUTCMinutes()).padStart(2, '0');
  return `${h}:${m}`;
}

// =====================================================
// テスト・デバッグ用関数
// =====================================================

/**
 * テスト① LINE WORKS API接続確認
 * スクリプトプロパティが正しく設定されているか確認できます
 */
function testLineWorksConnection() {
  try {
    Logger.log('LINE WORKS 接続テスト開始...');
    const token = getLineWorksAccessToken();
    const msg = `LINE WORKSへの接続に成功しました！\nトークン（先頭30文字）: ${token.substring(0, 30)}...`;
    Logger.log('✅ ' + msg);
    Browser.msgBox('接続成功', msg, Browser.Buttons.OK);
  } catch (e) {
    Logger.log('❌ ' + e.message);
    Browser.msgBox('接続失敗', `エラー:\n${e.message}`, Browser.Buttons.OK);
  }
}

/**
 * テスト② 空き時間の確認
 * メールを送信せずに、抽出された空き時間テキストだけを確認できます
 */
function testPreviewFreeSlots() {
  try {
    Logger.log('空き時間プレビュー開始...');
    const slots = getFreeSlotsFromLineWorks();

    if (slots.length === 0) {
      Browser.msgBox('空きなし', `指定期間内に空き時間が見つかりませんでした。`, Browser.Buttons.OK);
      return;
    }

    const text = formatCandidateDates(slots);
    Logger.log('\n--- 抽出された空き時間 ---\n' + text);

    const body = CONFIG.EMAIL_BODY.replace('{CANDIDATE_DATES}', text);
    Logger.log('\n--- メール本文プレビュー ---\n' + body);

    Browser.msgBox(
      'プレビュー',
      `抽出件数: ${slots.length}件\n\n【候補日時】\n${text}`,
      Browser.Buttons.OK
    );
  } catch (e) {
    Logger.log('❌ ' + e.message);
    Browser.msgBox('エラー', e.message, Browser.Buttons.OK);
  }
}

/**
 * テスト③ スプレッドシートのメールアドレス確認
 * Q列からどのメールアドレスが取得されるか確認できます
 */
function testReadEmails() {
  try {
    Logger.log('スプレッドシート読み込みテスト開始...');
    const emails = getClientEmailsFromSheet();
    const msg = `取得したメールアドレス: ${emails.length}件\n\n${emails.join('\n')}`;
    Logger.log(msg);
    Browser.msgBox(`メール取得結果（${emails.length}件）`, msg.substring(0, 500), Browser.Buttons.OK);
  } catch (e) {
    Logger.log('❌ ' + e.message);
    Browser.msgBox('エラー', e.message, Browser.Buttons.OK);
  }
}
