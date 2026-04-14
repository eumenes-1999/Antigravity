
// ============================================================
//  月初 BCC 一斉メール送信スクリプト
//  ・LINE Works カレンダーから空き枠を取得
//  ・顧客管理シートのR列=TRUEの宛先にBCC送信
//  ・送信前に確認ダイアログを表示
// ============================================================

// ─── 設定エリア ──────────────────────────────────────────────
const CONFIG = {
  // 顧客管理スプレッドシートURL
  SPREADSHEET_URL: 'https://docs.google.com/spreadsheets/d/1Qrd2HTrWSeYp2f5vrMg3rV7ey72GS2nN463LdzgrrQo/edit?gid=1447447650#gid=1447447650',
  SHEET_GID: '1447447650',          // 対象シートのGID

  // 送信元メール
  FROM_EMAIL: 'creative.yt@runbird.net',
  FROM_NAME: '株式会社ランバード 広告運用サポート窓口',

  // 列インデックス（0始まり）
  COL_EMAIL: 16,   // Q列 = 16
  COL_SEND_FLAG: 17, // R列 = 17

  // LINE Works API設定
  LW_USER_ID: 's.abe@runbird-biz',
  LW_ACCESS_TOKEN_KEY: 'LW_ACCESS_TOKEN', // PropertiesServiceに保存したキー名

  // MTG候補を何個ピックアップするか
  SLOT_COUNT: 10,

  // MTGの所要時間（分）
  MTG_DURATION_MIN: 30,

  // 1日の候補時間帯（24h表記）
  AVAILABLE_HOURS: [10, 11, 13, 14, 15, 16, 17, 18],

  // 何か月先まで検索するか
  SEARCH_WEEKS: 5,
};

// ─── メインエントリポイント ────────────────────────────────────

/**
 * この関数をトリガーで月初に実行する。
 * 手動実行も可能。
 */
function sendMonthlyBccMail() {
  // 1. LINE Works アクセストークンを取得
  const accessToken = getLineWorksAccessToken();
  if (!accessToken) {
    SpreadsheetApp.getUi().alert('❌ LINE Works のアクセストークンが取得できませんでした。\nスクリプトプロパティ「LW_ACCESS_TOKEN」を確認してください。');
    return;
  }

  // 2. 空き日程を取得
  const freeSlots = getFreeSlots_(accessToken);
  if (freeSlots.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ 空き日程が見つかりませんでした。LINE Worksカレンダーを確認してください。');
    return;
  }

  // 3. BCC宛先を収集
  const bccList = getBccList_();
  if (bccList.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ 送信対象のクライアントが見つかりません。スプレッドシートのR列を確認してください。');
    return;
  }

  // 4. メール本文を組み立て
  const { subject, body } = buildEmailContent_(freeSlots);

  // 5. 確認ダイアログ（プレビュー付き）を表示
  const confirmed = showConfirmDialog_(bccList, subject, body, freeSlots);
  if (!confirmed) {
    Logger.log('送信をキャンセルしました。');
    return;
  }

  // 6. BCC 送信
  sendBccMail_(bccList, subject, body);

  SpreadsheetApp.getUi().alert(`✅ 送信完了！\n${bccList.length} 件のアドレスにBCCで送信しました。`);
}

// ─── LINE Works: アクセストークン取得 ──────────────────────────

function getLineWorksAccessToken() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(CONFIG.LW_ACCESS_TOKEN_KEY);
  if (!token) {
    Logger.log('スクリプトプロパティにアクセストークンが設定されていません: ' + CONFIG.LW_ACCESS_TOKEN_KEY);
    return null;
  }
  return token;
}

// ─── LINE Works: 空き日程の取得 ───────────────────────────────

function getFreeSlots_(accessToken) {
  const now = new Date();
  const endDate = new Date();
  endDate.setDate(now.getDate() + CONFIG.SEARCH_WEEKS * 7);

  // ISO8601 形式で期間を指定
  const fromStr = Utilities.formatDate(now, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss'+09:00'");
  const toStr   = Utilities.formatDate(endDate, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss'+09:00'");

  // LINE Works Calendar API v2: 個人カレンダーのイベント一覧取得
  // GET /users/{userId}/calendar-personals
  const url = `https://www.worksapis.com/v1.0/users/${encodeURIComponent(CONFIG.LW_USER_ID)}/calendar-personals?fromDateTime=${encodeURIComponent(fromStr)}&toDateTime=${encodeURIComponent(toStr)}&limit=500`;

  let existingEvents = [];
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + accessToken,
        'Content-Type': 'application/json',
      },
      muteHttpExceptions: true,
    });

    const code = response.getResponseCode();
    const text = response.getContentText();
    Logger.log('LINE Works API status: ' + code);
    Logger.log('LINE Works API response: ' + text.substring(0, 500));

    if (code === 200) {
      const json = JSON.parse(text);
      // イベントの時間帯を抽出
      const events = json.calendarPersonals || json.calendars || json.events || [];
      for (const ev of events) {
        const start = ev.startDateTime || ev.start;
        const end   = ev.endDateTime   || ev.end;
        if (start && end) {
          existingEvents.push({
            start: new Date(start),
            end:   new Date(end),
          });
        }
      }
    } else {
      // APIエラー時はGoogleカレンダーをフォールバックとして使用
      Logger.log('LINE Works APIエラー。Googleカレンダーをフォールバックとして使用します。');
      existingEvents = getGoogleCalendarEvents_(now, endDate);
    }
  } catch (e) {
    Logger.log('LINE Works API 例外: ' + e.message);
    existingEvents = getGoogleCalendarEvents_(now, endDate);
  }

  // 空き枠を探す
  return findFreeSlots_(now, endDate, existingEvents);
}

/** Googleカレンダーからイベントを取得（フォールバック） */
function getGoogleCalendarEvents_(from, to) {
  const cal = CalendarApp.getDefaultCalendar();
  const events = cal.getEvents(from, to);
  return events.map(ev => ({ start: ev.getStartTime(), end: ev.getEndTime() }));
}

/** 空き枠を探してリスト返却 */
function findFreeSlots_(from, to, busyEvents) {
  const slots = [];
  const dur = CONFIG.MTG_DURATION_MIN * 60 * 1000;
  const today = new Date(from);
  today.setHours(0, 0, 0, 0);

  // 翌稼働日から開始（当日は除外）
  const startDate = new Date(today);
  startDate.setDate(startDate.getDate() + 1);

  const current = new Date(startDate);

  while (current <= to && slots.length < CONFIG.SLOT_COUNT) {
    const dow = current.getDay();
    // 土日はスキップ
    if (dow === 0 || dow === 6) {
      current.setDate(current.getDate() + 1);
      continue;
    }

    for (const hour of CONFIG.AVAILABLE_HOURS) {
      if (slots.length >= CONFIG.SLOT_COUNT) break;

      const slotStart = new Date(current);
      slotStart.setHours(hour, 0, 0, 0);
      const slotEnd = new Date(slotStart.getTime() + dur);

      if (!isConflict_(slotStart, slotEnd, busyEvents)) {
        slots.push(slotStart);
      }
    }
    current.setDate(current.getDate() + 1);
  }

  return slots;
}

function isConflict_(start, end, busyEvents) {
  for (const ev of busyEvents) {
    if (start < ev.end && end > ev.start) return true;
  }
  return false;
}

// ─── スプレッドシート: BCC宛先取得 ────────────────────────────

function getBccList_() {
  const ss = SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL);
  // GIDでシートを特定
  const sheets = ss.getSheets();
  let sheet = null;
  for (const s of sheets) {
    if (String(s.getSheetId()) === CONFIG.SHEET_GID) {
      sheet = s;
      break;
    }
  }
  if (!sheet) sheet = ss.getActiveSheet();

  const data = sheet.getDataRange().getValues();
  const bccEmails = [];

  for (let i = 1; i < data.length; i++) {
    const sendFlag = data[i][CONFIG.COL_SEND_FLAG]; // R列
    const email    = data[i][CONFIG.COL_EMAIL];      // Q列

    // R列がTRUE（またはチェックボックスで true）かつメールアドレスが存在する場合
    if (sendFlag === true && email && String(email).includes('@')) {
      bccEmails.push(String(email).trim());
    }
  }

  // 重複除去
  return [...new Set(bccEmails)];
}

// ─── メール本文の組み立て ──────────────────────────────────────

/**
 * 空きスロット配列（Date[]）を受け取り、
 * 日付ごとに時間帯をまとめた候補日文字列を生成する。
 * 例）・4月7日(火) 14:00~14:30 17:00~17:30
 */
function formatSlotsPerDay_(freeSlots) {
  const WEEKDAY_JA = ['日', '月', '火', '水', '木', '金', '土'];
  const durMin = CONFIG.MTG_DURATION_MIN;

  // 日付キーでグルーピング（キー: 'yyyy/MM/dd'）
  const dayMap = {};
  const dayOrder = [];
  for (const slot of freeSlots) {
    const key = Utilities.formatDate(slot, 'Asia/Tokyo', 'yyyy/MM/dd');
    if (!dayMap[key]) {
      dayMap[key] = { slot: slot, times: [] };
      dayOrder.push(key);
    }
    const hStart = String(slot.getHours()).padStart(2, '0');
    const mStart = String(slot.getMinutes()).padStart(2, '0');
    const endSlot = new Date(slot.getTime() + durMin * 60 * 1000);
    const hEnd = String(endSlot.getHours()).padStart(2, '0');
    const mEnd = String(endSlot.getMinutes()).padStart(2, '0');
    dayMap[key].times.push(`${hStart}:${mStart}~${hEnd}:${mEnd}`);
  }

  return dayOrder.map(key => {
    const info = dayMap[key];
    const s = info.slot;
    const m = s.getMonth() + 1;
    const d = s.getDate();
    const w = WEEKDAY_JA[s.getDay()];
    const timesStr = info.times.join(' ');
    return `・${m}月${d}日(${w}) ${timesStr}`;
  }).join('\n');
}

function buildEmailContent_(freeSlots) {
  const now = new Date();
  const month = now.getMonth() + 1;

  const subject = `月次報告 日程調整のご連絡`;

  // 日付ごとにまとめた候補日文字列
  const slotLines = formatSlotsPerDay_(freeSlots);

  const body =
`お客様各位

平素より大変お世話になっております。
株式会社ランバード　広告運用サポート窓口です。

本メールは皆様へのお知らせのためBCCで送信しております。

月次の報告をさせていただきたく思い、日程調整のご連絡をいたしました。
下記日程より、3つほどご都合の良いお日にちを挙げて頂ければ幸いでございます。

【候補日時】
${slotLines}
※所要時間は30分ほどを予定しております。

ご確認のほどよろしくお願いいたします。
＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
株式会社ランバード
広告運用サポート窓口
東京都中央区晴海1丁目8−10
晴海アイランド トリトンスクエア オフィスタワーX棟40階
TEL　0120-949-851　FAX　03-5843-6978
Mail   creative.yt@runbird.net	
HP　 https://runbird.net/
＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝`;

  return { subject, body };
}

// ─── 確認ダイアログ ────────────────────────────────────────────

function showConfirmDialog_(bccList, subject, body, freeSlots) {
  const slotPreview = formatSlotsPerDay_(freeSlots);

  const confirmMsg = [
    '─ 送信内容の確認 ─────────────────────',
    '',
    `▶ 件名：${subject}`,
    `▶ 送信元：${CONFIG.FROM_EMAIL}`,
    `▶ BCC送信先：${bccList.length} 件`,
    '',
    '【候補日時（メール本文に挿入される内容）】',
    slotPreview,
    '',
    '─────────────────────────────────────',
    '「OK」を押すと送信を実行します。',
    '「キャンセル」で中止します。',
  ].join('\n');

  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('📧 月初BCC一斉メール 送信確認', confirmMsg, ui.ButtonSet.OK_CANCEL);
  return result === ui.Button.OK;
}

// ─── BCC 送信 ──────────────────────────────────────────────────

function sendBccMail_(bccList, subject, body) {
  const bccStr = bccList.join(',');

  GmailApp.sendEmail(
    CONFIG.FROM_EMAIL,  // To: 自分自身（BCCなので実質空でもよいが必須フィールドのため）
    subject,
    body,
    {
      from: CONFIG.FROM_EMAIL,
      name: CONFIG.FROM_NAME,
      bcc: bccStr,
    }
  );

  Logger.log(`BCC送信完了: ${bccList.length}件 → ${bccStr}`);
}

// ─── テスト用: 空き枠だけ確認したい場合 ─────────────────────────

function testGetFreeSlots() {
  const now = new Date();
  const endDate = new Date();
  endDate.setDate(now.getDate() + CONFIG.SEARCH_WEEKS * 7);
  const token = getLineWorksAccessToken();
  const slots = token
    ? getFreeSlots_(token)
    : findFreeSlots_(now, endDate, getGoogleCalendarEvents_(now, endDate));
  const result = formatSlotsPerDay_(slots);
  Logger.log('空き枠一覧 (' + slots.length + '件):\n' + result);
  SpreadsheetApp.getUi().alert('空き枠一覧 (' + slots.length + '件)\n\n' + result);
}

/** BCC宛先リストだけ確認 */
function testGetBccList() {
  const list = getBccList_();
  Logger.log('BCC宛先 ' + list.length + '件: ' + list.join(', '));
  SpreadsheetApp.getUi().alert('BCC宛先 ' + list.length + '件\n\n' + list.join('\n'));
}
