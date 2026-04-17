const SPREADSHEET_ID = '1XudqpFJEmNCsgPRHX9fHkT4mTjKJgCDXiDdwzJQU4tY';
const SHEET_A_NAME = '案件進行シート';
const SHEET_B_NAME = '全案件管理';

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput("No data");
    }
    
    let eventData = JSON.parse(e.postData.contents);
    let messageText = "";
    if (eventData.content && eventData.content.text) {
      messageText = eventData.content.text;
    } else if (eventData.text) {
      messageText = eventData.text;
    } else {
      return ContentService.createTextOutput("Success (skipped)");
    }
    
    if (!messageText.includes("法人名") || !messageText.includes("キーワード")) {
       return ContentService.createTextOutput("Success (Not a target message)");
    }
    
    let data = extractDataFromText(messageText);
    writeToSheetA(data);
    writeToSheetB(data);
    
    return ContentService.createTextOutput("Success");
  } catch (error) {
    console.error(error);
    return ContentService.createTextOutput("Error: " + error.message);
  }
}

function extractDataFromText(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(l => l !== '');
  
  const parseFlexibly = (labelPattern) => {
    // 古い形式：同じ行にある場合 (例: 法人名 : 株式会社XXX)
    const exactLineRegex = new RegExp(`^(?:\\d+\\.\\s*)?${labelPattern}[\\s:：]+(.*)$`, 'i');
    for (let i = 0; i < lines.length; i++) {
      let match = lines[i].match(exactLineRegex);
      if (match && match[1].trim() !== '') return match[1].trim();
    }
    
    // 新しい形式：改行されている場合 (例: 法人名 \n (説明書き) \n 株式会社XXX)
    const labelRegex = new RegExp(`^(?:\\d+\\.\\s*)?${labelPattern}`, 'i');
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].match(labelRegex)) {
        // ラベルの次の行から3行先まで確認して、最初に現れた「意味のある文字」を取得
        for(let j = i + 1; j < Math.min(i + 4, lines.length); j++) {
           let val = lines[j];
           // 注釈・説明書きと考えられる行はスキップ
           if (val.includes('※') || val.includes('同じであれば') || val.includes('ハイフン付') || 
               val.includes('カレンダー') || val.includes('記載') || val.includes('危険度低') || val.includes('有無の選択')) {
             continue;
           }
           // キーワードの次等で、もし別の項目名に被ってしまったら終了
           if (val.match(/^(?:法人住所|屋号名|店舗住所|代表者名|店舗電話|架電先|メール|危険度|運用開始|契約期間|投稿代行|口コミ|インスタ)/)) {
             break;
           }
           if (val !== '') return val;
        }
      }
    }
    return '';
  };

  // キーワード欄は少し特殊（長い改行を含むかもしれない）
  let keywords = '';
  // 「希望キーワード」から次の見出しまでの行を結合した後に分割
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].includes('希望キーワード') || lines[i] === 'キーワード') {
       let kwdParts = [];
       for(let j = i + 1; j < lines.length; j++) {
         if (lines[j].match(/^(?:法人|店舗|代表|危険度|運用|契約)/)) break; // 他の項目なら終了
         kwdParts.push(lines[j]);
       }
       keywords = kwdParts.join('、');
       break;
    }
  }
  if (!keywords) keywords = parseFlexibly('希望キーワード.*');

  return {
    corpName: parseFlexibly('法人名'),
    corpAddress: parseFlexibly('法人住所'),
    meoAddressCorp: parseFlexibly('MEO表示住所(?:[\\(（]法人[\\)）])?'),
    storeName: parseFlexibly('屋号名(?:[\\(（]店舗名[\\)）])?|屋号名|店舗名'),
    storeAddress: parseFlexibly('店舗住所'),
    meoAddressStore: parseFlexibly('MEO表示住所(?:[\\(（]店舗[\\)）])?'),
    repName: parseFlexibly('代表者名(?:[\\(（]カタカナ[\\)）])?|代表者名'),
    storePhone: parseFlexibly('店舗電話番号'),
    callPhone: parseFlexibly('架電先電話番号'),
    email: parseFlexibly('メールアドレス'),
    riskLevel: parseFlexibly('危険度.*'),
    startDate: parseFlexibly('運用開始日'),
    contractMonths: parseFlexibly('契約期間.*'),
    postCount: parseFlexibly('投稿代行.*'),
    otherStore: parseFlexibly('他店舗情報'),
    reviewReply: parseFlexibly('口コミ返信.*'),
    hearing: parseFlexibly('店舗情報ヒアリング'),
    instaLink: parseFlexibly('インスタ連動'),
    keywords: keywords
  };
}

// 日付を 2026/5/1 のように安全な形式でフォーマットする関数（入力規則エラー回避用）
function formatYMD_Safe(dateStr) {
  if (!dateStr) return "";
  let parsedStr = dateStr.trim().replace(/\s*\(.*?\)/g, '').replace(/[.\-]/g, '/').trim();
  let date = new Date(parsedStr);
  if (isNaN(date.getTime())) return dateStr;
  
  return date.getFullYear() + "/" + (date.getMonth() + 1) + "/" + date.getDate();
}

// 実際の末尾行を取得する関数（関数やスペースによる誤作動防止）
function getActualLastRow(sheet, colIndex) {
  const vals = sheet.getRange(1, colIndex, sheet.getMaxRows(), 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] && vals[i][0].toString().trim() !== "") {
      return i + 2; 
    }
  }
  return 2; // デフォルトは2行目から
}

function writeToSheetA(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_A_NAME);
  if (!sheet) throw new Error("シートが見つかりません: " + SHEET_A_NAME);
  
  let lastRow = getActualLastRow(sheet, 5); // E列（店舗名）を基準に末尾行を探す
  
  const updates = [];
  if (data.startDate) updates.push({col: 2, val: data.startDate}); // B
  if (data.corpName) updates.push({col: 4, val: data.corpName}); // D
  if (data.storeName) updates.push({col: 5, val: data.storeName}); // E
  
  let insta = data.instaLink ? data.instaLink.trim() : '';
  if (insta === '有' || insta === '有り') updates.push({col: 7, val: '有'}); // G
  else if (insta === '無' || insta === '無し') updates.push({col: 7, val: '無'});
  
  if (data.repName) updates.push({col: 12, val: data.repName}); // L
  if (data.email) updates.push({col: 13, val: data.email}); // M
  
  let reviewText = data.reviewReply.replace(/[^\d]/g, '');
  let reviewCount = parseInt(reviewText) || 0;
  updates.push({col: 20, val: reviewCount > 0 ? '有' : '無'}); // T
  
  updates.push({col: 23, val: '毎月'}); // W
  
  let postCountVal = parseInt(data.postCount.replace(/[^\d]/g, ''));
  if (!isNaN(postCountVal)) updates.push({col: 24, val: postCountVal}); // X
  
  updates.forEach(u => {
    sheet.getRange(lastRow, u.col).setValue(u.val);
  });
}

function writeToSheetB(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_B_NAME);
  if (!sheet) throw new Error("シートが見つかりません: " + SHEET_B_NAME);
  
  let lastRow = getActualLastRow(sheet, 3); // C列（店舗名）を基準に末尾行を探す
  
  const updates = [];
  if (data.corpName) updates.push({col: 2, val: data.corpName}); // B
  if (data.storeName) updates.push({col: 3, val: data.storeName}); // C
  if (data.repName) updates.push({col: 4, val: data.repName}); // D
  if (data.callPhone) updates.push({col: 5, val: data.callPhone}); // E
  if (data.storePhone) updates.push({col: 6, val: data.storePhone}); // F
  if (data.email) updates.push({col: 7, val: data.email}); // G
  
  let riskMatch = data.riskLevel.match(/\d+/);
  let riskStars = "";
  if (riskMatch) {
    let num = parseInt(riskMatch[0], 10);
    if (num >= 1 && num <= 5) riskStars = "★".repeat(num);
  } else {
    let starMatch = data.riskLevel.match(/★/g);
    if (starMatch) riskStars = "★".repeat(starMatch.length);
  }
  if (riskStars || data.riskLevel) updates.push({col: 9, val: riskStars ? riskStars : data.riskLevel}); // I
  
  if (data.startDate) updates.push({col: 13, val: formatYMD_Safe(data.startDate)}); // M
  
  if (data.startDate && data.contractMonths) {
    let months = parseInt(data.contractMonths.replace(/[^\d]/g, ''), 10);
    if (!isNaN(months)) {
      let parsedStr = data.startDate.trim().replace(/\s*\(.*?\)/g, '').replace(/[.\-]/g, '/').trim();
      let date = new Date(parsedStr);
      if (!isNaN(date.getTime())) {
        date.setMonth(date.getMonth() + months);
        date.setDate(date.getDate() - 1);
        updates.push({col: 14, val: date.getFullYear() + "/" + (date.getMonth() + 1) + "/" + date.getDate()}); // N
      }
    }
  }
  
  if (data.storeAddress) updates.push({col: 15, val: data.storeAddress}); // O
  let meoAddress = (data.meoAddressStore && data.meoAddressStore.trim() !== '') ? data.meoAddressStore : data.storeAddress;
  if (meoAddress) updates.push({col: 16, val: meoAddress}); // P
  
  let postCountVal = data.postCount.replace(/[^\d]/g, '');
  if (postCountVal) updates.push({col: 17, val: `毎月${postCountVal}本`}); // Q
  
  let reviewNumStr = data.reviewReply.match(/\d+/);
  if (reviewNumStr) updates.push({col: 18, val: `${reviewNumStr[0]}回`}); // R
  
  let insta = data.instaLink ? data.instaLink.trim() : '';
  if (insta === '有' || insta === '有り') updates.push({col: 21, val: '未実行'}); // U
  else if (insta === '無' || insta === '無し') updates.push({col: 21, val: '無し'}); // U
  else if (insta) updates.push({col: 21, val: insta});
  
  const keywords = data.keywords.split(/[,、\s]+/).filter(k => k.trim() !== '');
  for (let i = 0; i < 7; i++) {
    if (keywords[i]) updates.push({col: 22 + i, val: keywords[i]}); // V(22)〜
  }
  
  updates.forEach(u => {
    sheet.getRange(lastRow, u.col).setValue(u.val);
  });
}

// 【テスト実行用】
function testInsertExtractedData() {
  const extractedData = {
    corpName: "",
    storeName: "おそうじ本舗 高知旭店",
    repName: "ヤマナカ ヒロマサ",
    callPhone: "090-9777-8614",
    storePhone: "090-9777-8614",
    email: "hiromax3211@gmail.com",
    riskLevel: "★★",
    startDate: "2026.05.01 (金)",
    contractMonths: "12カ月",
    postCount: "月に4回",
    reviewReply: "月に1回",
    instaLink: "無",
    storeAddress: "高知県吾川郡いの町枝川８５８－１２",
    meoAddressStore: "",
    keywords: "エアコンクリーニング、ハウスクリーニング、お風呂クリーニング、エアコン掃除業者"
  };
  
  writeToSheetA(extractedData);
  writeToSheetB(extractedData);
}
