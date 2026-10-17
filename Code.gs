/**
 * 小學生練習營 — 成績雲端存檔（貼到 Google 試算表的 Apps Script）
 *
 * 1. 把下面的 FAMILY_KEY 改成你自己的家庭密碼（平板設定時要輸入一樣的）。
 * 2. 部署 → 新增部署作業 → 類型選「網頁應用程式」
 *    執行身分：我；誰可以存取：所有人。
 * 3. 把產生的網址（…/exec）貼到平板上的「家長設定」。
 *
 * 試算表會自動建立兩個工作表：
 *   「營員」：每位小朋友一列（刪掉那一列 = 刪除這位小朋友的所有學習資料）
 *   「作答紀錄」：每答一題一列，可以自己篩選、做樞紐分析
 */
const FAMILY_KEY = '請改成你的家庭密碼';

const KIDS_SHEET = '營員';
const LOG_SHEET = '作答紀錄';
const KIDS_HEADER = ['營員ID', '名字', '年級學期', '金幣', '經驗值', '勳章數', '練習天數', '最後更新', '資料（請勿修改）'];
const LOG_HEADER = ['時間', '名字', '年級學期', '科目', '課次', '單元', '難度', '題目', '選的答案', '正確答案', '對錯', '營員ID'];

function doGet(e) {
  const p = e.parameter || {};
  if (p.key !== FAMILY_KEY) return json_({ ok: false, error: '家庭密碼不對' });
  if (p.action === 'load') return json_({ ok: true, kids: readKids_() });
  return json_({ ok: true, message: '練習營雲端存檔運作中' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '資料格式錯誤' }); }
  if (body.key !== FAMILY_KEY) return json_({ ok: false, error: '家庭密碼不對' });
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (body.action === 'saveKid') saveKid_(body.kid);
    else if (body.action === 'log') appendLogs_(body.rows || []);
    else return json_({ ok: false, error: '不認得的動作' });
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sh;
}

function readKids_() {
  const sh = sheet_(KIDS_SHEET, KIDS_HEADER);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(2, 9, last - 1, 1).getValues();
  const kids = [];
  data.forEach(function (r) {
    if (!r[0]) return;
    try { kids.push(JSON.parse(r[0])); } catch (err) {}
  });
  return kids;
}

function saveKid_(kid) {
  if (!kid || !kid.id) return;
  const sh = sheet_(KIDS_SHEET, KIDS_HEADER);
  const term = ['', '一', '二', '三', '四', '五', '六'][kid.grade] + (kid.sem || '');
  const row = [kid.id, kid.name, term, kid.coins, kid.xp, (kid.medals || []).length, (kid.days || []).length,
    new Date(), JSON.stringify(kid)];
  const last = sh.getLastRow();
  if (last >= 2) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === kid.id) { sh.getRange(i + 2, 1, 1, row.length).setValues([row]); return; }
    }
  }
  sh.appendRow(row);
}

function appendLogs_(rows) {
  if (!rows.length) return;
  const sh = sheet_(LOG_SHEET, LOG_HEADER);
  const values = rows.map(function (r) {
    return [new Date(r.time), r.kid, r.term, r.subject, r.lesson, r.unit, r.level, r.question, r.chosen, r.answer, r.ok ? '對' : '錯', r.kidId];
  });
  sh.getRange(sh.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
