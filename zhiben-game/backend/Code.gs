/**
 * 脂寶烤烤樂｜排行榜後端（Google Apps Script）
 * ------------------------------------------------------------
 * 部署步驟：
 * 1. 建一個 Google 試算表，命名「脂寶烤烤樂_排行榜」。
 * 2. 擴充功能 → Apps Script，貼上本檔全部內容，存檔。
 * 3. （不用手動執行）第一次被呼叫時會自動建立工作表、核銷碼與每週發獎排程。
 * 4. 部署 → 新增部署作業 → 類型「網頁應用程式」
 *    執行身分：我；存取權：所有人 → 取得網址。
 * 5. 把網址貼到遊戲 HTML 的 CFG.API_URL。
 *
 * 遊戲端呼叫：
 *  POST  body = {"action":"submit","name":"小米","store":"ch","score":2350,"level":4}
 *  GET   ?action=board&scope=ch     （scope：yl 員林／ch 彰化／bt 北屯／all 總榜）
 *  POST  body = {"action":"redeem","voucher":"T-2026-11-02-ch-小米","code":"ZB-UDG-ARF","type":"treat","name":"小米","store":"ch","pin":"2202"}
 *
 * 核銷碼：在 Apps Script「專案設定 → 指令碼屬性」新增 PINS，值為
 *   {"yl":"四位數","ch":"四位數","bt":"四位數"}
 * 核銷碼只存在後端，網頁原始碼看不到；建議每月更換一次。
 * 週結算一律以伺服器的台灣時間為準（每週一 00:00 重置），不採用手機時間。
 */

const SHEET = 'scores';
const STORES = { yl: '員林店', ch: '彰化店', bt: '北屯店' };
const MAX_PER_LEVEL = 3500;   // 防作弊：每關分數上限（依實際數據調整）
const BOARD_SIZE = 50;
// 三店核銷碼預設值（只存在後端；之後可在「專案設定 → 指令碼屬性 → PINS」修改）
const DEFAULT_PINS = null;   // 公開版不放核銷碼；請在「專案設定 → 指令碼屬性」新增 PINS

/** 第一次被呼叫時自動建立工作表、核銷碼與每週發獎排程，不需要手動執行 setup */
function ensure_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('READY') === '2') return;
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.getRange(1, 1, 1, 7).setValues([['時間', '週', '分店', '名字', '分數', '關卡', '可疑']]); sh.setFrozenRows(1); }
  let rd = ss.getSheetByName('核銷紀錄');
  if (!rd) { rd = ss.insertSheet('核銷紀錄'); rd.getRange(1, 1, 1, 7).setValues([['核銷時間', '分店', '票券類型', '票券ID', '票券編號', '名字', '備註']]); rd.setFrozenRows(1); }
  let wn = ss.getSheetByName('得獎名單');
  if (!wn) { wn = ss.insertSheet('得獎名單'); wn.appendRow(['週', '榜別', '名次', '名字', '分店', '分數', '獎品', '已發放']); wn.setFrozenRows(1); }
  const s1 = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  if (!props.getProperty('PINS') && DEFAULT_PINS) props.setProperty('PINS', JSON.stringify(DEFAULT_PINS));
  try {
    if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'weeklyWinners')) {
      ScriptApp.newTrigger('weeklyWinners').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(0).nearMinute(5).inTimezone('Asia/Taipei').create();
    }
  } catch (err) { /* 排程建立失敗不影響排行榜，可之後手動加 */ }
  props.setProperty('READY', '2');
}

function setup() { ensure_(); }

/** 台灣時間的本週一日期，例如 2026-11-02 */
function weekKey_(d) {
  const tw = new Date((d || new Date()).getTime() + 8 * 3600 * 1000);
  const dow = (tw.getUTCDay() + 6) % 7;
  const mon = new Date(Date.UTC(tw.getUTCFullYear(), tw.getUTCMonth(), tw.getUTCDate() - dow));
  return Utilities.formatDate(mon, 'UTC', 'yyyy-MM-dd');
}

function cleanName_(n) {
  return String(n || '').replace(/[<>"'&]/g, '').replace(/\s+/g, ' ').trim().slice(0, 10);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body = {};
  ensure_();
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  if (body.action === 'redeem') return redeem_(body);
  if (body.action !== 'submit') return json_({ ok: false, error: 'bad_action' });

  const name = cleanName_(body.name);
  const store = String(body.store || '');
  const score = Math.floor(Number(body.score) || 0);
  const level = Math.floor(Number(body.level) || 1);
  if (!name || !STORES[store] || score < 0 || level < 1 || level > 60) return json_({ ok: false, error: 'invalid' });

  const suspicious = score > level * MAX_PER_LEVEL;   // 超出人類可能範圍 → 標記、不上榜
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    SpreadsheetApp.getActive().getSheetByName(SHEET)
      .appendRow([new Date(), "'" + weekKey_(), store, name, score, level, suspicious ? 'Y' : '']);   // 加 ' 讓週別保持文字，不被轉成日期
  } finally { lock.releaseLock(); }
  const wkNow = weekKey_();
  CacheService.getScriptCache().removeAll(['yl', 'ch', 'bt', 'all'].map(x => 'board_' + x + '_' + wkNow));
  return json_({ ok: true, flagged: suspicious });
}

function doGet(e) {
  ensure_();
  const p = e.parameter || {};
  if (p.action === 'ping') return json_({ ok: true, week: weekKey_() });
  if (p.action !== 'board') return json_({ ok: false, error: 'bad_action' });
  const scope = STORES[p.scope] ? p.scope : 'all';
  // 可查本週或上週（發獎用）；格式不對就用本週
  const wk = /^\d{4}-\d{2}-\d{2}$/.test(p.week || '') ? p.week : weekKey_();
  const cache = CacheService.getScriptCache();
  const ckey = 'board_' + scope + '_' + wk;
  const hit = cache.get(ckey);
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);

  const rows = SpreadsheetApp.getActive().getSheetByName(SHEET).getDataRange().getValues().slice(1);
  const best = {};
  rows.forEach(r => {
    const [, week, store, name, score, level, flag] = r;
    const wkStr = week instanceof Date ? Utilities.formatDate(week, 'Asia/Taipei', 'yyyy-MM-dd') : String(week);
    if (wkStr !== wk || flag === 'Y') return;
    if (scope !== 'all' && store !== scope) return;
    const k = name + '|' + store;
    if (!best[k] || score > best[k].score) best[k] = { name: name, store: store, score: Number(score), level: Number(level) };
  });
  const out = JSON.stringify(Object.values(best).sort((a, b) => b.score - a.score).slice(0, BOARD_SIZE));
  cache.put(ckey, out, wk === weekKey_() ? 20 : 600);   // 本週快取 20 秒，上週 10 分鐘
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 每週一 00:05 自動執行（觸發條件 → 時間驅動 → 每週）：
 * 把上週各店前 3 名與總榜第 1 名寫到「得獎名單」工作表，方便店長發獎。
 */
function weeklyWinners() {
  const ss = SpreadsheetApp.getActive();
  const lastWeek = weekKey_(new Date(Date.now() - 7 * 864e5));
  const rows = ss.getSheetByName(SHEET).getDataRange().getValues().slice(1)
    .filter(r => String(r[1]) === lastWeek && r[6] !== 'Y');
  const best = {};
  rows.forEach(([, , store, name, score, level]) => {
    const k = name + '|' + store;
    if (!best[k] || score > best[k].score) best[k] = { name, store, score, level };
  });
  const all = Object.values(best).sort((a, b) => b.score - a.score);
  let sh = ss.getSheetByName('得獎名單') || ss.insertSheet('得獎名單');
  if (sh.getLastRow() === 0) sh.appendRow(['週', '榜別', '名次', '名字', '分店', '分數', '獎品', '已發放']);
  Object.keys(STORES).forEach(s => {
    all.filter(x => x.store === s).slice(0, 3).forEach((x, i) =>
      sh.appendRow(["'" + lastWeek, STORES[s], i + 1, x.name, STORES[x.store], x.score, i === 0 ? 'A5和牛霜降一份' : '手切松阪豬一份', '']));
  });
  if (all[0]) sh.appendRow(["'" + lastWeek, '總榜', 1, all[0].name, STORES[all[0].store], all[0].score, '脂本烤神稱號＋月見和牛飯', '']);
}


/**
 * 核銷：檢查核銷碼 → 檢查是否已核銷 → 週排行獎檢查得獎名單 → 寫入核銷紀錄
 * 回傳 {ok:true} 或 {ok:false, error:'bad_pin'|'locked'|'used'|'not_winner'|'invalid'}
 */
function redeem_(b) {
  const store = String(b.store || '');
  const id = String(b.voucher || '');
  const name = cleanName_(b.name);
  if (!STORES[store] || !id || !name) return json_({ ok: false, error: 'invalid' });

  // 防暴力猜碼：同一分店 10 分鐘內錯 10 次就鎖 10 分鐘
  const cache = CacheService.getScriptCache();
  const failKey = 'pinfail_' + store;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= 10) return json_({ ok: false, error: 'locked' });

  const pins = JSON.parse(PropertiesService.getScriptProperties().getProperty('PINS') || '{}');
  if (String(b.pin) !== String(pins[store] || '')) {
    cache.put(failKey, String(fails + 1), 600);
    return json_({ ok: false, error: 'bad_pin' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const ss = SpreadsheetApp.getActive();
    const rd = ss.getSheetByName('核銷紀錄');
    const used = rd.getDataRange().getValues().slice(1).some(r => String(r[3]) === id);
    if (used) return json_({ ok: false, error: 'used' });

    // 週排行獎：必須出現在「得獎名單」上
    if (b.type === 'weekly' || b.type === 'grand') {
      const parts = id.split('-');            // W-2026-10-19-ch-1-名字 ／ G-2026-10-19-名字
      const week = parts.slice(1, 4).join('-');
      const wins = (ss.getSheetByName('得獎名單') || { getDataRange: () => ({ getValues: () => [[]] }) })
        .getDataRange().getValues().slice(1);
      const ok = wins.some(w => String(w[0]) === week && String(w[3]) === name &&
        (b.type === 'grand' ? w[1] === '總榜' : (w[1] === STORES[store] && String(w[2]) === parts[5])));
      if (!ok) return json_({ ok: false, error: 'not_winner' });
    }
    rd.appendRow([new Date(), STORES[store], b.type, id, String(b.code || ''), name, '']);
  } finally { lock.releaseLock(); }
  return json_({ ok: true });
}
