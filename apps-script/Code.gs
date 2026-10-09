/**
 * 倉庫盤點掃描 — Google Apps Script 後端（v2）
 *
 * 部署方式：擴充功能 → Apps Script → 貼上本檔 → 部署 → 管理部署作業 → 編輯（鉛筆）
 *          → 版本選「新版本」→ 部署。這樣網址不變，前端不用改。
 *          執行身分：我；存取權：任何人。
 *
 * 存取碼（強烈建議設定）：專案設定 → 指令碼屬性 → 新增 API_TOKEN = 自訂一組密碼。
 *          設定後，手機要在網頁的「⚙ 設定」輸入同一組存取碼才能讀寫。
 *
 * 欄位是依「標題列文字」自動尋找的，欄位順序不限。若你的標題名稱不同，請改下方 CONFIG.COLS。
 */

const CONFIG = {
  SHEET_NAME: '',          // 庫存工作表名稱；空白 = 第一個工作表
  HEADER_ROW: 1,           // 標題列在第幾列
  COLS: {                  // 每個欄位可接受的標題名稱（找到第一個符合的就用）
    name:        ['品名', '商品名稱', '名稱', '品項'],
    model:       ['型號', '規格', '料號'],
    barcode:     ['國際條碼', '條碼', 'Barcode', 'EAN'],
    book:        ['帳面數量', '帳面庫存', '系統數量'],   // 選用：盤點差異報表用
    location:    ['儲位', '位置', '櫃位'],               // 選用：沒有會在第一次設定儲位時自動新增
    lastCounted: ['最後盤點時間'],                       // 自動新增
    counter:     ['盤點人員'],                           // 自動新增
  },
  STATUSES: ['新品', '福利品', '瑕疵', '報廢', '樣品', '採購樣品'],  // 每個狀態一欄，標題需完全相同
  LOG_SHEET: '異動紀錄',   // 每次寫入都會記一筆到這個工作表（自動建立）
};
const VERSION = 2;

// ===================== 進入點 =====================
function doGet(e)  { return respond_(() => handle_(e.parameter || {})); }
function doPost(e) {
  return respond_(() => {
    let p = {};
    try { p = JSON.parse((e.postData && e.postData.contents) || '{}'); } catch (err) { throw new Error('資料格式錯誤'); }
    return handle_(p);
  });
}
function respond_(fn) {
  let out;
  try { out = fn(); } catch (err) { out = { error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(p) {
  checkToken_(p.token);
  switch (p.action) {
    case 'ping':        return { ok: true, version: VERSION, tokenRequired: !!getToken_() };
    case 'lookup':      return lookup_(p);
    case 'search':      return search_(p);
    case 'list':        return list_();
    case 'save':        return withLock_(() => save_(p));
    case 'create':      return withLock_(() => create_(p));
    case 'bindBarcode': return withLock_(() => bindBarcode_(p));
    case 'setLocation': return withLock_(() => setLocation_(p));
    case 'startRound':  return withLock_(() => startRound_(p));
    default: throw new Error('未知的動作：' + p.action);
  }
}

// ===================== 驗證 / 鎖 =====================
function getToken_() { return PropertiesService.getScriptProperties().getProperty('API_TOKEN') || ''; }
function checkToken_(t) {
  const need = getToken_();
  if (need && String(t || '') !== need) throw new Error('存取碼錯誤，請到「⚙ 設定」輸入正確的存取碼');
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('伺服器忙碌中，請稍後再試');
  try { return fn(); } finally { lock.releaseLock(); }
}

// ===================== 工作表 =====================
function ctx_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = CONFIG.SHEET_NAME ? ss.getSheetByName(CONFIG.SHEET_NAME) : ss.getSheets()[0];
  if (!sh) throw new Error('找不到工作表：' + CONFIG.SHEET_NAME);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const col = {};
  Object.keys(CONFIG.COLS).forEach(k => {
    const i = headers.findIndex(h => CONFIG.COLS[k].indexOf(h) >= 0);
    col[k] = i >= 0 ? i + 1 : 0;
  });
  const st = {};
  CONFIG.STATUSES.forEach(s => { const i = headers.indexOf(s); st[s] = i >= 0 ? i + 1 : 0; });
  if (!col.name && !col.barcode) throw new Error('找不到「品名」或「國際條碼」欄，請檢查標題列或 CONFIG.COLS');
  return { ss, sh, headers, col, st };
}
// 欄位不存在時，在最右邊新增一欄
function ensureCol_(c, key) {
  if (c.col[key]) return c.col[key];
  const n = c.sh.getLastColumn() + 1;
  c.sh.getRange(CONFIG.HEADER_ROW, n).setValue(CONFIG.COLS[key][0]);
  c.col[key] = n; c.headers.push(CONFIG.COLS[key][0]);
  return n;
}
function readRows_(c) {
  const first = CONFIG.HEADER_ROW + 1, last = c.sh.getLastRow();
  if (last < first) return [];
  const vals = c.sh.getRange(first, 1, last - first + 1, c.sh.getLastColumn()).getValues();
  const out = [];
  vals.forEach((v, i) => {
    const it = toItem_(c, v, first + i);
    if (it.name || it.barcode) out.push(it);
  });
  return out;
}
function readRow_(c, row) {
  if (!row || row <= CONFIG.HEADER_ROW || row > c.sh.getLastRow()) return null;
  const v = c.sh.getRange(row, 1, 1, c.sh.getLastColumn()).getValues()[0];
  return toItem_(c, v, row);
}
function toItem_(c, v, row) {
  const g = k => c.col[k] ? v[c.col[k] - 1] : '';
  const q = {};
  CONFIG.STATUSES.forEach(s => { q[s] = c.st[s] ? (Number(v[c.st[s] - 1]) || 0) : 0; });
  const book = g('book'), lc = g('lastCounted');
  return {
    row,
    name: String(g('name')).trim(),
    model: String(g('model')).trim(),
    barcode: normBarcode_(g('barcode')),
    location: String(g('location')).trim(),
    book: c.col.book ? (book === '' || book === null ? null : Number(book) || 0) : null,
    lastCounted: lc instanceof Date ? lc.toISOString() : (lc ? String(lc) : ''),
    counter: String(g('counter')).trim(),
    quantities: q,
  };
}
function normBarcode_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (typeof v === 'number') return v.toFixed(0);
  return String(v).trim();
}
function sameBarcode_(a, b) {
  a = String(a || '').trim(); b = String(b || '').trim();
  if (!a || !b) return false;
  if (a === b) return true;
  return /^\d+$/.test(a) && /^\d+$/.test(b) && a.replace(/^0+/, '') === b.replace(/^0+/, '');  // 容忍開頭 0 被吃掉
}
function findByBarcode_(c, barcode) {
  return readRows_(c).find(it => sameBarcode_(it.barcode, barcode)) || null;
}

// ===================== 查詢 =====================
function lookup_(p) {
  const bc = String(p.barcode || '').trim();
  if (!bc) throw new Error('缺少條碼');
  const it = findByBarcode_(ctx_(), bc);
  return it ? Object.assign({ found: true }, it) : { found: false, barcode: bc };
}
function search_(p) {
  const q = String(p.q || '').trim().toLowerCase();
  if (!q) return { results: [] };
  const words = q.split(/\s+/);
  const res = readRows_(ctx_()).filter(it => {
    const s = (it.name + ' ' + it.model + ' ' + it.barcode + ' ' + it.location).toLowerCase();
    return words.every(w => s.indexOf(w) >= 0);
  });
  return { results: res.slice(0, 50) };
}
function list_() {
  const c = ctx_();
  return {
    version: VERSION,
    results: readRows_(c),
    roundStart: PropertiesService.getScriptProperties().getProperty('ROUND_START') || '',
    columns: { book: !!c.col.book, location: !!c.col.location },
  };
}

// ===================== 寫入 =====================
function resolveRow_(c, p) {
  let it = readRow_(c, Number(p.row));
  // 有條碼時，確認列號沒有因為插入/刪除列而跑掉
  if (p.barcode && (!it || !sameBarcode_(it.barcode, p.barcode))) it = findByBarcode_(c, p.barcode);
  if (!it) throw new Error('找不到這個商品（列號或條碼不符），請重新查詢');
  return it;
}
function stamp_(c, row, operator) {
  c.sh.getRange(row, ensureCol_(c, 'lastCounted')).setValue(new Date());
  c.sh.getRange(row, ensureCol_(c, 'counter')).setValue(String(operator || ''));
}
function log_(c, op, it, extra) {
  let lg = c.ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!lg) {
    lg = c.ss.insertSheet(CONFIG.LOG_SHEET);
    lg.appendRow(['時間', '盤點人員', '動作', '列', '國際條碼', '品名', '型號', '狀態', '寫入方式', '本次數量', '原數量', '新數量', '備註']);
    lg.setFrozenRows(1);
  }
  const e = extra || {};
  lg.appendRow([new Date(), op || '', e.action || '', it ? it.row : '', it ? it.barcode : '', it ? it.name : '', it ? it.model : '',
    e.status || '', e.mode || '', e.qty === undefined ? '' : e.qty, e.oldValue === undefined ? '' : e.oldValue,
    e.newValue === undefined ? '' : e.newValue, e.note || '']);
}

function save_(p) {
  const c = ctx_();
  const status = String(p.status || '');
  if (!c.st[status]) throw new Error('試算表沒有「' + status + '」欄');
  const qty = Number(p.qty);
  if (!isFinite(qty) || Math.floor(qty) !== qty || qty < 0) throw new Error('數量必須是 0 以上的整數');
  const mode = p.mode === 'set' ? 'set' : 'add';
  const it = resolveRow_(c, p);
  const cell = c.sh.getRange(it.row, c.st[status]);
  const oldValue = Number(cell.getValue()) || 0;
  const newValue = mode === 'set' ? qty : oldValue + qty;
  cell.setValue(newValue);
  stamp_(c, it.row, p.operator);
  log_(c, p.operator, it, { action: p.undo ? '撤銷' : '盤點', status, mode: mode === 'set' ? '覆蓋' : '累加', qty, oldValue, newValue });
  return { ok: true, row: it.row, status, added: qty, oldValue, newValue, mode };
}

function create_(p) {
  const c = ctx_();
  const name = String(p.name || '').trim(), bc = String(p.barcode || '').trim();
  if (!name) throw new Error('請輸入品名');
  if (bc) {
    const dup = findByBarcode_(c, bc);
    if (dup) throw new Error('條碼已存在：第 ' + dup.row + ' 列「' + dup.name + '」');
  }
  if (!c.col.name) ensureCol_(c, 'name');
  if (bc && !c.col.barcode) ensureCol_(c, 'barcode');
  if (p.location) ensureCol_(c, 'location');
  const row = Math.max(c.sh.getLastRow(), CONFIG.HEADER_ROW) + 1;
  c.sh.getRange(row, c.col.name).setValue(name);
  if (c.col.model && p.model) c.sh.getRange(row, c.col.model).setValue(String(p.model).trim());
  if (bc) c.sh.getRange(row, c.col.barcode).setNumberFormat('@').setValue(bc);   // 存成文字，保留開頭 0
  if (p.location) c.sh.getRange(row, c.col.location).setValue(String(p.location).trim());
  CONFIG.STATUSES.forEach(s => { if (c.st[s]) c.sh.getRange(row, c.st[s]).setValue(0); });
  const it = readRow_(c, row);
  log_(c, p.operator, it, { action: '新增商品' });
  return { ok: true, item: it };
}

function bindBarcode_(p) {
  const c = ctx_();
  const bc = String(p.barcode || '').trim();
  if (!bc) throw new Error('缺少條碼');
  const it = readRow_(c, Number(p.row));
  if (!it) throw new Error('找不到第 ' + p.row + ' 列');
  const dup = findByBarcode_(c, bc);
  if (dup && dup.row !== it.row) throw new Error('條碼已屬於第 ' + dup.row + ' 列「' + dup.name + '」');
  if (it.barcode && !sameBarcode_(it.barcode, bc) && !p.force) {
    throw new Error('此商品已有條碼 ' + it.barcode + '，不會覆蓋');
  }
  c.sh.getRange(it.row, ensureCol_(c, 'barcode')).setNumberFormat('@').setValue(bc);
  const out = readRow_(c, it.row);
  log_(c, p.operator, out, { action: '綁定條碼', note: it.barcode ? '原條碼 ' + it.barcode : '' });
  return { ok: true, item: out };
}

function setLocation_(p) {
  const c = ctx_();
  const it = resolveRow_(c, p);
  const loc = String(p.location || '').trim();
  c.sh.getRange(it.row, ensureCol_(c, 'location')).setValue(loc);
  log_(c, p.operator, it, { action: '設定儲位', note: (it.location || '(空)') + ' → ' + (loc || '(空)') });
  return { ok: true, item: readRow_(c, it.row) };
}

// 開始新一輪盤點：記下開始時間；可選擇先備份再把所有狀態數量歸零
function startRound_(p) {
  const c = ctx_();
  const now = new Date();
  let backup = '';
  if (p.reset) {
    backup = '備份_' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    c.sh.copyTo(c.ss).setName(backup);
    const first = CONFIG.HEADER_ROW + 1, n = c.sh.getLastRow() - CONFIG.HEADER_ROW;
    if (n > 0) CONFIG.STATUSES.forEach(s => {
      if (c.st[s]) c.sh.getRange(first, c.st[s], n, 1).setValue(0);
    });
  }
  PropertiesService.getScriptProperties().setProperty('ROUND_START', now.toISOString());
  log_(c, p.operator, null, { action: '開始新一輪盤點', note: p.reset ? '已歸零，備份於 ' + backup : '未歸零' });
  return { ok: true, roundStart: now.toISOString(), backup };
}
