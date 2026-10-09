/**
 * 倉庫盤點掃描 — 後端 API (Google Apps Script)  v2
 * 含「每人密碼登入 + 修改紀錄」，並新增：新增商品 / 綁定條碼 / 儲位 / 盤點進度與差異 / 開始新一輪
 *
 * 綁定在你的試算表上：擴充功能 → Apps Script，把整個 Code.gs 內容換成本檔。
 *
 * ★ 需要的分頁：
 *   1) 「使用者」分頁：第一列標題為  姓名 | 密碼 | 啟用  （一人一列；啟用留空或填是=可用，填 否/N=停用）
 *   2) 「紀錄」分頁：可不用自己建，第一次有人存入時會自動建立並寫入標題。
 *
 * ★ 庫存分頁的選用欄位（沒有也能用）：
 *   - 帳面數量：有這欄才會計算盤盈 / 盤虧
 *   - 儲位、最後盤點時間、盤點人員：第一次用到時會自動在最右邊新增
 *
 * ★ 改完務必「部署 → 管理部署作業 → 編輯 → 版本：新版本 → 部署」，改動才會生效。
 */

// ===================== 設定 =====================
const SHEET_NAME = '';   // 留空 = 第一個工作表（庫存資料）；或填分頁名稱
const HEADER_ROW = 1;

const BARCODE_HEADER = '國際條碼';
const NAME_HEADER    = '品名';
const MODEL_HEADER   = '型號';
const STATUS_COLUMNS = ['新品', '福利品', '瑕疵', '報廢', '樣品', '採購樣品'];

const BOOK_HEADER     = '帳面數量';      // 選用
const LOCATION_HEADER = '儲位';          // 選用（自動新增）
const COUNTED_HEADER  = '最後盤點時間';  // 自動新增
const COUNTER_HEADER  = '盤點人員';      // 自動新增

const USERS_SHEET = '使用者';  // 使用者清單分頁（姓名/密碼/啟用）
const LOG_SHEET   = '紀錄';    // 修改紀錄分頁（自動建立）
const SEARCH_LIMIT = 30;
const API_VERSION = 2;
// ==============================================


function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (!p.action) return json_({ ok: true, message: '倉庫盤點 API 運作中', version: API_VERSION }); // 存活測試（免密碼）
    const user = findUser_(p.code);
    if (p.action === 'auth') {
      return json_(user ? { ok: true, name: user.name } : { ok: false, error: '密碼錯誤或已停用' });
    }
    if (!user) return json_({ ok: false, error: '未授權（密碼錯誤或已停用）', needAuth: true });

    if (p.action === 'lookup') return json_(lookupProduct(p.barcode));
    if (p.action === 'search') return json_(searchProducts(p.q));
    if (p.action === 'list')   return json_(listAll());
    return json_({ ok: false, error: '未知的動作' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const user = findUser_(body.code);
    if (!user) return json_({ ok: false, error: '未授權（密碼錯誤或已停用）', needAuth: true });
    switch (body.action) {
      case 'save':        return json_(withLock_(() => saveCount(body, user)));
      case 'create':      return json_(withLock_(() => createProduct(body, user)));
      case 'bindBarcode': return json_(withLock_(() => bindBarcode(body, user)));
      case 'setLocation': return json_(withLock_(() => setLocation(body, user)));
      case 'startRound':  return json_(withLock_(() => startRound(body, user)));
    }
    return json_({ ok: false, error: '未知的動作' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}


// ============ 登入驗證 / 紀錄 ============

/* 依密碼在「使用者」分頁比對，回傳 {name} 或 null */
function findUser_(code) {
  code = String(code || '').trim();
  if (!code) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);
  if (!sheet) throw new Error('找不到「' + USERS_SHEET + '」分頁（請新增，第一列標題：姓名 / 密碼 / 啟用）');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return null;

  const hmap = {};
  sheet.getRange(1, 1, 1, lastCol).getValues()[0].forEach((h, i) => { const k = String(h).trim(); if (k) hmap[k] = i; });
  const ci = hmap['密碼'], ni = hmap['姓名'], ei = hmap['啟用'];
  if (ci === undefined) throw new Error('「' + USERS_SHEET + '」分頁找不到「密碼」欄');

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (let i = 0; i < data.length; i++) {
    const pass = String(data[i][ci]).trim();
    if (pass && pass === code) {
      if (ei !== undefined) {
        const en = String(data[i][ei]).trim().toLowerCase();
        if (['n', 'no', '否', '停用', '0', 'false', 'x'].indexOf(en) !== -1) return null; // 已停用
      }
      const name = (ni !== undefined) ? String(data[i][ni]).trim() : '';
      return { name: name || '(未命名)' };
    }
  }
  return null;
}

/* 追加一筆修改紀錄到「紀錄」分頁（不存在則自動建立）
 * o.modeText 有值時直接用（非盤點動作，如「新增商品」），否則依 o.mode 顯示 覆蓋 / 累加 */
function appendLog_(o) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.appendRow(['時間', '操作人', '品名', '型號', '國際條碼', '狀態', '本次數量', '寫入方式', '前值', '後值', '列', '備註']);
  } else if (String(sheet.getRange(1, 12).getValue()).trim() === '') {
    sheet.getRange(1, 12).setValue('備註');   // 舊紀錄表補上「備註」標題
  }
  const modeText = o.modeText !== undefined ? o.modeText : (o.mode === 'set' ? '覆蓋' : '累加');
  sheet.appendRow([
    new Date(), o.user || '', o.name || '', o.model || '', o.barcode || '',
    o.status || '', o.qty === undefined ? '' : o.qty, modeText,
    o.prev === undefined ? '' : o.prev, o.next === undefined ? '' : o.next, o.row || '', o.note || ''
  ]);
}


// ============ 資料存取 ============

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (SHEET_NAME) {
    const s = ss.getSheetByName(SHEET_NAME);
    if (!s) throw new Error('找不到工作表：' + SHEET_NAME);
    return s;
  }
  return ss.getSheets()[0];
}

function getHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { const key = String(h).trim(); if (key) map[key] = i + 1; });
  return map;
}

/* 欄位不存在時，在最右邊新增一欄，回傳欄號 */
function ensureCol_(sheet, headerMap, header) {
  if (headerMap[header]) return headerMap[header];
  const col = sheet.getLastColumn() + 1;
  sheet.getRange(HEADER_ROW, col).setValue(header);
  headerMap[header] = col;
  return col;
}

function normalizeBarcode_(v) {
  if (v === null || v === undefined) return '';
  let s = String(v).trim();
  if (s === '') return '';
  if (/e/i.test(s) && !isNaN(Number(s))) s = Number(s).toFixed(0);
  return s.replace(/[^0-9]/g, '');
}

function locate_(barcode) {
  const sheet = getSheet_();
  const headerMap = getHeaderMap_(sheet);
  const barcodeCol = headerMap[BARCODE_HEADER];
  if (!barcodeCol) throw new Error('表頭找不到「' + BARCODE_HEADER + '」欄，請確認標題文字。');

  const target = normalizeBarcode_(barcode);
  const lastRow = sheet.getLastRow();
  let row = null;
  if (target && lastRow > HEADER_ROW) {
    const values = sheet.getRange(HEADER_ROW + 1, barcodeCol, lastRow - HEADER_ROW, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (normalizeBarcode_(values[i][0]) === target) { row = HEADER_ROW + 1 + i; break; }
    }
  }
  return { sheet, headerMap, row };
}

function rowToProduct_(sheet, headerMap, row, rowValues) {
  const getVal = (h) => headerMap[h] ? rowValues[headerMap[h] - 1] : '';
  const quantities = {};
  STATUS_COLUMNS.forEach(st => { quantities[st] = Number(getVal(st)) || 0; });
  const book = getVal(BOOK_HEADER), counted = getVal(COUNTED_HEADER);
  return {
    found: true, row: row,
    barcode: String(getVal(BARCODE_HEADER) || ''),
    name: getVal(NAME_HEADER), model: getVal(MODEL_HEADER),
    quantities: quantities,
    location: String(getVal(LOCATION_HEADER) || '').trim(),
    book: (headerMap[BOOK_HEADER] && book !== '' && book !== null) ? (Number(book) || 0) : null,
    lastCounted: counted instanceof Date ? counted.toISOString() : (counted ? String(counted) : ''),
    counter: String(getVal(COUNTER_HEADER) || '').trim()
  };
}

function readProduct_(sheet, headerMap, row) {
  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  return rowToProduct_(sheet, headerMap, row, rowValues);
}

function lookupProduct(barcode) {
  const { sheet, headerMap, row } = locate_(barcode);
  if (!row) return { found: false, barcode: barcode };
  const p = readProduct_(sheet, headerMap, row);
  p.barcode = String(barcode);
  return p;
}

function searchProducts(q) {
  q = String(q || '').trim().toLowerCase();
  if (!q) return { ok: true, results: [] };
  const sheet = getSheet_();
  const headerMap = getHeaderMap_(sheet);
  const nameCol = headerMap[NAME_HEADER];
  const modelCol = headerMap[MODEL_HEADER];
  const locCol = headerMap[LOCATION_HEADER];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= HEADER_ROW) return { ok: true, results: [] };

  const data = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol).getValues();
  const results = [];
  for (let i = 0; i < data.length && results.length < SEARCH_LIMIT; i++) {
    const rowValues = data[i];
    const name  = nameCol  ? String(rowValues[nameCol - 1])  : '';
    const model = modelCol ? String(rowValues[modelCol - 1]) : '';
    const loc   = locCol   ? String(rowValues[locCol - 1])   : '';
    if (name.toLowerCase().indexOf(q) !== -1 || model.toLowerCase().indexOf(q) !== -1 || loc.toLowerCase().indexOf(q) !== -1) {
      results.push(rowToProduct_(sheet, headerMap, HEADER_ROW + 1 + i, rowValues));
    }
  }
  return { ok: true, results: results };
}

function listAll() {
  const sheet = getSheet_();
  const headerMap = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const info = {
    ok: true, version: API_VERSION,
    roundStart: PropertiesService.getScriptProperties().getProperty('ROUND_START') || '',
    columns: { book: !!headerMap[BOOK_HEADER], location: !!headerMap[LOCATION_HEADER] }
  };
  if (lastRow <= HEADER_ROW) return Object.assign(info, { results: [] });

  const data = sheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol).getValues();
  const results = [];
  for (let i = 0; i < data.length; i++) {
    const prod = rowToProduct_(sheet, headerMap, HEADER_ROW + 1 + i, data[i]);
    if (!String(prod.name).trim() && !String(prod.barcode).trim()) continue;
    results.push(prod);
  }
  return Object.assign(info, { results: results });
}

/* 依條碼（優先）或列號找到商品列；找不到回傳 null */
function resolveRow_(params) {
  let sheet, headerMap, row = null;
  if (params.barcode && String(params.barcode).trim() !== '') {
    const loc = locate_(params.barcode);
    sheet = loc.sheet; headerMap = loc.headerMap; row = loc.row;
  }
  if (!sheet) { sheet = getSheet_(); headerMap = getHeaderMap_(sheet); }
  if (!row && params.row) {
    const r = Number(params.row);
    if (r > HEADER_ROW && r <= sheet.getLastRow()) row = r;
  }
  return { sheet, headerMap, row };
}

/* 寫入數量並記錄操作人。params: { barcode?, row?, name?, model?, status, qty, mode, undo? } */
function saveCount(params, user) {
  const status = params.status;
  const qty = Number(params.qty);
  const mode = params.mode;
  if (isNaN(qty)) throw new Error('數量無效');
  if (STATUS_COLUMNS.indexOf(status) === -1) throw new Error('狀態無效：' + status);

  const { sheet, headerMap, row } = resolveRow_(params);
  if (!row) return { ok: false, message: '找不到要寫入的商品' };

  const col = headerMap[status];
  if (!col) throw new Error('表頭找不到「' + status + '」欄');

  const cell = sheet.getRange(row, col);
  const current = Number(cell.getValue()) || 0;
  const newVal = (mode === 'set') ? qty : current + qty;
  cell.setValue(newVal);

  // 記下最後盤點時間與人員（盤點進度 / 未盤清單用）
  sheet.getRange(row, ensureCol_(sheet, headerMap, COUNTED_HEADER)).setValue(new Date());
  sheet.getRange(row, ensureCol_(sheet, headerMap, COUNTER_HEADER)).setValue(user ? user.name : '');

  appendLog_({
    user: user ? user.name : '',
    name: params.name || '', model: params.model || '', barcode: params.barcode || '',
    status: status, qty: qty, mode: mode, prev: current, next: newVal, row: row,
    note: params.undo ? '撤銷' : ''
  });

  return { ok: true, row: row, status: status, previous: current, added: qty, newValue: newVal, mode: mode, user: user ? user.name : '' };
}

/* 新增商品。params: { barcode?, name, model?, location? } */
function createProduct(params, user) {
  const name = String(params.name || '').trim();
  const barcode = normalizeBarcode_(params.barcode);
  if (!name) throw new Error('請輸入品名');
  const sheet = getSheet_();
  const headerMap = getHeaderMap_(sheet);
  if (barcode) {
    const dup = locate_(barcode);
    if (dup.row) throw new Error('條碼已存在於第 ' + dup.row + ' 列');
  }
  const row = Math.max(sheet.getLastRow(), HEADER_ROW) + 1;
  sheet.getRange(row, ensureCol_(sheet, headerMap, NAME_HEADER)).setValue(name);
  if (params.model) sheet.getRange(row, ensureCol_(sheet, headerMap, MODEL_HEADER)).setValue(String(params.model).trim());
  if (barcode) sheet.getRange(row, ensureCol_(sheet, headerMap, BARCODE_HEADER)).setNumberFormat('@').setValue(barcode); // 文字格式，保留開頭 0
  if (params.location) sheet.getRange(row, ensureCol_(sheet, headerMap, LOCATION_HEADER)).setValue(String(params.location).trim());
  STATUS_COLUMNS.forEach(st => { if (headerMap[st]) sheet.getRange(row, headerMap[st]).setValue(0); });

  const item = readProduct_(sheet, headerMap, row);
  appendLog_({ user: user.name, name: name, model: item.model, barcode: barcode, row: row, modeText: '新增商品' });
  return { ok: true, item: item };
}

/* 把條碼綁定到既有商品。params: { row, barcode, force? } */
function bindBarcode(params, user) {
  const barcode = normalizeBarcode_(params.barcode);
  if (!barcode) throw new Error('缺少條碼');
  const sheet = getSheet_();
  const headerMap = getHeaderMap_(sheet);
  const row = Number(params.row);
  if (!(row > HEADER_ROW && row <= sheet.getLastRow())) throw new Error('找不到第 ' + params.row + ' 列');
  const dup = locate_(barcode);
  if (dup.row && dup.row !== row) throw new Error('條碼已屬於第 ' + dup.row + ' 列');
  const before = readProduct_(sheet, headerMap, row);
  if (before.barcode && normalizeBarcode_(before.barcode) !== barcode && !params.force) {
    throw new Error('此商品已有條碼 ' + before.barcode + '，不會覆蓋');
  }
  sheet.getRange(row, ensureCol_(sheet, headerMap, BARCODE_HEADER)).setNumberFormat('@').setValue(barcode);
  const item = readProduct_(sheet, headerMap, row);
  appendLog_({ user: user.name, name: item.name, model: item.model, barcode: barcode, row: row, modeText: '綁定條碼',
    note: before.barcode ? '原條碼 ' + before.barcode : '' });
  return { ok: true, item: item };
}

/* 設定儲位。params: { barcode?, row, location } */
function setLocation(params, user) {
  const { sheet, headerMap, row } = resolveRow_(params);
  if (!row) throw new Error('找不到這個商品，請重新查詢');
  const before = readProduct_(sheet, headerMap, row);
  const loc = String(params.location || '').trim();
  sheet.getRange(row, ensureCol_(sheet, headerMap, LOCATION_HEADER)).setValue(loc);
  appendLog_({ user: user.name, name: before.name, model: before.model, barcode: before.barcode, row: row, modeText: '設定儲位',
    note: (before.location || '(空)') + ' → ' + (loc || '(空)') });
  return { ok: true, item: readProduct_(sheet, headerMap, row) };
}

/* 開始新一輪盤點：記下開始時間；reset=true 時先備份整張表，再把狀態數量全部歸零 */
function startRound(params, user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_();
  const headerMap = getHeaderMap_(sheet);
  const now = new Date();
  let backup = '';
  if (params.reset) {
    backup = '備份_' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    sheet.copyTo(ss).setName(backup);
    const n = sheet.getLastRow() - HEADER_ROW;
    if (n > 0) STATUS_COLUMNS.forEach(st => {
      if (headerMap[st]) sheet.getRange(HEADER_ROW + 1, headerMap[st], n, 1).setValue(0);
    });
  }
  PropertiesService.getScriptProperties().setProperty('ROUND_START', now.toISOString());
  appendLog_({ user: user.name, modeText: '開始新一輪盤點', note: params.reset ? '已歸零，備份於 ' + backup : '數量未歸零' });
  return { ok: true, roundStart: now.toISOString(), backup: backup };
}
