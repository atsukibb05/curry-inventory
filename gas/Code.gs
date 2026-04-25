/**
 * カレー食品加工 在庫管理システム - バックエンド
 * Google Apps Script Web App
 *
 * == API 仕様 ==
 * すべて GET / POST。レスポンスは { ok: boolean, data?: any, error?: string }
 *
 * GET  ?action=list          → 在庫一覧 + カテゴリー + 設定 を返す
 * GET  ?action=history&limit=50 → 履歴を新しい順で返す
 * POST { action: 'inOut', code, type: '入庫'|'出庫', qty, operator, note }
 *      → 入出庫処理(在庫更新 + 履歴記録)
 * POST { action: 'addItem', code, name, category, unit, stock, minStock, expiry, note }
 *      → 新規商品登録
 * POST { action: 'updateItem', code, ...更新フィールド }
 *      → 商品マスタ更新
 *
 * == セットアップ ==
 * 1. このファイルを Apps Script プロジェクトに貼り付け
 * 2. SPREADSHEET_ID を自分のスプレッドシート ID に書き換え
 * 3. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 4. デプロイ URL を web/app.js の API_URL に貼る
 */

// ====== 設定 ======
const SPREADSHEET_ID = '1UgENYWd0qs9eh0V2263De4c6LxbLeCM-2NsdViQukJU'; // ← 変更必須
const SHEET_STOCK = '在庫マスタ';
const SHEET_HISTORY = '履歴';
const SHEET_CATEGORY = 'カテゴリー';
const SHEET_CONFIG = '設定';

// 在庫マスタの列インデックス(1始まり)
const COL_CODE = 1;
const COL_NAME = 2;
const COL_CATEGORY = 3;
const COL_UNIT = 4;
const COL_STOCK = 5;
const COL_MIN_STOCK = 6;
const COL_EXPIRY = 7;
const COL_NOTE = 8;
const COL_UPDATED = 9;

// ====== エントリーポイント ======
function doGet(e) {
  return handle(e, false);
}

function doPost(e) {
  return handle(e, true);
}

function handle(e, isPost) {
  try {
    let params;
    if (isPost && e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else {
      params = e.parameter || {};
    }

    const action = params.action;
    let result;

    switch (action) {
      case 'list':
        result = getList();
        break;
      case 'history':
        result = getHistory(parseInt(params.limit) || 50);
        break;
      case 'inOut':
        result = doInOut(params);
        break;
      case 'addItem':
        result = addItem(params);
        break;
      case 'updateItem':
        result = updateItem(params);
        break;
      case 'init':
        // 初期データ投入(セットアップ時に1回だけ使う)
        result = initializeSheets();
        break;
      default:
        throw new Error('Unknown action: ' + action);
    }

    return jsonResponse({ ok: true, data: result });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====== 在庫取得 ======
function getList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const stockSheet = ss.getSheetByName(SHEET_STOCK);
  const catSheet = ss.getSheetByName(SHEET_CATEGORY);
  const configSheet = ss.getSheetByName(SHEET_CONFIG);

  // 在庫
  const stockValues = stockSheet.getDataRange().getValues();
  const stockHeader = stockValues.shift();
  const items = stockValues
    .filter(row => row[COL_CODE - 1])
    .map(row => ({
      code: row[COL_CODE - 1],
      name: row[COL_NAME - 1],
      category: row[COL_CATEGORY - 1],
      unit: row[COL_UNIT - 1],
      stock: Number(row[COL_STOCK - 1]) || 0,
      minStock: Number(row[COL_MIN_STOCK - 1]) || 0,
      expiry: row[COL_EXPIRY - 1] ? Utilities.formatDate(new Date(row[COL_EXPIRY - 1]), 'Asia/Tokyo', 'yyyy-MM-dd') : '',
      note: row[COL_NOTE - 1] || '',
      updated: row[COL_UPDATED - 1] ? Utilities.formatDate(new Date(row[COL_UPDATED - 1]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : ''
    }));

  // カテゴリー
  const catValues = catSheet.getDataRange().getValues();
  catValues.shift();
  const categories = catValues
    .filter(r => r[0])
    .map(r => ({ name: r[0], order: Number(r[1]) || 0, icon: r[2] || '' }))
    .sort((a, b) => a.order - b.order);

  // 設定
  const config = {};
  if (configSheet) {
    const cfgValues = configSheet.getDataRange().getValues();
    cfgValues.shift();
    cfgValues.forEach(r => { if (r[0]) config[r[0]] = r[1]; });
  }

  return { items, categories, config };
}

// ====== 履歴取得 ======
function getHistory(limit) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_HISTORY);
  const values = sheet.getDataRange().getValues();
  values.shift(); // ヘッダー
  const rows = values
    .filter(r => r[0])
    .map(r => ({
      datetime: r[0] ? Utilities.formatDate(new Date(r[0]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss') : '',
      operator: r[1],
      code: r[2],
      name: r[3],
      type: r[4],
      qty: Number(r[5]) || 0,
      stockAfter: Number(r[6]) || 0,
      note: r[7] || ''
    }))
    .reverse() // 新しい順
    .slice(0, limit);
  return rows;
}

// ====== 入出庫処理 ======
function doInOut(params) {
  const { code, type, qty, operator, note } = params;
  if (!code || !type || !qty) throw new Error('引数不足: code, type, qty が必要');
  const qtyNum = Number(qty);
  if (qtyNum <= 0) throw new Error('数量は1以上にしてください');
  if (type !== '入庫' && type !== '出庫') throw new Error('type は 入庫 または 出庫');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_STOCK);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    let item = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][COL_CODE - 1] === code) {
        targetRow = i + 1;
        item = data[i];
        break;
      }
    }
    if (targetRow === -1) throw new Error('商品コードが見つかりません: ' + code);

    const currentStock = Number(item[COL_STOCK - 1]) || 0;
    const newStock = type === '入庫' ? currentStock + qtyNum : currentStock - qtyNum;
    if (newStock < 0) throw new Error('在庫が足りません(現在 ' + currentStock + ')');

    // 更新
    sheet.getRange(targetRow, COL_STOCK).setValue(newStock);
    sheet.getRange(targetRow, COL_UPDATED).setValue(new Date());

    // 履歴
    logHistory({
      operator: operator || '不明',
      code: code,
      name: item[COL_NAME - 1],
      type: type,
      qty: qtyNum,
      stockAfter: newStock,
      note: note || ''
    });

    return { code, name: item[COL_NAME - 1], stockBefore: currentStock, stockAfter: newStock };
  } finally {
    lock.releaseLock();
  }
}

function logHistory(rec) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_HISTORY);
  sheet.appendRow([
    new Date(),
    rec.operator,
    rec.code,
    rec.name,
    rec.type,
    rec.qty,
    rec.stockAfter,
    rec.note
  ]);
}

// ====== 商品追加・更新 ======
function addItem(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_STOCK);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_CODE - 1] === p.code) throw new Error('商品コードが重複: ' + p.code);
  }
  sheet.appendRow([
    p.code, p.name, p.category, p.unit,
    Number(p.stock) || 0,
    Number(p.minStock) || 0,
    p.expiry || '',
    p.note || '',
    new Date()
  ]);
  return { code: p.code };
}

function updateItem(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_STOCK);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_CODE - 1] === p.code) {
      const row = i + 1;
      if (p.name !== undefined) sheet.getRange(row, COL_NAME).setValue(p.name);
      if (p.category !== undefined) sheet.getRange(row, COL_CATEGORY).setValue(p.category);
      if (p.unit !== undefined) sheet.getRange(row, COL_UNIT).setValue(p.unit);
      if (p.minStock !== undefined) sheet.getRange(row, COL_MIN_STOCK).setValue(Number(p.minStock));
      if (p.expiry !== undefined) sheet.getRange(row, COL_EXPIRY).setValue(p.expiry);
      if (p.note !== undefined) sheet.getRange(row, COL_NOTE).setValue(p.note);
      sheet.getRange(row, COL_UPDATED).setValue(new Date());
      return { code: p.code };
    }
  }
  throw new Error('商品コードが見つかりません: ' + p.code);
}

// ====== 初期化(初回のみ手動実行) ======
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 在庫マスタ
  let s = ss.getSheetByName(SHEET_STOCK) || ss.insertSheet(SHEET_STOCK);
  s.clear();
  s.getRange(1, 1, 1, 9).setValues([[
    '商品コード', '商品名', 'カテゴリー', '単位', '現在庫', '最低在庫', '賞味期限', '備考', '更新日時'
  ]]);
  s.setFrozenRows(1);

  // カテゴリー
  let c = ss.getSheetByName(SHEET_CATEGORY) || ss.insertSheet(SHEET_CATEGORY);
  c.clear();
  c.getRange(1, 1, 1, 3).setValues([['カテゴリー名', '表示順', 'アイコン']]);
  c.getRange(2, 1, 6, 3).setValues([
    ['香辛料', 1, '🌶️'],
    ['野菜', 2, '🥕'],
    ['肉類', 3, '🍖'],
    ['乳製品', 4, '🥛'],
    ['調味料', 5, '🧂'],
    ['包材', 6, '📦']
  ]);
  c.setFrozenRows(1);

  // 履歴
  let h = ss.getSheetByName(SHEET_HISTORY) || ss.insertSheet(SHEET_HISTORY);
  h.clear();
  h.getRange(1, 1, 1, 8).setValues([[
    '日時', '操作者', '商品コード', '商品名', '区分', '数量', '操作後在庫', '備考'
  ]]);
  h.setFrozenRows(1);

  // 設定
  let cfg = ss.getSheetByName(SHEET_CONFIG) || ss.insertSheet(SHEET_CONFIG);
  cfg.clear();
  cfg.getRange(1, 1, 1, 2).setValues([['キー', '値']]);
  cfg.getRange(2, 1, 2, 2).setValues([
    ['alert_threshold_ratio', 1.0],
    ['default_operator', '田中']
  ]);
  cfg.setFrozenRows(1);

  // サンプル商品
  s.getRange(2, 1, 10, 9).setValues([
    ['SP001', 'クミンパウダー', '香辛料', 'g', 500, 200, '2026-12-31', '', new Date()],
    ['SP002', 'ターメリック', '香辛料', 'g', 300, 200, '2026-12-31', '', new Date()],
    ['SP003', 'コリアンダー', '香辛料', 'g', 400, 200, '2026-12-31', '', new Date()],
    ['VG001', '玉ねぎ', '野菜', '個', 50, 20, '2026-05-10', '北海道産', new Date()],
    ['VG002', 'にんにく', '野菜', 'kg', 2, 1, '2026-05-15', '', new Date()],
    ['VG003', '生姜', '野菜', 'kg', 1.5, 0.5, '2026-05-12', '', new Date()],
    ['MT001', '鶏もも肉', '肉類', 'kg', 10, 5, '2026-04-28', '冷蔵', new Date()],
    ['DR001', 'ヨーグルト', '乳製品', 'g', 2000, 1000, '2026-05-05', '', new Date()],
    ['SC001', 'トマト缶', '調味料', '缶', 24, 12, '2027-06-30', '400g缶', new Date()],
    ['PK001', '使い捨て容器500ml', '包材', '個', 100, 50, '', '', new Date()]
  ]);

  return { message: '初期化完了。サンプルデータを投入しました。' };
}
