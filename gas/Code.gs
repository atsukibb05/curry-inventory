/**
 * カレー食品加工 在庫管理システム - バックエンド
 * Google Apps Script Web App
 *
 * == API 仕様 ==
 * すべて GET / POST。レスポンスは { ok: boolean, data?: any, error?: string }
 *
 * GET  ?action=list          → 在庫一覧 + カテゴリー + 拠点 + 設定
 * GET  ?action=history&limit=50 → 履歴を新しい順
 * POST { action: 'inOut', code, location, type: '入庫'|'出庫', qty, operator, note }
 *      → 入出庫処理(指定拠点の在庫更新 + 履歴記録)
 * POST { action: 'addItem', code, name, category, unit, stocks, minStock, expiry, note }
 *      → 新規商品登録 (stocks は { 塩山店:数, 百間店:数, ... })
 * POST { action: 'updateItem', code, ...更新フィールド }
 *      → 商品マスタ更新
 *
 * == セットアップ ==
 * 1. このファイルを Apps Script プロジェクトに貼り付け
 * 2. SPREADSHEET_ID を自分のスプレッドシート ID に書き換え
 * 3. initializeSheets を実行 → 拠点別シートが構築される
 * 4. 「デプロイ」→「デプロイを管理」→ 編集 → 新バージョン → デプロイ
 */

// ====== 設定 ======
const SPREADSHEET_ID = '1UgENYWd0qs9eh0V2263De4c6LxbLeCM-2NsdViQukJU';
const SHEET_STOCK = '在庫マスタ';
const SHEET_HISTORY = '履歴';
const SHEET_CATEGORY = 'カテゴリー';
const SHEET_LOCATION = '拠点';
const SHEET_CONFIG = '設定';

// 在庫マスタの列インデックス(1始まり)
const COL_CODE = 1;
const COL_NAME = 2;
const COL_CATEGORY = 3;
const COL_UNIT = 4;
const COL_STOCK_SHIOYAMA = 5;
const COL_STOCK_HYAKKEN = 6;
const COL_STOCK_KAWAGUCHIKO = 7;
const COL_STOCK_MS = 8;
const COL_TOTAL = 9;       // 数式 =SUM(E:H)
const COL_MIN_STOCK = 10;
const COL_EXPIRY = 11;
const COL_NOTE = 12;
const COL_UPDATED = 13;
const STOCK_COL_COUNT = 13;

// 拠点の定義(拠点名と在庫マスタ列の対応)
const LOCATIONS = [
  { name: '塩山店',   type: '店舗', order: 1, col: COL_STOCK_SHIOYAMA },
  { name: '百間店',   type: '店舗', order: 2, col: COL_STOCK_HYAKKEN },
  { name: '河口湖店', type: '店舗', order: 3, col: COL_STOCK_KAWAGUCHIKO },
  { name: 'MS',       type: '倉庫', order: 4, col: COL_STOCK_MS },
];

function findLocation(name) {
  return LOCATIONS.find(l => l.name === name);
}

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
      case 'list':       result = getList(); break;
      case 'history':    result = getHistory(parseInt(params.limit) || 50); break;
      case 'inOut':      result = doInOut(params); break;
      case 'addItem':    result = addItem(params); break;
      case 'updateItem': result = updateItem(params); break;
      case 'init':       result = initializeSheets(); break;
      default: throw new Error('Unknown action: ' + action);
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

// 在庫セルを安全に数値化(列書式が日付などになっていた場合でも 0 にフォールバック)
function safeStockNum(raw) {
  if (raw === '' || raw === null || raw === undefined) return 0;
  if (raw instanceof Date) return 0; // 日付書式の混入を防御
  const n = Number(raw);
  return isFinite(n) ? n : 0;
}

// ====== 在庫取得 ======
function getList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const stockSheet = ss.getSheetByName(SHEET_STOCK);
  const catSheet = ss.getSheetByName(SHEET_CATEGORY);
  const locSheet = ss.getSheetByName(SHEET_LOCATION);
  const configSheet = ss.getSheetByName(SHEET_CONFIG);

  // 在庫
  const stockValues = stockSheet.getDataRange().getValues();
  stockValues.shift(); // ヘッダー除去
  const items = stockValues
    .filter(row => row[COL_CODE - 1])
    .map(row => {
      const stocks = {};
      let total = 0;
      LOCATIONS.forEach(loc => {
        const v = safeStockNum(row[loc.col - 1]);
        stocks[loc.name] = v;
        total += v;
      });
      return {
        code: row[COL_CODE - 1],
        name: row[COL_NAME - 1],
        category: row[COL_CATEGORY - 1],
        unit: row[COL_UNIT - 1],
        stocks: stocks,
        total: total,
        minStock: Number(row[COL_MIN_STOCK - 1]) || 0,
        expiry: row[COL_EXPIRY - 1] ? Utilities.formatDate(new Date(row[COL_EXPIRY - 1]), 'Asia/Tokyo', 'yyyy-MM-dd') : '',
        note: row[COL_NOTE - 1] || '',
        updated: row[COL_UPDATED - 1] ? Utilities.formatDate(new Date(row[COL_UPDATED - 1]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : ''
      };
    });

  // カテゴリー
  const catValues = catSheet.getDataRange().getValues();
  catValues.shift();
  const categories = catValues
    .filter(r => r[0])
    .map(r => ({ name: r[0], order: Number(r[1]) || 0, icon: r[2] || '' }))
    .sort((a, b) => a.order - b.order);

  // 拠点
  let locations = [];
  if (locSheet) {
    const locValues = locSheet.getDataRange().getValues();
    locValues.shift();
    locations = locValues
      .filter(r => r[0])
      .map(r => ({ name: r[0], type: r[1] || '店舗', order: Number(r[2]) || 0 }))
      .sort((a, b) => a.order - b.order);
  }
  if (locations.length === 0) {
    locations = LOCATIONS.map(l => ({ name: l.name, type: l.type, order: l.order }));
  }

  // 設定
  const config = {};
  if (configSheet) {
    const cfgValues = configSheet.getDataRange().getValues();
    cfgValues.shift();
    cfgValues.forEach(r => { if (r[0]) config[r[0]] = r[1]; });
  }

  return { items, categories, locations, config };
}

// ====== 履歴取得 ======
function getHistory(limit) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_HISTORY);
  const values = sheet.getDataRange().getValues();
  values.shift();
  const rows = values
    .filter(r => r[0])
    .map(r => ({
      datetime: r[0] ? Utilities.formatDate(new Date(r[0]), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss') : '',
      operator: r[1],
      location: r[2],
      code: r[3],
      name: r[4],
      type: r[5],
      qty: Number(r[6]) || 0,
      stockAfter: Number(r[7]) || 0,
      note: r[8] || ''
    }))
    .reverse()
    .slice(0, limit);
  return rows;
}

// ====== 入出庫処理 ======
function doInOut(params) {
  const { code, location, type, qty, operator, note } = params;
  if (!code || !location || !type || !qty) throw new Error('引数不足: code, location, type, qty が必要');
  const loc = findLocation(location);
  if (!loc) throw new Error('未知の拠点: ' + location);
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

    const currentStock = safeStockNum(item[loc.col - 1]);
    const newStock = type === '入庫' ? currentStock + qtyNum : currentStock - qtyNum;
    if (newStock < 0) throw new Error(loc.name + ' の在庫が足りません(現在 ' + currentStock + ')');

    sheet.getRange(targetRow, loc.col).setValue(newStock);
    sheet.getRange(targetRow, COL_UPDATED).setValue(new Date());

    logHistory({
      operator: operator || '不明',
      location: loc.name,
      code: code,
      name: item[COL_NAME - 1],
      type: type,
      qty: qtyNum,
      stockAfter: newStock,
      note: note || ''
    });

    return { code, name: item[COL_NAME - 1], location: loc.name, stockBefore: currentStock, stockAfter: newStock };
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
    rec.location,
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
  const stocks = p.stocks || {};
  const row = new Array(STOCK_COL_COUNT).fill('');
  row[COL_CODE - 1] = p.code;
  row[COL_NAME - 1] = p.name;
  row[COL_CATEGORY - 1] = p.category;
  row[COL_UNIT - 1] = p.unit;
  LOCATIONS.forEach(loc => {
    row[loc.col - 1] = Number(stocks[loc.name]) || 0;
  });
  // 合計セルは appendRow 後に数式を入れるためここでは空のまま
  row[COL_MIN_STOCK - 1] = Number(p.minStock) || 0;
  row[COL_EXPIRY - 1] = p.expiry || '';
  row[COL_NOTE - 1] = p.note || '';
  row[COL_UPDATED - 1] = new Date();
  sheet.appendRow(row);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, COL_TOTAL).setFormula(`=SUM(E${lastRow}:H${lastRow})`);
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

// シートを削除して新規作成(書式までリセットするため)
function recreateSheet(ss, name) {
  const old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  return ss.insertSheet(name);
}

// ====== 初期化(初回のみ手動実行) ======
function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 何もシートが残らない状態を避けるため、ダミーシートを先に作る(削除順序対策)
  let dummy = null;
  if (ss.getSheets().length <= 5) {
    dummy = ss.insertSheet('__tmp_init__');
  }

  // 在庫マスタ(シート削除→再作成で書式を完全リセット)
  const s = recreateSheet(ss, SHEET_STOCK);
  s.getRange(1, 1, 1, STOCK_COL_COUNT).setValues([[
    '商品コード', '商品名', 'カテゴリー', '単位',
    '塩山店', '百間店', '河口湖店', 'MS',
    '合計', '最低在庫', '賞味期限', '備考', '更新日時'
  ]]);
  s.setFrozenRows(1);
  s.setFrozenColumns(2); // 商品コード・商品名を固定

  // ヘッダー書式
  const header = s.getRange(1, 1, 1, STOCK_COL_COUNT);
  header.setBackground('#1f2937')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
  s.setRowHeight(1, 36);

  // 数値・日付書式(全行)
  const maxRow = s.getMaxRows();
  s.getRange(2, COL_STOCK_SHIOYAMA, maxRow - 1, 4).setNumberFormat('#,##0.##').setHorizontalAlignment('center');
  s.getRange(2, COL_TOTAL, maxRow - 1, 1).setNumberFormat('#,##0.##').setHorizontalAlignment('center').setFontWeight('bold');
  s.getRange(2, COL_MIN_STOCK, maxRow - 1, 1).setNumberFormat('#,##0.##').setHorizontalAlignment('center');
  s.getRange(2, COL_EXPIRY, maxRow - 1, 1).setNumberFormat('yyyy/mm/dd').setHorizontalAlignment('center');
  s.getRange(2, COL_UPDATED, maxRow - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm').setHorizontalAlignment('center').setFontColor('#6b7280');
  s.getRange(2, COL_UNIT, maxRow - 1, 1).setHorizontalAlignment('center');
  s.getRange(2, COL_CODE, maxRow - 1, 1).setFontFamily('Roboto Mono');

  // 合計列: 淡黄色のハイライト
  s.getRange(2, COL_TOTAL, maxRow - 1, 1).setBackground('#fefce8');

  // 列幅
  s.setColumnWidth(COL_CODE, 80);
  s.setColumnWidth(COL_NAME, 180);
  s.setColumnWidth(COL_CATEGORY, 90);
  s.setColumnWidth(COL_UNIT, 60);
  [COL_STOCK_SHIOYAMA, COL_STOCK_HYAKKEN, COL_STOCK_KAWAGUCHIKO, COL_STOCK_MS].forEach(c => s.setColumnWidth(c, 75));
  s.setColumnWidth(COL_TOTAL, 80);
  s.setColumnWidth(COL_MIN_STOCK, 80);
  s.setColumnWidth(COL_EXPIRY, 100);
  s.setColumnWidth(COL_NOTE, 140);
  s.setColumnWidth(COL_UPDATED, 130);

  // 条件付き書式: 合計 ≦ 最低在庫 → 行を薄赤背景
  const dangerRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($J2>0, $I2<=$J2)')
    .setBackground('#fee2e2')
    .setRanges([s.getRange(2, 1, maxRow - 1, STOCK_COL_COUNT)])
    .build();
  // 注意レベル: 合計 ≦ 最低在庫 × 1.5
  const warnRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($J2>0, $I2<=$J2*1.5)')
    .setBackground('#fef3c7')
    .setRanges([s.getRange(2, 1, maxRow - 1, STOCK_COL_COUNT)])
    .build();
  s.setConditionalFormatRules([dangerRule, warnRule]);

  // 共通ヘッダー書式関数
  const styleHeader = (sheet, cols) => {
    const range = sheet.getRange(1, 1, 1, cols);
    range.setBackground('#1f2937')
         .setFontColor('#ffffff')
         .setFontWeight('bold')
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 34);
    sheet.setFrozenRows(1);
  };

  // カテゴリー
  const c = recreateSheet(ss, SHEET_CATEGORY);
  c.getRange(1, 1, 1, 3).setValues([['カテゴリー名', '表示順', 'アイコン']]);
  c.getRange(2, 1, 6, 3).setValues([
    ['香辛料', 1, '🌶️'],
    ['野菜', 2, '🥕'],
    ['肉類', 3, '🍖'],
    ['乳製品', 4, '🥛'],
    ['調味料', 5, '🧂'],
    ['包材', 6, '📦']
  ]);
  styleHeader(c, 3);
  c.setColumnWidth(1, 140); c.setColumnWidth(2, 80); c.setColumnWidth(3, 80);
  c.getRange(2, 2, c.getMaxRows() - 1, 1).setHorizontalAlignment('center');
  c.getRange(2, 3, c.getMaxRows() - 1, 1).setHorizontalAlignment('center');

  // 拠点
  const l = recreateSheet(ss, SHEET_LOCATION);
  l.getRange(1, 1, 1, 3).setValues([['拠点名', '種別', '表示順']]);
  l.getRange(2, 1, LOCATIONS.length, 3).setValues(
    LOCATIONS.map(loc => [loc.name, loc.type, loc.order])
  );
  styleHeader(l, 3);
  l.setColumnWidth(1, 140); l.setColumnWidth(2, 80); l.setColumnWidth(3, 80);
  l.getRange(2, 2, l.getMaxRows() - 1, 2).setHorizontalAlignment('center');

  // 履歴
  const h = recreateSheet(ss, SHEET_HISTORY);
  h.getRange(1, 1, 1, 9).setValues([[
    '日時', '操作者', '拠点', '商品コード', '商品名', '区分', '数量', '操作後在庫', '備考'
  ]]);
  styleHeader(h, 9);
  h.getRange(2, 1, h.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss').setHorizontalAlignment('center').setFontColor('#6b7280');
  h.getRange(2, 2, h.getMaxRows() - 1, 1).setHorizontalAlignment('center');
  h.getRange(2, 3, h.getMaxRows() - 1, 1).setHorizontalAlignment('center');
  h.getRange(2, 4, h.getMaxRows() - 1, 1).setFontFamily('Roboto Mono');
  h.getRange(2, 6, h.getMaxRows() - 1, 1).setHorizontalAlignment('center').setFontWeight('bold');
  h.getRange(2, 7, h.getMaxRows() - 1, 2).setNumberFormat('#,##0.##').setHorizontalAlignment('center');
  h.setColumnWidth(1, 150); h.setColumnWidth(2, 80); h.setColumnWidth(3, 90);
  h.setColumnWidth(4, 90); h.setColumnWidth(5, 160); h.setColumnWidth(6, 60);
  h.setColumnWidth(7, 70); h.setColumnWidth(8, 90); h.setColumnWidth(9, 140);

  // 区分列の条件付き書式(入庫=緑、出庫=赤)
  const inRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('入庫')
    .setBackground('#dcfce7').setFontColor('#15803d')
    .setRanges([h.getRange(2, 6, h.getMaxRows() - 1, 1)])
    .build();
  const outRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('出庫')
    .setBackground('#fee2e2').setFontColor('#b91c1c')
    .setRanges([h.getRange(2, 6, h.getMaxRows() - 1, 1)])
    .build();
  h.setConditionalFormatRules([inRule, outRule]);

  // 設定
  const cfg = recreateSheet(ss, SHEET_CONFIG);
  cfg.getRange(1, 1, 1, 2).setValues([['キー', '値']]);
  cfg.getRange(2, 1, 2, 2).setValues([
    ['alert_threshold_ratio', 1.0],
    ['default_operator', '田中']
  ]);
  styleHeader(cfg, 2);
  cfg.setColumnWidth(1, 200); cfg.setColumnWidth(2, 140);

  // ダミーシートがあれば削除
  if (dummy) ss.deleteSheet(dummy);

  // サンプル商品(拠点別在庫付き)
  // 列構造: [code, name, category, unit, 塩山店, 百間店, 河口湖店, MS, 合計(数式), 最低在庫, 賞味期限, 備考, 更新日時]
  const now = new Date();
  const sampleData = [
    ['SP001', 'クミンパウダー',     '香辛料', 'g',  100,  100,  100,  200, '', 200,    '2026-12-31', '',       now],
    ['SP002', 'ターメリック',       '香辛料', 'g',   50,   50,   50,  150, '', 200,    '2026-12-31', '',       now],
    ['SP003', 'コリアンダー',       '香辛料', 'g',   80,   80,   80,  160, '', 200,    '2026-12-31', '',       now],
    ['VG001', '玉ねぎ',             '野菜',   '個',  15,   10,   10,   15, '',  20,    '2026-05-10', '北海道産', now],
    ['VG002', 'にんにく',           '野菜',   'kg', 0.5,  0.5,  0.5,  0.5, '',   1,    '2026-05-15', '',       now],
    ['VG003', '生姜',               '野菜',   'kg', 0.5,  0.3,  0.2,  0.5, '', 0.5,    '2026-05-12', '',       now],
    ['MT001', '鶏もも肉',           '肉類',   'kg',   3,    2,    2,    3, '',   5,    '2026-04-28', '冷蔵',   now],
    ['DR001', 'ヨーグルト',         '乳製品', 'g',  500,  500,  500,  500, '', 1000,   '2026-05-05', '',       now],
    ['SC001', 'トマト缶',           '調味料', '缶',   6,    6,    6,    6, '',  12,    '2027-06-30', '400g缶', now],
    ['PK001', '使い捨て容器500ml', '包材',   '個',  25,   25,   25,   25, '',  50,    '',           '',       now]
  ];
  s.getRange(2, 1, sampleData.length, STOCK_COL_COUNT).setValues(sampleData);
  // 合計列に SUM 数式を入れる
  const totalFormulas = sampleData.map((_, i) => [`=SUM(E${i + 2}:H${i + 2})`]);
  s.getRange(2, COL_TOTAL, totalFormulas.length, 1).setFormulas(totalFormulas);

  return { message: '初期化完了。合計列・条件付き書式を含む拠点別在庫マスタとサンプルデータを投入しました。' };
}
