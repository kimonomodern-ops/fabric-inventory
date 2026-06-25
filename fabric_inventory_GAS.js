/**
 * 生地在庫管理 — Google Apps Script (GAS) バックエンド
 * KIMONO MODERN 社内向け
 *
 * セットアップ手順:
 *   1. Googleスプレッドシートを新規作成
 *   2. URLから SPREADSHEET_ID をコピーして下記に貼り付け
 *   3. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *      アクセス権: 「全員」→ デプロイ → URLをコピー
 */

// ============================================================
// ★ ここにスプレッドシートIDを貼り付けてください
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
// ============================================================

const SHEET_INVENTORY = '在庫マスタ';
const SHEET_HISTORY   = '入出庫履歴';

// 在庫マスタのヘッダー列順
const INV_HEADERS = [
  '品番', '品名', '色', '規格', 'カテゴリ', '管理単位',
  '在庫数', '単価', '仕入先', '写真URL', '商品紐付けID',
  '弊社品番', '弊社商品名', '備考', '登録日', '更新日'
];

// 入出庫履歴のヘッダー列順
const HIST_HEADERS = [
  '日時', '品番', '品名', '種別', '増減数', '更新後在庫', '担当者', '備考'
];

// ============================================================
// GETリクエスト処理
// ============================================================
function doGet(e) {
  const action = e.parameter.action || '';
  let result;
  try {
    switch (action) {
      case 'getInventory':
        result = getInventory();
        break;
      case 'getHistory':
        result = getHistory(Number(e.parameter.limit) || 500);
        break;
      case 'getStats':
        result = getStats();
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return buildResponse(result);
}

// ============================================================
// POSTリクエスト処理
// ============================================================
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return buildResponse({ success: false, error: 'Invalid JSON: ' + err.message });
  }

  const action = body.action || '';
  let result;
  try {
    switch (action) {
      case 'addItem':
        result = addItem(body);
        break;
      case 'updateItem':
        result = updateItem(body);
        break;
      case 'deleteItem':
        result = deleteItem(body._row);
        break;
      case 'updateStock':
        result = updateStock(body);
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }
  return buildResponse(result);
}

// ============================================================
// レスポンスビルダー（CORS対応）
// ============================================================
function buildResponse(data) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============================================================
// スプレッドシート初期化（シートが無ければ作成）
// ============================================================
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    // ヘッダー行のスタイル設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f0ece4');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getInvSheet()  { return getOrCreateSheet(SHEET_INVENTORY, INV_HEADERS); }
function getHistSheet() { return getOrCreateSheet(SHEET_HISTORY,   HIST_HEADERS); }

// ============================================================
// 在庫マスタ取得
// ============================================================
function getInventory() {
  const sheet = getInvSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { items: [] };

  const headers = data[0];
  const items = data.slice(1).map((row, i) => {
    const obj = { _row: i + 2 }; // 実際の行番号（1-indexed, ヘッダー除く）
    headers.forEach((h, j) => { obj[h] = row[j] !== undefined ? row[j] : ''; });
    // 数値型を確実に変換
    obj['在庫数'] = Number(obj['在庫数']) || 0;
    obj['単価']   = Number(obj['単価'])   || 0;
    // 日付をISO文字列に変換
    if (obj['登録日'] instanceof Date) obj['登録日'] = obj['登録日'].toISOString();
    if (obj['更新日'] instanceof Date) obj['更新日'] = obj['更新日'].toISOString();
    return obj;
  });
  return { items };
}

// ============================================================
// 在庫マスタ: 生地追加
// ============================================================
function addItem(data) {
  const sheet = getInvSheet();
  const now = new Date();

  // 品番重複チェック
  if (data['品番']) {
    const existing = findRowByCode(sheet, data['品番']);
    if (existing > 0) {
      // 既存行を更新
      return updateItemAtRow(sheet, existing, data);
    }
  }

  const row = INV_HEADERS.map(h => {
    if (h === '登録日') return now;
    if (h === '更新日') return now;
    if (h === '在庫数') return Number(data[h]) || 0;
    if (h === '単価')   return Number(data[h]) || 0;
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(row);
  return { success: true };
}

// ============================================================
// 在庫マスタ: 生地更新
// ============================================================
function updateItem(data) {
  const sheet = getInvSheet();
  const rowNum = data._row || (data['品番'] ? findRowByCode(sheet, data['品番']) : -1);
  if (!rowNum || rowNum < 2) return { success: false, error: '対象行が見つかりません' };
  return updateItemAtRow(sheet, rowNum, data);
}

function updateItemAtRow(sheet, rowNum, data) {
  const now = new Date();
  const headers = sheet.getRange(1, 1, 1, INV_HEADERS.length).getValues()[0];
  headers.forEach((h, j) => {
    if (h === '登録日') return; // 登録日は変更しない
    if (h === '更新日') { sheet.getRange(rowNum, j + 1).setValue(now); return; }
    if (data[h] !== undefined) {
      const val = (h === '在庫数' || h === '単価') ? (Number(data[h]) || 0) : data[h];
      sheet.getRange(rowNum, j + 1).setValue(val);
    }
  });
  return { success: true };
}

// ============================================================
// 在庫マスタ: 生地削除
// ============================================================
function deleteItem(rowNum) {
  if (!rowNum || rowNum < 2) return { success: false, error: '対象行が見つかりません' };
  const sheet = getInvSheet();
  sheet.deleteRow(rowNum);
  return { success: true };
}

// ============================================================
// 在庫マスタ: 在庫数増減 + 履歴記録
// ============================================================
function updateStock(data) {
  const sheet = getInvSheet();
  const rowNum = data._row;
  if (!rowNum || rowNum < 2) return { success: false, error: '対象行が見つかりません' };

  // 在庫数列のインデックスを取得
  const headers = sheet.getRange(1, 1, 1, INV_HEADERS.length).getValues()[0];
  const stockCol = headers.indexOf('在庫数') + 1;
  const nameCol  = headers.indexOf('品名')   + 1;
  const codeCol  = headers.indexOf('品番')   + 1;
  const unitCol  = headers.indexOf('管理単位') + 1;
  const updCol   = headers.indexOf('更新日')  + 1;

  if (stockCol < 1) return { success: false, error: '在庫数列が見つかりません' };

  const currentStock = Number(sheet.getRange(rowNum, stockCol).getValue()) || 0;
  const delta = Number(data.delta) || 0;
  const newStock = currentStock + delta;

  if (newStock < 0) return { success: false, error: '在庫がマイナスになります（現在: ' + currentStock + '）' };

  const now = new Date();
  sheet.getRange(rowNum, stockCol).setValue(Math.round(newStock * 100) / 100);
  if (updCol > 0) sheet.getRange(rowNum, updCol).setValue(now);

  // 履歴記録
  const itemCode = sheet.getRange(rowNum, codeCol).getValue();
  const itemName = sheet.getRange(rowNum, nameCol).getValue();
  addHistory({
    日時: now,
    品番: itemCode,
    品名: itemName,
    種別: data.type || '入庫',
    増減数: delta,
    更新後在庫: Math.round(newStock * 100) / 100,
    担当者: data.operator || '',
    備考: data.memo || ''
  });

  return { success: true, newStock: Math.round(newStock * 100) / 100 };
}

// ============================================================
// 入出庫履歴: 追加
// ============================================================
function addHistory(data) {
  const sheet = getHistSheet();
  const row = HIST_HEADERS.map(h => data[h] !== undefined ? data[h] : '');
  sheet.appendRow(row);
}

// ============================================================
// 入出庫履歴: 取得
// ============================================================
function getHistory(limit) {
  const sheet = getHistSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { history: [] };

  const headers = data[0];
  let rows = data.slice(1).map((row, i) => {
    const obj = {};
    headers.forEach((h, j) => { obj[h] = row[j] !== undefined ? row[j] : ''; });
    if (obj['日時'] instanceof Date) obj['日時'] = obj['日時'].toISOString();
    obj['増減数'] = Number(obj['増減数']) || 0;
    obj['更新後在庫'] = Number(obj['更新後在庫']) || 0;
    return obj;
  });

  // 新しい順に並べ替え
  rows.sort((a, b) => new Date(b['日時']) - new Date(a['日時']));
  if (limit > 0) rows = rows.slice(0, limit);

  return { history: rows };
}

// ============================================================
// 統計情報取得
// ============================================================
function getStats() {
  const sheet = getInvSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { total: 0, value: 0, lowCount: 0, categories: 0 };

  const headers = data[0];
  const stockIdx = headers.indexOf('在庫数');
  const priceIdx = headers.indexOf('単価');
  const catIdx   = headers.indexOf('カテゴリ');

  let total = 0, value = 0, lowCount = 0;
  const cats = new Set();

  data.slice(1).forEach(row => {
    const stock = Number(row[stockIdx]) || 0;
    const price = Number(row[priceIdx]) || 0;
    const cat   = row[catIdx] || '';
    total++;
    value += stock * price;
    if (stock < 5) lowCount++;
    if (cat) cats.add(cat);
  });

  return { success: true, total, value, lowCount, categories: cats.size };
}

// ============================================================
// ユーティリティ: 品番で行番号を検索
// ============================================================
function findRowByCode(sheet, code) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const codeIdx = headers.indexOf('品番');
  if (codeIdx < 0) return -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][codeIdx]).trim() === String(code).trim()) return i + 1;
  }
  return -1;
}
