const BOOK_SHEET_NAME = '書籍データ';
const BOOK_FIELDS = [
  'ID', '転記', '書名', '著者', '出版社', 'ISBN', 'ISC', '発行年',
  '本体価格', '本体税率', '税込価格', '仕入時の状態', '仕入日', '仕入先', '束親',
  '仕入数', '仕入値', '仕入値（税抜）', '掛率', '仕入送料', '仕入送料（税抜）',
  '仕入合計', '売上数', '在庫数', '在庫金額', '販売場所', '登録日',
  '販売価格', '販売価格（税抜）', '販売税率', '販売送料', '販売送料（税抜）',
  '販売送料税率', '売上合計', '発送手段', '発送費用', '決済手数料名',
  '決済手数料', '利益', '利益率', '売上日', '入金日', 'ステータス',
  '備考', '証憑', '記入日'
];

function doGet(event) {
  if (event.parameter.action === 'app') {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('書籍・在庫管理 Webシステム');
  }

  const callback = String(event.parameter.callback || '');
  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput('Invalid callback');
  }

  const requestId = String(event.parameter.requestId || '').trim();
  if (requestId) {
    const cachedResult = CacheService.getScriptCache().get(syncCacheKey_(requestId));
    return jsonpResponse_(callback, cachedResult
      ? JSON.parse(cachedResult)
      : { protocol: 2, pending: true });
  }

  return jsonpResponse_(callback, { protocol: 2, mode: 'google.script.run' });
}

function doPost(event) {
  let requestId = String(event.parameter.requestId || '').trim();
  try {
    const request = JSON.parse(event.parameter.payload || '{}');
    requestId = requestId || String(request.requestId || '').trim();
    const books = synchronizeBooksFromClient(request);
    if (requestId) {
      putSyncResult_(requestId, { protocol: 2, ok: true, books: books });
    }
    return ContentService.createTextOutput('');
  } catch (error) {
    console.error(error);
    if (requestId) {
      putSyncResult_(requestId, {
        protocol: 2,
        ok: false,
        error: error.message || String(error)
      });
    }
    return ContentService.createTextOutput('');
  }
}

function syncCacheKey_(requestId) {
  return `book-sync-result:${requestId}`;
}

function putSyncResult_(requestId, result) {
  CacheService.getScriptCache().put(syncCacheKey_(requestId), JSON.stringify(result), 21600);
}

function synchronizeBooksFromClient(request) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return synchronizeBooks_(request);
  } finally {
    lock.releaseLock();
  }
}

function synchronizeBooks_(request) {
  if (!request || typeof request !== 'object') {
    throw new Error('同期データの形式が正しくありません。');
  }

  const localBooks = recordsById_(request.books, 'アプリ');
  const baselineBooks = recordsById_(request.baseline, '同期履歴');
  const sheet = getBookSheet_();
  const remoteBooks = readSheetBooks_(sheet);
  const mergedBooks = mergeBooks_(localBooks, remoteBooks, baselineBooks);

  writeSheetBooks_(sheet, mergedBooks);
  return Array.from(mergedBooks.values());
}

function getBookSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('このApps Scriptを対象スプレッドシートに紐づけてください。');
  }

  let sheet = spreadsheet.getSheetByName(BOOK_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.getSheets().find(candidate => {
      if (candidate.getLastRow() === 0 || candidate.getLastColumn() < BOOK_FIELDS.length) return false;
      const headers = candidate.getRange(1, 1, 1, BOOK_FIELDS.length).getDisplayValues()[0];
      return BOOK_FIELDS.every((field, index) => headers[index] === field);
    });
  }
  if (!sheet) sheet = spreadsheet.insertSheet(BOOK_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, BOOK_FIELDS.length).setValues([BOOK_FIELDS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, BOOK_FIELDS.length).setFontWeight('bold');
  } else {
    const headers = sheet.getRange(1, 1, 1, BOOK_FIELDS.length).getDisplayValues()[0];
    if (BOOK_FIELDS.some((field, index) => headers[index] !== field)) {
      throw new Error(`「${BOOK_SHEET_NAME}」シートの1行目が想定した項目名と一致しません。既存データを保護するため同期を中止しました。`);
    }
  }

  return sheet;
}

function readSheetBooks_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Map();

  const values = sheet.getRange(2, 1, lastRow - 1, BOOK_FIELDS.length).getValues();
  const books = [];
  values.forEach(row => {
    if (row[0] === '' || row[0] === null) return;
    const book = {};
    BOOK_FIELDS.forEach((field, index) => {
      const value = row[index];
      book[field] = value instanceof Date
        ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : value;
    });
    books.push(book);
  });
  return recordsById_(books, 'スプレッドシート');
}

function recordsById_(records, sourceName) {
  if (!Array.isArray(records)) {
    throw new Error(`${sourceName}の書籍データが配列ではありません。`);
  }

  const result = new Map();
  records.forEach(source => {
    const book = {};
    BOOK_FIELDS.forEach(field => {
      book[field] = source[field] === undefined || source[field] === null ? '' : source[field];
    });
    const id = String(book.ID).trim();
    if (!id) throw new Error(`${sourceName}にIDが未入力の書籍があります。`);
    if (result.has(id)) throw new Error(`${sourceName}にID「${id}」が重複しています。`);
    book.ID = id;
    result.set(id, book);
  });
  return result;
}

function mergeBooks_(localBooks, remoteBooks, baselineBooks) {
  const ids = new Set([
    ...localBooks.keys(),
    ...remoteBooks.keys(),
    ...baselineBooks.keys()
  ]);
  const result = new Map();

  ids.forEach(id => {
    const local = localBooks.get(id);
    const remote = remoteBooks.get(id);
    const baseline = baselineBooks.get(id);

    if (!baseline) {
      const newBook = remote || local;
      if (newBook) result.set(id, newBook);
      return;
    }

    if (!local) {
      if (remote && !sameBook_(remote, baseline)) result.set(id, remote);
      return;
    }

    if (!remote) {
      return;
    }

    result.set(id, sameBook_(remote, baseline) ? local : remote);
  });

  return result;
}

function sameBook_(left, right) {
  return BOOK_FIELDS.every(field => String(left[field] ?? '') === String(right[field] ?? ''));
}

function writeSheetBooks_(sheet, books) {
  const lastRow = sheet.getLastRow();
  const existingRows = lastRow < 2
    ? []
    : sheet.getRange(2, 1, lastRow - 1, BOOK_FIELDS.length).getValues();
  const desiredIds = new Set(books.keys());
  const rowsToDelete = [];
  const existingRowById = new Map();

  existingRows.forEach((row, index) => {
    const id = String(row[0] ?? '').trim();
    if (!id) return;
    if (existingRowById.has(id)) throw new Error(`シート内でID「${id}」が重複しています。`);
    if (!desiredIds.has(id)) rowsToDelete.push(index + 2);
    else existingRowById.set(id, index + 2);
  });

  rowsToDelete.sort((left, right) => right - left).forEach(row => sheet.deleteRow(row));

  books.forEach((book, id) => {
    const values = BOOK_FIELDS.map(field => book[field] === undefined ? '' : book[field]);
    const row = existingRowById.get(id);
    if (row) {
      const removedAbove = rowsToDelete.filter(deletedRow => deletedRow < row).length;
      sheet.getRange(row - removedAbove, 1, 1, BOOK_FIELDS.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
  });
}

function jsonpResponse_(callback, message) {
  const encoded = JSON.stringify(message)
    .replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
  return ContentService.createTextOutput(`${callback}(${encoded});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
