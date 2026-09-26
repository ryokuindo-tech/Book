const BOOK_SHEET_NAME = '書籍データ';
const BOOK_FIELDS = [
  'ID', '転記', '書名', '著者', '出版社', 'ISBN', 'ISC', '発行年',
  '本体価格', '本体税率', '税込価格', '仕入時の状態', '仕入日', '仕入先', '束親',
  '仕入数', '仕入値', '仕入値（税抜）', '掛率', '仕入送料', '仕入送料（税抜）',
  '仕入合計', '売上数', '在庫数', '在庫金額', '販売場所', '登録日',
  '販売価格', '販売価格（税抜）', '販売税率', '販売送料', '販売送料（税抜）',
  '販売送料税率', '売上合計', '売上合計（税抜）', '発送手段', '発送費用', '決済手数料名',
  '決済手数料', '利益', '利益率', '売上日', '入金日', 'ステータス',
  '備考', '証憑', '記入日'
];

const SYNC_RESULT_CACHE_PREFIX = 'book-sync-result:';
const SYNC_RESULT_CHUNK_LENGTH = 20000;
const SYNC_RESULT_CACHE_TTL_SECONDS = 600;

function doGet(event) {
  const parameters = event && event.parameter ? event.parameter : {};
  if (parameters.requestId) return getSyncResult_(parameters);

  const action = String(parameters.action || 'app');
  const isReport = action === 'report';
  const template = HtmlService.createTemplateFromFile(isReport ? 'Report' : 'Index');
  template.WEB_APP_URL = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle(isReport ? '損益計算書 | 書籍管理システム' : '書籍管理システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(event) {
  const parameters = event && event.parameter ? event.parameter : {};
  const payloadText = parameters.payload;
  if (!payloadText) throw new Error('同期データを受信できませんでした。');

  let request;
  try {
    request = JSON.parse(payloadText);
  } catch (error) {
    throw new Error('同期データの形式が正しくありません。');
  }

  const requestId = String(request.requestId || '');
  if (!/^[a-zA-Z0-9_-]{1,100}$/.test(requestId)) {
    throw new Error('同期リクエストIDが正しくありません。');
  }

  try {
    const books = synchronizeBooksFromClient(request);
    storeSyncResult_(requestId, { ok: true, books: books });
  } catch (error) {
    console.error('Googleスプレッドシート同期エラー:', error);
    storeSyncResult_(requestId, {
      ok: false,
      error: error && error.message ? error.message : '同期に失敗しました。'
    });
  }

  return ContentService.createTextOutput('同期リクエストを処理しました。');
}

function getSyncResult_(parameters) {
  const requestId = String(parameters.requestId || '');
  const callback = String(parameters.callback || '');
  if (!/^[a-zA-Z0-9_-]{1,100}$/.test(requestId)) {
    throw new Error('同期リクエストIDが正しくありません。');
  }
  if (!/^bookSyncResult_[a-zA-Z0-9_]+$/.test(callback)) {
    throw new Error('同期結果のコールバックが正しくありません。');
  }

  const cache = CacheService.getScriptCache();
  const result = readSyncResult_(cache, requestId);
  const json = JSON.stringify(result).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  return ContentService.createTextOutput(`${callback}(${json});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function storeSyncResult_(requestId, result) {
  const cache = CacheService.getScriptCache();
  const serialized = JSON.stringify(result);
  const chunks = [];
  for (let offset = 0; offset < serialized.length; offset += SYNC_RESULT_CHUNK_LENGTH) {
    chunks.push(serialized.slice(offset, offset + SYNC_RESULT_CHUNK_LENGTH));
  }

  const cacheKeys = chunks.map((chunk, index) => `${SYNC_RESULT_CACHE_PREFIX}${requestId}:${index}`);
  cache.putAll(
    Object.fromEntries(chunks.map((chunk, index) => [cacheKeys[index], chunk])),
    SYNC_RESULT_CACHE_TTL_SECONDS
  );
  cache.put(
    `${SYNC_RESULT_CACHE_PREFIX}${requestId}:meta`,
    JSON.stringify({ chunkCount: chunks.length }),
    SYNC_RESULT_CACHE_TTL_SECONDS
  );
}

function readSyncResult_(cache, requestId) {
  const prefix = `${SYNC_RESULT_CACHE_PREFIX}${requestId}:`;
  const metadata = cache.get(`${prefix}meta`);
  if (!metadata) return { pending: true };

  const { chunkCount } = JSON.parse(metadata);
  const chunkKeys = Array.from({ length: chunkCount }, (_, index) => `${prefix}${index}`);
  const chunks = cache.getAll(chunkKeys);
  if (chunkKeys.some(key => chunks[key] === undefined)) {
    return { ok: false, error: '同期結果の保存期間が切れました。もう一度同期してください。' };
  }

  return JSON.parse(chunkKeys.map(key => chunks[key]).join(''));
}

function uploadEvidencePdf(request) {
  return uploadEvidencePdf_(request);
}

function uploadEvidencePdf_(request) {
  const bookId = String(request.bookId || '').trim();
  const fileName = String(request.fileName || '');
  const base64 = String(request.base64 || '');
  const purchaseDate = String(request.purchaseDate || '');
  const supplier = sanitizeEvidenceFileNamePart_(request.supplier);
  const evidenceType = sanitizeEvidenceFileNamePart_(request.evidenceType);
  const purchaseTotal = Number(request.purchaseTotal);
  if (!bookId) throw new Error('書籍IDがないため、証憑を保存できません。');
  if (!/\.pdf$/i.test(fileName)) throw new Error('PDFファイルを選択してください。');
  if (!base64) throw new Error('PDFファイルの内容を読み取れませんでした。');
  if (!/^\d{8}$/.test(purchaseDate)) throw new Error('仕入日をYYYYMMDD形式で入力してください。');
  const purchaseYear = Number(purchaseDate.slice(0, 4));
  const purchaseMonth = Number(purchaseDate.slice(4, 6));
  const day = Number(purchaseDate.slice(6, 8));
  const parsedPurchaseDate = new Date(purchaseYear, purchaseMonth - 1, day);
  if (
    parsedPurchaseDate.getFullYear() !== purchaseYear ||
    parsedPurchaseDate.getMonth() !== purchaseMonth - 1 ||
    parsedPurchaseDate.getDate() !== day
  ) throw new Error('仕入日が正しくありません。');
  if (!supplier) throw new Error('仕入先を入力してください。');
  if (!evidenceType) throw new Error('証憑の種類を入力してください。');
  if (!Number.isFinite(purchaseTotal) || purchaseTotal < 0) throw new Error('仕入合計が正しくありません。');

  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > 10 * 1024 * 1024) {
    throw new Error('PDFは10MB以下のファイルを選択してください。');
  }
  const signature = Utilities.newBlob(bytes).getDataAsString().slice(0, 5);
  if (signature !== '%PDF-') throw new Error('選択したファイルは有効なPDFではありません。');

  const timeZone = Session.getScriptTimeZone();
  const now = new Date();
  const year = Utilities.formatDate(now, timeZone, 'yyyy');
  const month = Utilities.formatDate(now, timeZone, 'MM');
  const root = DriveApp.getRootFolder();
  const evidenceFolder = getOrCreateFolder_(root, '証憑');
  const yearFolder = getOrCreateFolder_(evidenceFolder, Utilities.formatDate(now, timeZone, 'yyyy'));
  const monthFolder = getOrCreateFolder_(yearFolder, month);
  const safeAmount = String(Math.round(purchaseTotal));
  const pdf = Utilities.newBlob(
    bytes,
    'application/pdf',
    `${purchaseDate}_${safeAmount}_${supplier}_${evidenceType}.pdf`
  );
  const file = monthFolder.createFile(pdf);

  return {
    fileId: file.getId(),
    fileName: file.getName()
  };
}

function sanitizeEvidenceFileNamePart_(value) {
  return String(value || '')
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '-')
    .replace(/[\s_]+/g, '-')
    .replace(/^[.-]+|[.-]+$/g, '')
    .slice(0, 60);
}

function authorizeEvidenceDriveAccess() {
  const authorizationFolder = DriveApp.createFolder(
    `.book-drive-authorization-${Utilities.getUuid()}`
  );
  authorizationFolder.setTrashed(true);
}

function getOrCreateFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
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

  mergedBooks.forEach(normalizeInventoryFields_);
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
      if (candidate.getLastRow() === 0) return false;
      return sheetHasHeaders_(candidate, BOOK_FIELDS) ||
        sheetHasHeaders_(candidate, BOOK_FIELDS.filter(field => field !== '売上合計（税抜）'));
    });
  }
  if (!sheet) sheet = spreadsheet.insertSheet(BOOK_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, BOOK_FIELDS.length).setValues([BOOK_FIELDS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, BOOK_FIELDS.length).setFontWeight('bold');
  } else {
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const legacyFields = BOOK_FIELDS.filter(field => field !== '売上合計（税抜）');
    if (
      headers.length === legacyFields.length &&
      legacyFields.every((field, index) => headers[index] === field)
    ) {
      const salesTotalColumn = BOOK_FIELDS.indexOf('売上合計') + 1;
      sheet.insertColumnAfter(salesTotalColumn);
      sheet.getRange(1, salesTotalColumn + 1).setValue('売上合計（税抜）').setFontWeight('bold');
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    }
    if (BOOK_FIELDS.some((field, index) => headers[index] !== field)) {
      throw new Error(`「${BOOK_SHEET_NAME}」シートの1行目が想定した項目名と一致しません。既存データを保護するため同期を中止しました。`);
    }
  }

  return sheet;
}

function sheetHasHeaders_(sheet, fields) {
  if (sheet.getLastColumn() < fields.length) return false;
  const headers = sheet.getRange(1, 1, 1, fields.length).getDisplayValues()[0];
  return fields.every((field, index) => headers[index] === field);
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
    if (book['売上合計（税抜）'] === '') {
      const salesPriceNet = book['販売価格（税抜）'] === ''
        ? taxExclusiveAmount_(book['販売価格'], book['販売税率'])
        : Number(book['販売価格（税抜）']) || 0;
      const salesShippingNet = book['販売送料（税抜）'] === ''
        ? taxExclusiveAmount_(book['販売送料'], book['販売送料税率'])
        : Number(book['販売送料（税抜）']) || 0;
      book['売上合計（税抜）'] = salesPriceNet + salesShippingNet;
    }
    normalizeInventoryFields_(book);
    const id = String(book.ID).trim();
    if (!id) throw new Error(`${sourceName}にIDが未入力の書籍があります。`);
    if (result.has(id)) throw new Error(`${sourceName}にID「${id}」が重複しています。`);
    book.ID = id;
    result.set(id, book);
  });
  return result;
}

function normalizeInventoryFields_(book) {
  const purchaseQty = Number(book['仕入数']) || 0;
  const salesQty = Number(book['売上数']) || 0;
  const stockQty = Math.max(0, purchaseQty - salesQty);
  const purchaseUnitCost = Number(book['仕入値']) || 0;
  book['在庫数'] = stockQty;
  book['在庫金額'] = stockQty * purchaseUnitCost;
}

function taxExclusiveAmount_(inclusiveAmount, taxRate) {
  const amount = Number(inclusiveAmount) || 0;
  const rate = Number(taxRate) || 0;
  return Math.round(amount / (1 + rate / 100));
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

    const mergedBook = { ...local };
    BOOK_FIELDS.forEach(field => {
      if (!sameValue_(remote[field], baseline[field])) {
        mergedBook[field] = remote[field];
      }
    });
    result.set(id, mergedBook);
  });

  return result;
}

function sameValue_(left, right) {
  return String(left ?? '') === String(right ?? '');
}

function sameBook_(left, right) {
  return BOOK_FIELDS.every(field => sameValue_(left[field], right[field]));
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
