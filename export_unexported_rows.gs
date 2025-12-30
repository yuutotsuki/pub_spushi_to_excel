/**
 * 設定シートで指定した対象シートのうち、L列（エクスポート）が空の行だけを
 * L列を除外したExcelに出力し、成功後にL列へ「済」を書き戻す。
 */
function exportUnexportedRowsToXlsx() {
  const config = {
    exportMarker: '済',              // 出力後に書き戻す値
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings_(ss);
  if (settings.sheetName === SETTINGS_SHEET_NAME) {
    throw new Error(`Target sheet name cannot be "${SETTINGS_SHEET_NAME}".`);
  }

  const sheet = ss.getSheetByName(settings.sheetName);
  if (!sheet) throw new Error(`Sheet "${settings.sheetName}" not found.`);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    Logger.log('No data rows to export.');
    return;
  }

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0];
  const exportColIdx = Number(settings.exportColumn) - 1;
  const markerColIdx = resolveMarkerColumnIndex_(header, exportColIdx);
  const exportLabelIndices = getExportLabelIndices_(header);
  if (exportLabelIndices.length > 1) {
    throw new Error('「エクスポート」ヘッダが複数あります。1つだけにしてください。');
  }
  if (exportLabelIndices.length === 1 && exportLabelIndices[0] !== markerColIdx) {
    throw new Error('設定のエクスポート列番号と、シートの「エクスポート」列が一致していません。');
  }
  const exportRemoveIdx = markerColIdx;

  const rowsToExport = [];
  const rowsToMark = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    // データの無い行はスキップ（全列空 or エクスポート列以外が空）
    if (isEmptyRow(row, markerColIdx)) continue;

    const marker = row[markerColIdx];
    if (marker === '' || marker === null) {
      rowsToExport.push(removeExportColumn(row, exportRemoveIdx));
      rowsToMark.push(i + 1); // シート上の行番号（ヘッダー込み）
    }
  }

  if (rowsToExport.length === 0) {
    Logger.log('No unexported rows found.');
    return;
  }

  const output = [
    removeExportColumn(header, exportRemoveIdx),
    ...rowsToExport,
  ];

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const baseName = settings.outputBaseName || 'export';
  const fileName = `${baseName}_${timestamp}.xlsx`;

  let tempFileId = '';
  try {
    // 一時スプレッドシートを作成し、L列を除外したデータを書き込み。
    const tempSs = SpreadsheetApp.create(fileName);
    tempFileId = tempSs.getId();
    const tempSheet = tempSs.getActiveSheet();
    tempSheet.clearContents();
    tempSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
    SpreadsheetApp.flush(); // ensure data is written before export

    // .xlsx としてエクスポート（Drive API v3 を直接叩いて alt=media を指定）。
    const exportBlob = fetchExportAsXlsx(tempFileId).setName(fileName);

    const parentFolder = settings.outputFolderId
      ? getFolderOrThrow_(settings.outputFolderId)
      : getParentFolder(ss.getId());
    parentFolder.createFile(exportBlob);

    // エクスポート済み行にマーカーを書き戻し。
    const markerColumn = columnToLetter_(markerColIdx + 1);
    sheet.getRangeList(rowsToMark.map((row) => `${markerColumn}${row}`)).setValue(config.exportMarker);

    Logger.log(`Exported ${rowsToMark.length} rows -> ${exportBlob.getName()}`);
  } finally {
    // 一時シートを削除（失敗時もクリーンアップ）。
    if (tempFileId) DriveApp.getFileById(tempFileId).setTrashed(true);
  }
}

function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSettingsSheet_(ss);
  SpreadsheetApp.getUi()
    .createMenu('エクスポート')
    .addItem('Excelに出力', 'exportUnexportedRowsToXlsx')
    .addToUi();
}

function removeExportColumn(row, exportColIdx) {
  const copy = row.slice();
  copy.splice(exportColIdx, 1);
  return copy;
}

function isEmptyRow(row, exportColIdx) {
  // エクスポート列以外に何も入っていなければ空行として扱う
  for (let i = 0; i < row.length; i++) {
    if (i === exportColIdx) continue;
    const v = row[i];
    if (v !== '' && v !== null) return false;
  }
  return true;
}

function getParentFolder(fileId) {
  const parents = DriveApp.getFileById(fileId).getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}
const SETTINGS_SHEET_NAME = '設定';
const SETTINGS_DEFAULTS = {
  sheetName: 'test',
  exportColumn: 12,
  outputBaseName: 'test',
  outputFolderId: '',
};
const SETTINGS_LABELS = {
  sheetName: '対象シート名',
  exportColumn: 'エクスポート列番号',
  outputBaseName: '出力ファイル名ベース',
  outputFolderId: '出力先フォルダID（空なら同じフォルダ）',
};

function getSettings_(ss) {
  const sheet = ensureSettingsSheet_(ss);
  const lastRow = Math.max(2, sheet.getLastRow());
  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = {};
  rows.forEach(([label, value]) => {
    if (!label) return;
    map[normalizeLabel_(label)] = value;
  });

  const sheetName = map[normalizeLabel_(SETTINGS_LABELS.sheetName)] || SETTINGS_DEFAULTS.sheetName;
  const exportColumn = parseColumnNumber_(map[normalizeLabel_(SETTINGS_LABELS.exportColumn)]) ||
    SETTINGS_DEFAULTS.exportColumn;
  const outputBaseName = map[normalizeLabel_(SETTINGS_LABELS.outputBaseName)] || SETTINGS_DEFAULTS.outputBaseName;
  const rawFolder =
    map[normalizeLabel_(SETTINGS_LABELS.outputFolderId)] ||
    map[normalizeLabel_('出力先フォルダID')];
  const outputFolderId = normalizeFolderId_(rawFolder) || SETTINGS_DEFAULTS.outputFolderId;
  return { sheetName, exportColumn, outputBaseName, outputFolderId };
}

function ensureSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  sheet.getRange('A1:B1').setValues([['設定項目', '値']]);
  sheet.getRange('A2:B5').setValues([
    [SETTINGS_LABELS.sheetName, SETTINGS_DEFAULTS.sheetName],
    [SETTINGS_LABELS.exportColumn, SETTINGS_DEFAULTS.exportColumn],
    [SETTINGS_LABELS.outputBaseName, SETTINGS_DEFAULTS.outputBaseName],
    [SETTINGS_LABELS.outputFolderId, SETTINGS_DEFAULTS.outputFolderId],
  ]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);
  return sheet;
}

function normalizeFolderId_(value) {
  if (value === null || typeof value === 'undefined') return '';
  const text = String(value).trim();
  if (!text) return '';
  const match = text.match(/folders\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  return text;
}

function parseColumnNumber_(value) {
  if (value === null || typeof value === 'undefined') return NaN;
  const text = String(value).trim();
  if (!text) return NaN;
  const digits = text.replace(/[^\d]/g, '');
  const num = parseInt(digits, 10);
  return Number.isFinite(num) && num > 0 ? num : NaN;
}

function columnToLetter_(col) {
  let n = col;
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

function normalizeLabel_(value) {
  return String(value)
    .trim()
    .replace(/\s+/g, '')
    .replace(/（.*?）/g, '');
}

function resolveMarkerColumnIndex_(headerRow, exportColIdx) {
  const lastIdx = headerRow.length - 1;
  const idxInRange = Number.isInteger(exportColIdx) && exportColIdx >= 0 && exportColIdx <= lastIdx;
  if (idxInRange) return exportColIdx;

  const normalizedHeaders = headerRow.map((v) => normalizeLabel_(v || ''));
  const labelIdx = normalizedHeaders.indexOf('エクスポート');
  if (labelIdx !== -1) return labelIdx;

  throw new Error('Invalid export column. Check "エクスポート列番号" in 設定シート.');
}

function getExportLabelIndices_(headerRow) {
  const normalizedHeaders = headerRow.map((v) => normalizeLabel_(v || ''));
  const indices = [];
  for (let i = 0; i < normalizedHeaders.length; i++) {
    if (normalizedHeaders[i] === 'エクスポート') indices.push(i);
  }
  return indices;
}


function getFolderOrThrow_(folderId) {
  try {
    return DriveApp.getFolderById(folderId);
  } catch (err) {
    throw new Error(`Invalid OUTPUT_FOLDER_ID: ${folderId}`);
  }
}

/**
 * Drive API v3 で alt=media を指定してエクスポートを取得。
 */
function fetchExportAsXlsx(fileId) {
  const url = `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=${encodeURIComponent('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}&alt=media`;
  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  });

  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Export failed (${code}): ${resp.getContentText()}`);
  }
  return resp.getBlob();
}
