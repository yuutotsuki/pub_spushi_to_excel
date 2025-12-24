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

  const rowsToExport = [];
  const rowsToMark = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    // データの無い行はスキップ（全列空 or エクスポート列以外が空）
    if (isEmptyRow(row, exportColIdx)) continue;

    const marker = row[exportColIdx];
    if (marker === '' || marker === null) {
      rowsToExport.push(removeExportColumn(row, exportColIdx));
      rowsToMark.push(i + 1); // シート上の行番号（ヘッダー込み）
    }
  }

  if (rowsToExport.length === 0) {
    Logger.log('No unexported rows found.');
    return;
  }

  const output = [
    removeExportColumn(header, exportColIdx),
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
      ? DriveApp.getFolderById(settings.outputFolderId)
      : getParentFolder(ss.getId());
    parentFolder.createFile(exportBlob);

    // エクスポート済み行にマーカーを書き戻し。
    sheet.getRangeList(rowsToMark.map((row) => `L${row}`)).setValue(config.exportMarker);

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

function getSettings_(ss) {
  const sheet = ensureSettingsSheet_(ss);
  const values = sheet.getRange(2, 2, 4, 1).getValues().flat();
  const sheetName = values[0] || SETTINGS_DEFAULTS.sheetName;
  const exportColumn = Number(values[1]) || SETTINGS_DEFAULTS.exportColumn;
  const outputBaseName = values[2] || SETTINGS_DEFAULTS.outputBaseName;
  const outputFolderId = values[3] || SETTINGS_DEFAULTS.outputFolderId;
  return { sheetName, exportColumn, outputBaseName, outputFolderId };
}

function ensureSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  sheet.getRange('A1:B1').setValues([['設定項目', '値']]);
  sheet.getRange('A2:B5').setValues([
    ['対象シート名', SETTINGS_DEFAULTS.sheetName],
    ['エクスポート列番号', SETTINGS_DEFAULTS.exportColumn],
    ['出力ファイル名ベース', SETTINGS_DEFAULTS.outputBaseName],
    ['出力先フォルダID（空なら同じフォルダ）', SETTINGS_DEFAULTS.outputFolderId],
  ]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);
  return sheet;
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
