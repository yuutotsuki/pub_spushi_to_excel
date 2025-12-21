/**
 * testシートのうち、L列（エクスポート）が空の行だけを
 * L列を除外したExcelに出力し、成功後にL列へ「済」を書き戻す。
 *
 * 事前準備:
 * - `config.gs` に対象スプレッドシートIDなどの設定を入れる。
 */
function exportUnexportedRowsToXlsx() {
  const config = {
    sheetName: 'test',               // 対象シート名（既定値）
    exportColumn: 12,                // L列（1始まり）
    exportMarker: '済',              // 出力後に書き戻す値
    outputBaseName: 'test',          // 出力ファイル名のベース（既定値）
    outputFolderId: '',              // 空なら元スプレッドシートと同じフォルダ
  };

  if (typeof CONFIG === 'undefined') {
    throw new Error('CONFIG is required. Copy config.sample.gs to config.gs and edit values.');
  }

  const spreadsheetId = getRequiredConfig('SPREADSHEET_ID');
  const sheetName = getConfigOrDefault('SHEET_NAME', config.sheetName);
  const exportColumn = getConfigOrDefault('EXPORT_COLUMN', config.exportColumn);
  const outputBaseName = getConfigOrDefault('OUTPUT_BASENAME', config.outputBaseName);
  const outputFolderId = getConfigOrDefault('OUTPUT_FOLDER_ID', config.outputFolderId);
  const ss = SpreadsheetApp.openById(spreadsheetId);

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    Logger.log('No data rows to export.');
    return;
  }

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0];
  const exportColIdx = Number(exportColumn) - 1;

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
  const baseName = outputBaseName || 'export';
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

    const parentFolder = outputFolderId
      ? DriveApp.getFolderById(outputFolderId)
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

function getRequiredConfig(key) {
  const value = CONFIG[key];
  if (!value) throw new Error(`${key} is required in CONFIG.`);
  return value;
}

function getConfigOrDefault(key, defaultValue) {
  const value = CONFIG[key];
  return value === null || value === '' || typeof value === 'undefined' ? defaultValue : value;
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
