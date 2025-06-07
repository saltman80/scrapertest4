function highlightInvalidRows(rows) {
  if (!Array.isArray(rows)) {
    throw new Error('highlightInvalidRows: rows must be an array');
  }
  var globals = getGlobalSheetConstants();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(globals.SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + globals.SHEET_NAME + '" not found');

  var bgColor = globals.TOAST_ERROR_BG || '#C62828';
  rows.forEach(function(row) {
    if (typeof row !== 'number' || row < globals.DATA_START_ROW_INDEX) return;
    sheet.getRange(row, globals.COL_URL).setBackground(bgColor);
  });
}
