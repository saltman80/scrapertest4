function getOrCreateScrapeSheet() {
  var globals = getGlobalSheetConstants();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet');
  }
  var sheet = ss.getSheetByName(globals.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(globals.SHEET_NAME);
    sheet.getRange(globals.HEADER_ROW_INDEX, 1, 1, globals.HEADER_LABELS.length)
         .setValues([globals.HEADER_LABELS]);
  }
  return sheet;
}

