const HEADER_FONT_COLOR = '#FFFFFF';
const HEADER_BORDER_COLOR = '#666666';
const DATA_FONT_COLOR = '#DDDDDD';

/**
 * Applies dark-mode styling to the header row of the SEO URL Scraper sheet.
 * Freezes the header row, sets background, font styling, and border.
 * @throws {Error} If the spreadsheet or target sheet is unavailable.
 */
function applyHeaderStyles() {
  const globals = getGlobalSheetConstants();
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Unable to access the active spreadsheet.');
  const sheetName = globals.SHEET_NAME;
  if (typeof sheetName !== 'string' || !sheetName.trim()) {
    throw new Error('SHEET_NAME is not defined or invalid.');
  }
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
  
  sheet.setFrozenRows(1);
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    throw new Error(`Sheet "${sheetName}" has no columns.`);
  }
  
  const headerRange = sheet.getRange(1, 1, 1, lastColumn);
  headerRange
    .setBackground(globals.COLOR_HEADER_BG)
    .setFontColor(HEADER_FONT_COLOR)
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true, HEADER_BORDER_COLOR);
  
  sheet.setRowHeight(1, 40);
}

/**
 * Applies dark-mode styling to all data rows of the SEO URL Scraper sheet.
 * Alternates row background colors and sets font color.
 * @throws {Error} If there are no data rows or the target sheet is unavailable.
 */
function applyDataRowStyles() {
  const globals = getGlobalSheetConstants();
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Unable to access the active spreadsheet.');
  const sheetName = globals.SHEET_NAME;
  if (typeof sheetName !== 'string' || !sheetName.trim()) {
    throw new Error('SHEET_NAME is not defined or invalid.');
  }
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No data rows to style.');
  const lastColumn = sheet.getLastColumn();
  
  const numRows = lastRow - 1;
  const numCols = lastColumn;
  const backgrounds = [];
  const fontColors = [];
  
  for (let i = 0; i < numRows; i++) {
    const rowIndex = i + 2;
    const bgColor = (rowIndex % 2 === 0) ? globals.COLOR_ROW_EVEN : globals.COLOR_ROW_ODD;
    backgrounds.push(new Array(numCols).fill(bgColor));
    fontColors.push(new Array(numCols).fill(DATA_FONT_COLOR));
  }
  
  const dataRange = sheet.getRange(2, 1, numRows, numCols);
  dataRange
    .setBackgrounds(backgrounds)
    .setFontColors(fontColors);
}
