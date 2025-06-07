const { SHEET_NAME, COLOR_HEADER_BG, COLOR_ROW_EVEN, COLOR_ROW_ODD } = getGlobalSheetConstants();
const HEADER_FONT_COLOR = '#FFFFFF';
const HEADER_BORDER_COLOR = '#666666';
const DATA_FONT_COLOR = '#DDDDDD';

/**
 * Applies dark-mode styling to the header row of the SEO URL Scraper sheet.
 * Freezes the header row, sets background, font styling, and border.
 * @throws {Error} If the spreadsheet or target sheet is unavailable.
 */
function applyHeaderStyles() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Unable to access the active spreadsheet.');
  if (typeof SHEET_NAME !== 'string' || !SHEET_NAME.trim()) {
    throw new Error('SHEET_NAME is not defined or invalid.');
  }
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
  
  sheet.setFrozenRows(1);
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    throw new Error(`Sheet "${SHEET_NAME}" has no columns.`);
  }
  
  const headerRange = sheet.getRange(1, 1, 1, lastColumn);
  headerRange
    .setBackground(COLOR_HEADER_BG)
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
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Unable to access the active spreadsheet.');
  if (typeof SHEET_NAME !== 'string' || !SHEET_NAME.trim()) {
    throw new Error('SHEET_NAME is not defined or invalid.');
  }
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No data rows to style.');
  const lastColumn = sheet.getLastColumn();
  
  const numRows = lastRow - 1;
  const numCols = lastColumn;
  const backgrounds = [];
  const fontColors = [];
  
  for (let i = 0; i < numRows; i++) {
    const rowIndex = i + 2;
    const bgColor = (rowIndex % 2 === 0) ? COLOR_ROW_EVEN : COLOR_ROW_ODD;
    backgrounds.push(new Array(numCols).fill(bgColor));
    fontColors.push(new Array(numCols).fill(DATA_FONT_COLOR));
  }
  
  const dataRange = sheet.getRange(2, 1, numRows, numCols);
  dataRange
    .setBackgrounds(backgrounds)
    .setFontColors(fontColors);
}