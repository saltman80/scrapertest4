function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('URL Scraper')
      .addItem('Open Scraper', 'openSidebar')
      .addToUi();
  } catch (e) {
    Logger.log('onOpen error: ' + e.message);
  }
}

function openSidebar() {
  try {
    var globals = getGlobalSheetConstants();
    var SCRAPE_SHEET_NAME = globals.SHEET_NAME;
    var URL_COLUMN = globals.COL_URL;
    var DATA_START_ROW = globals.DATA_START_ROW_INDEX;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SCRAPE_SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + SCRAPE_SHEET_NAME + '" not found.');
    SheetStyler.applyHeaderStyles();
    sheet.activate();
    sheet.getRange(DATA_START_ROW, URL_COLUMN).activate();
    var html = HtmlService.createTemplateFromFile('SidebarHandler')
      .evaluate()
      .setTitle('URL Scraper');
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showToast('Error opening scraper: ' + e.message, 'ERROR');
  }
}