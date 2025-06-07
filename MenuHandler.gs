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
    var sheet = getOrCreateScrapeSheet();
    SheetStyler.applyHeaderStyles();
    sheet.activate();
    sheet.getRange(globals.DATA_START_ROW_INDEX, globals.COL_URL).activate();
    var html = HtmlService.createTemplateFromFile('SidebarHandler')
      .evaluate()
      .setTitle('URL Scraper');
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showToast('Error opening scraper: ' + e.message, 'ERROR');
  }
}

