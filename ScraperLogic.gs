function runScrape() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet');
  var sheetName = getGlobalSheetConstants().SHEET_NAME;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" not found');
  var urls = getUrls();
  var totalCount = urls.length;
  var successCount = 0;
  var errors = [];
  for (var i = 0; i < urls.length; i++) {
    if (typeof isCancelled === 'function' && isCancelled()) break;
    var url = urls[i];
    var row = 2 + i;
    try {
      var html = fetchUrlData(url);
      var data = parseHtml(html);
      writeData(row, data);
      successCount++;
    } catch (e) {
      errors.push({row: row, message: e.message || e.toString()});
    }
  }
  SheetStyler.applyDataRowStyles();
  return {totalCount: totalCount, successCount: successCount, errors: errors};
}

function getUrls() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet');
  var sheetName = getGlobalSheetConstants().SHEET_NAME;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" not found');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var numRows = lastRow - 1;
  var values = sheet.getRange(2, 1, numRows, 1).getValues();
  var urls = [];
  values.forEach(function(row) {
    var cell = row[0];
    if (cell != null) {
      var url = cell.toString().trim();
      if (url) urls.push(url);
    }
  });
  return urls;
}

function fetchUrlData(url) {
  if (!url || typeof url !== 'string') throw new Error('Invalid URL input');
  var sanitized = url.trim();
  if (!/^https?:\/\//i.test(sanitized)) throw new Error('URL must start with http:// or https://');
  var response;
  try {
    response = UrlFetchApp.fetch(sanitized, {
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true
    });
  } catch (e) {
    throw new Error('Fetch failed for ' + sanitized + ': ' + e.message);
  }
  var code = response.getResponseCode();
  if (code < 200 || code >= 400) throw new Error('HTTP Error ' + code + ' for URL ' + sanitized);
  return response.getContentText();
}

function parseHtml(html) {
  if (html == null) throw new Error('Missing HTML content');
  html = html.toString();
  var titleMatch = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  var title = titleMatch ? titleMatch[1].trim() : '';
  var h1Matches = html.match(/<h1\b[^>]*>/gi);
  var h1Count = h1Matches ? h1Matches.length : 0;
  var descMatch = html.match(/<meta\s+name=["']description["']\s+content=["']([\s\S]*?)["'][^>]*>/i) ||
                  html.match(/<meta\s+content=["']([\s\S]*?)["']\s+name=["']description["'][^>]*>/i);
  var metaDesc = descMatch ? descMatch[1].trim() : '';
  var robotsMatch = html.match(/<meta\s+name=["']robots["']\s+content=["']([\s\S]*?)["'][^>]*>/i) ||
                    html.match(/<meta\s+content=["']([\s\S]*?)["']\s+name=["']robots["'][^>]*>/i);
  var robotsMeta = robotsMatch ? robotsMatch[1].trim() : '';
  return [title, h1Count, metaDesc, robotsMeta];
}

function writeData(row, data) {
  if (typeof row !== 'number' || row < 2) throw new Error('Invalid row number: ' + row);
  if (!Array.isArray(data) || data.length === 0) throw new Error('Invalid data for writeData');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet');
  var sheetName = getGlobalSheetConstants().SHEET_NAME;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" not found');
  sheet.getRange(row, 2, 1, data.length).setValues([data]);
}