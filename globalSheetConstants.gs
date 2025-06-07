const GLOBAL_SHEET_CONSTANTS = Object.freeze({
  // Sheet and range settings
  SHEET_NAME: 'URL Scraping',
  HEADER_ROW_INDEX: 1,
  DATA_START_ROW_INDEX: 2,
  HEADER_LABELS: ['URL', 'Title', 'H1 Count', 'Meta Description', 'Robots'],
  // Column indexes (1-based)
  COL_URL: 1,
  COL_TITLE: 2,
  COL_H1_COUNT: 3,
  COL_META_DESCRIPTION: 4,
  COL_ROBOTS: 5,
  // Column letters
  COLUMN_LETTERS: Object.freeze({
    URL: 'A',
    TITLE: 'B',
    H1_COUNT: 'C',
    META_DESCRIPTION: 'D',
    ROBOTS: 'E'
  }),
  // Dark theme colors
  COLOR_HEADER_BG: '#222222',
  COLOR_HEADER_FONT: '#FFFFFF',
  COLOR_ROW_ODD: '#1A1A1A',
  COLOR_ROW_EVEN: '#2A2A2A',
  // Sidebar UI colors
  COLOR_SIDEBAR_BG: '#121212',
  COLOR_SIDEBAR_TEXT: '#FFFFFF',
  COLOR_BUTTON_BG: '#1F1F1F',
  COLOR_BUTTON_BORDER: '#444444',
  COLOR_BUTTON_HOVER_BG: '#2A2A2A',
  COLOR_ACCENT_GREEN: '#00E676',
  COLOR_ACCENT_BLUE: '#90CAF9',
  // Toast colors
  TOAST_SUCCESS_BG: '#2E7D32',
  TOAST_ERROR_BG: '#C62828',
  // Sidebar settings
  SIDEBAR_WIDTH: 300,
  // URL validation
  URL_REGEX: /^(https?:\/\/)([\w-]+(\.[\w-]+)+)([^\s]*)$/i,
  // UrlFetchApp options
  FETCH_OPTIONS: Object.freeze({
    muteHttpExceptions: true,
    followRedirects: true,
    validateHttpsCertificates: true,
    timeout: 30000
  })
});

/**
 * Returns the global constants object.
 * @returns {Object}
 */
function getGlobalSheetConstants() {
  return GLOBAL_SHEET_CONSTANTS;
}