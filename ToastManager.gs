var ToastType = {
  INFO: 'INFO',
  SUCCESS: 'SUCCESS',
  WARNING: 'WARNING',
  ERROR: 'ERROR'
};

var ALLOWED_TOAST_TYPES = [
  ToastType.INFO,
  ToastType.SUCCESS,
  ToastType.WARNING,
  ToastType.ERROR
];

/**
 * Error thrown when validation fails.
 *
 * @param {string} message - The error message.
 * @constructor
 * @extends {Error}
 */
function ValidationError(message) {
  this.name = 'ValidationError';
  this.message = message || 'Validation error';
  if (Error.captureStackTrace) {
    Error.captureStackTrace(this, ValidationError);
  }
}
ValidationError.prototype = Object.create(Error.prototype);
ValidationError.prototype.constructor = ValidationError;

/**
 * Creates a toast payload for the sidebar UI and displays a toast in the spreadsheet.
 *
 * @param {string} message - The message to display; must be a non-empty string.
 * @param {string} type - The toast type; must be one of ToastType values.
 * @returns {{message: string, type: string}}
 * @throws {ValidationError} If inputs are invalid.
 */
function showToast(message, type) {
  if (typeof message !== 'string' || message.trim() === '') {
    throw new ValidationError('Invalid "message" parameter: must be a non-empty string.');
  }
  if (typeof type !== 'string' || type.trim() === '') {
    throw new ValidationError('Invalid "type" parameter: must be a non-empty string.');
  }
  var normalizedType = type.trim().toUpperCase();
  if (ALLOWED_TOAST_TYPES.indexOf(normalizedType) === -1) {
    throw new ValidationError('Invalid toast type: "' + type + '". Allowed types: ' + ALLOWED_TOAST_TYPES.join(', ') + '.');
  }
  var trimmedMessage = message.trim();
  // Display the toast in the spreadsheet UI
  SpreadsheetApp.getActiveSpreadsheet().toast(trimmedMessage, normalizedType, 5);
  return {
    message: trimmedMessage,
    type: normalizedType
  };
}