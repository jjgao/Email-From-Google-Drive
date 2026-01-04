/**
 * Recipient Manager
 * Handles reading recipients from Google Sheets and validating email addresses
 */

/**
 * Get the recipient sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Recipient sheet
 */
function getRecipientSheet() {
  const sheetName = getConfig(CONFIG_KEYS.RECIPIENT_SHEET_NAME);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Recipient sheet "${sheetName}" not found. Please create it or update configuration.`);
  }

  return sheet;
}

/**
 * Get all recipients from the sheet
 * @returns {Array<Object>} Array of recipient objects with all data
 */
function getAllRecipients() {
  const sheet = getRecipientSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    throw new Error('No recipients found. Please add recipients to the sheet.');
  }

  // First row is header
  const headers = data[0];
  const recipients = [];

  // Process each row (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const recipient = {};

    // Create object with header as key and cell value as value
    for (let j = 0; j < headers.length; j++) {
      recipient[headers[j]] = row[j];
    }

    // Skip empty rows (no email)
    if (recipient.Email && recipient.Email.toString().trim() !== '') {
      recipient._rowIndex = i + 1; // Store row index for updating status
      recipients.push(recipient);
    }
  }

  return recipients;
}

/**
 * Get recipients ready for document creation (no Doc ID yet)
 * Status is only used for email sending, not document creation
 * @returns {Array<Object>} Array of recipients with {row, data} structure
 */
function getPendingRecipients() {
  const allRecipients = getAllRecipients();

  return allRecipients
    .filter(recipient => {
      const docId = (recipient['Doc ID'] || '').toString().trim();
      // Include only recipients without a Doc ID
      return docId === '';
    })
    .map(recipient => {
      // Return in format expected by document generator
      const rowIndex = recipient._rowIndex;
      delete recipient._rowIndex; // Remove internal property
      return {
        row: rowIndex,
        data: recipient
      };
    });
}

/**
 * Get all recipients formatted for document/PDF operations
 * @returns {Array<Object>} Array of recipients with {row, data} structure
 */
function getAllRecipientsFormatted() {
  const allRecipients = getAllRecipients();

  return allRecipients.map(recipient => {
    const rowIndex = recipient._rowIndex;
    delete recipient._rowIndex;
    return {
      row: rowIndex,
      data: recipient
    };
  });
}

/**
 * Get first recipient (for test email)
 * @returns {Object|null} First recipient or null if none found
 */
function getFirstRecipient() {
  const recipients = getAllRecipients();
  return recipients.length > 0 ? recipients[0] : null;
}

/**
 * Validate email address format
 * @param {string} email - Email address to validate
 * @returns {boolean} True if valid email format
 */
function isValidEmail(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }

  // Basic email regex pattern
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email.trim());
}

/**
 * Update recipient status in sheet
 * @param {number} rowIndex - Row index in sheet (1-based)
 * @param {string} status - Status to set (e.g., 'sent', 'failed', 'pending')
 */
function updateRecipientStatus(rowIndex, status) {
  const sheet = getRecipientSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find Status column index
  const statusColumnIndex = headers.indexOf('Status') + 1;

  if (statusColumnIndex > 0) {
    sheet.getRange(rowIndex, statusColumnIndex).setValue(status);
  } else {
    // Status column doesn't exist, add it
    const lastColumn = sheet.getLastColumn();
    sheet.getRange(1, lastColumn + 1).setValue('Status');
    sheet.getRange(rowIndex, lastColumn + 1).setValue(status);
  }
}

/**
 * Ensure Status column exists in recipient sheet
 */
function ensureStatusColumn() {
  const sheet = getRecipientSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Check if Status column exists
  if (headers.indexOf('Status') === -1) {
    const lastColumn = sheet.getLastColumn();
    sheet.getRange(1, lastColumn + 1).setValue('Status');
  }
}

/**
 * Validate recipient data
 * @param {Object} recipient - Recipient object
 * @returns {Object} Object with isValid boolean and error message
 */
function validateRecipient(recipient) {
  // Check if email exists
  if (!recipient.Email) {
    return {
      isValid: false,
      error: 'Email address is missing'
    };
  }

  // Check email format
  if (!isValidEmail(recipient.Email)) {
    return {
      isValid: false,
      error: 'Invalid email format'
    };
  }

  return {
    isValid: true,
    error: null
  };
}

/**
 * Get recipient count summary
 * @returns {Object} Summary with total, pending, sent, failed counts
 */
function getRecipientSummary() {
  const recipients = getAllRecipients();
  const summary = {
    total: recipients.length,
    pending: 0,
    sent: 0,
    failed: 0
  };

  recipients.forEach(recipient => {
    const status = (recipient.Status || 'pending').toString().toLowerCase();
    if (status === 'sent') {
      summary.sent++;
    } else if (status === 'failed') {
      summary.failed++;
    } else {
      summary.pending++;
    }
  });

  return summary;
}
