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
 * @returns {Array<Object>} Array of recipients with {row, data} structure
 */
function getRecipientsForDocumentCreation() {
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
 * Get recipients pending email (status is 'pending' or empty)
 * @returns {Array<Object>} Array of recipients ready for email sending
 */
function getPendingRecipients() {
  const allRecipients = getAllRecipients();

  return allRecipients.filter(recipient => {
    const status = (recipient['Email Status'] || '').toString().trim().toLowerCase();
    // Include if status is empty or 'pending'
    return status === '' || status === 'pending';
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

  // Find Email Status column index
  const statusColumnIndex = headers.indexOf('Email Status') + 1;

  if (statusColumnIndex > 0) {
    sheet.getRange(rowIndex, statusColumnIndex).setValue(status);
  } else {
    // Email Status column doesn't exist, add it
    const lastColumn = sheet.getLastColumn();
    sheet.getRange(1, lastColumn + 1).setValue('Email Status');
    sheet.getRange(rowIndex, lastColumn + 1).setValue(status);
  }
}

/**
 * Ensure Email Status column exists in recipient sheet
 */
function ensureStatusColumn() {
  const sheet = getRecipientSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Check if Email Status column exists
  if (headers.indexOf('Email Status') === -1) {
    const lastColumn = sheet.getLastColumn();
    sheet.getRange(1, lastColumn + 1).setValue('Email Status');
  }
}

/**
 * Ensure all required columns exist: Filename, Doc ID, PDF ID, Email Status
 * Creates missing columns and fills default values
 * @returns {Object} Results with columns created and values filled
 */
function ensureRequiredColumns() {
  const sheet = getRecipientSheet();
  let lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

  const results = {
    columnsCreated: [],
    filenamesFilled: 0,
    statusesFilled: 0
  };

  // Handle empty sheet (no columns)
  if (lastColumn === 0) {
    throw new Error('Recipient sheet has no columns. Please add at least an Email column first.');
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  // Define required columns in order they should appear
  const requiredColumns = ['Filename', 'Doc ID', 'PDF ID', 'Email Status'];

  // Track current headers (will be updated as we add columns)
  let currentHeaders = [...headers];
  let currentLastColumn = lastColumn;

  // Create missing columns
  for (const colName of requiredColumns) {
    if (currentHeaders.indexOf(colName) === -1) {
      // Add column at the end
      currentLastColumn++;
      const headerCell = sheet.getRange(1, currentLastColumn);
      headerCell.setValue(colName);

      // Format header to match others
      headerCell.setFontWeight('bold');
      headerCell.setBackground('#4285f4');
      headerCell.setFontColor('#ffffff');

      // Set column width based on column type
      if (colName === 'Filename') {
        sheet.setColumnWidth(currentLastColumn, 200);
      } else if (colName === 'Email Status') {
        sheet.setColumnWidth(currentLastColumn, 100);
      } else {
        sheet.setColumnWidth(currentLastColumn, 120);
      }

      currentHeaders.push(colName);
      results.columnsCreated.push(colName);
    }
  }

  // Flush changes to ensure columns are created before reading
  SpreadsheetApp.flush();

  // If no data rows, just return
  if (lastRow < 2) {
    return results;
  }

  // Refresh headers after adding columns
  const updatedLastColumn = sheet.getLastColumn();
  const updatedHeaders = sheet.getRange(1, 1, 1, updatedLastColumn).getValues()[0];

  // Get column indices
  const filenameColIndex = updatedHeaders.indexOf('Filename') + 1;
  const statusColIndex = updatedHeaders.indexOf('Email Status') + 1;

  // Get all data for processing
  const dataRange = sheet.getRange(2, 1, lastRow - 1, updatedLastColumn);
  const allData = dataRange.getValues();

  // Process each row
  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const rowIndex = i + 2; // Account for header row and 0-based index

    // Check if this row has an email (valid row)
    const emailColIndex = updatedHeaders.indexOf('Email');
    if (emailColIndex === -1 || !row[emailColIndex] || row[emailColIndex].toString().trim() === '') {
      continue; // Skip empty rows
    }

    // Build recipient data object for this row
    const recipientData = {};
    for (let j = 0; j < updatedHeaders.length; j++) {
      recipientData[updatedHeaders[j]] = row[j];
    }

    // Fill Filename if empty
    if (filenameColIndex > 0) {
      const currentFilename = row[filenameColIndex - 1];
      if (!currentFilename || currentFilename.toString().trim() === '') {
        const defaultFilename = generateDefaultDocumentName(recipientData);
        sheet.getRange(rowIndex, filenameColIndex).setValue(defaultFilename);
        results.filenamesFilled++;
      }
    }

    // Fill Status if empty
    if (statusColIndex > 0) {
      const currentStatus = row[statusColIndex - 1];
      if (!currentStatus || currentStatus.toString().trim() === '') {
        sheet.getRange(rowIndex, statusColIndex).setValue('pending');
        results.statusesFilled++;
      }
    }
  }

  return results;
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
    const status = (recipient['Email Status'] || 'pending').toString().toLowerCase();
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
