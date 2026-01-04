/**
 * Document Generator
 * Creates personalized Google Docs from template for each recipient
 */

/**
 * Create a personalized document for a single recipient
 * @param {string} templateDocId - Template document ID
 * @param {Object} recipientData - Recipient data with placeholders
 * @param {string} folderId - Output folder ID
 * @returns {Object} Object with docId and docUrl
 */
function createPersonalizedDocument(templateDocId, recipientData, folderId) {
  try {
    // Open template document
    const templateDoc = DocumentApp.openById(templateDocId);

    // Create a copy of the template
    const templateFile = DriveApp.getFileById(templateDocId);
    const folder = DriveApp.getFolderById(folderId);

    // Generate document name using recipient data
    const docName = generateDocumentName(recipientData);

    // Make a copy in the specified folder
    const newFile = templateFile.makeCopy(docName, folder);
    const newDocId = newFile.getId();

    // Open the new document and personalize it
    const newDoc = DocumentApp.openById(newDocId);
    const body = newDoc.getBody();

    // Replace all placeholders in the document
    replaceInDocument(body, recipientData);

    // Save and close
    newDoc.saveAndClose();

    return {
      docId: newDocId,
      docUrl: newFile.getUrl(),
      docName: docName
    };

  } catch (error) {
    throw new Error(`Failed to create document for ${recipientData.Name || recipientData.Email}: ${error.message}`);
  }
}

/**
 * Generate document name from recipient data
 * @param {Object} recipientData - Recipient data
 * @returns {string} Document name
 */
function generateDocumentName(recipientData) {
  // Use Name if available, otherwise use Email
  const identifier = recipientData.Name || recipientData.Email || 'Document';
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return `${identifier} - ${timestamp}`;
}

/**
 * Replace placeholders in the entire document
 * @param {GoogleAppsScript.Document.Body} body - Document body
 * @param {Object} data - Replacement data
 */
function replaceInDocument(body, data) {
  // Replace in regular text
  for (const key in data) {
    const placeholder = `{{${key}}}`;
    const value = data[key] || '';

    // Replace in body text
    body.replaceText(placeholder, value);
  }
}

/**
 * Create documents for all pending recipients
 * @returns {Object} Results with success and failure counts
 */
function createAllDocuments() {
  const validation = validateConfig();
  if (!validation.isValid) {
    throw new Error(`Missing required configuration: ${validation.missing.join(', ')}`);
  }

  const templateDocId = getConfig(CONFIG_KEYS.TEMPLATE_DOC_ID);
  const folderId = getConfig(CONFIG_KEYS.OUTPUT_FOLDER_ID);
  const recipients = getPendingRecipients();

  if (recipients.length === 0) {
    throw new Error('No pending recipients found');
  }

  const results = {
    total: recipients.length,
    success: 0,
    failed: 0,
    errors: []
  };

  for (const recipient of recipients) {
    try {
      // Create personalized document
      const result = createPersonalizedDocument(templateDocId, recipient.data, folderId);

      // Update recipient with document ID
      updateRecipientDocId(recipient.row, result.docId);
      updateRecipientStatus(recipient.row, 'created');

      // Log success
      logDocumentCreation(recipient.data, result.docId, result.docUrl, 'success');

      results.success++;

    } catch (error) {
      // Log failure
      logDocumentCreation(recipient.data, null, null, 'failed', error.message);
      updateRecipientStatus(recipient.row, 'failed');

      results.failed++;
      results.errors.push({
        recipient: recipient.data.Email || recipient.data.Name,
        error: error.message
      });
    }
  }

  return results;
}

/**
 * Create a single test document
 * @returns {Object} Document info
 */
function createTestDocument() {
  const validation = validateConfig();
  if (!validation.isValid) {
    throw new Error(`Missing required configuration: ${validation.missing.join(', ')}`);
  }

  const templateDocId = getConfig(CONFIG_KEYS.TEMPLATE_DOC_ID);
  const folderId = getConfig(CONFIG_KEYS.OUTPUT_FOLDER_ID);
  const testEmail = getConfig(CONFIG_KEYS.TEST_EMAIL);

  // Get first recipient as sample data
  const recipients = getAllRecipients();
  if (recipients.length === 0) {
    throw new Error('No recipients found to use as test data');
  }

  // Use first recipient's data but override email
  const testData = {...recipients[0].data};
  testData.Email = testEmail;
  testData.Name = testData.Name || 'Test User';

  const result = createPersonalizedDocument(templateDocId, testData, folderId);

  logDocumentCreation(testData, result.docId, result.docUrl, 'test');

  return result;
}

/**
 * Update recipient row with document ID
 * @param {number} row - Row number in sheet
 * @param {string} docId - Document ID
 */
function updateRecipientDocId(row, docId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getConfig(CONFIG_KEYS.RECIPIENT_SHEET_NAME);
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }

  // Find Doc ID column (should be column 9 based on our headers)
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const docIdColumn = headers.indexOf('Doc ID') + 1;

  if (docIdColumn === 0) {
    throw new Error('Doc ID column not found in recipient sheet');
  }

  // Update the cell with document ID
  sheet.getRange(row, docIdColumn).setValue(docId);
}

/**
 * Log document creation activity
 * @param {Object} recipientData - Recipient data
 * @param {string} docId - Document ID (null if failed)
 * @param {string} docUrl - Document URL (null if failed)
 * @param {string} status - Status (success/failed/test)
 * @param {string} errorMessage - Error message if failed
 */
function logDocumentCreation(recipientData, docId, docUrl, status, errorMessage = '') {
  const logSheet = getLogSheet();
  const timestamp = new Date();

  const rowData = [
    timestamp,
    'Document Creation',
    recipientData.Email || '',
    recipientData.Name || '',
    status,
    docId || '',
    docUrl || '',
    errorMessage
  ];

  logSheet.appendRow(rowData);

  // Color code the row
  const lastRow = logSheet.getLastRow();
  const range = logSheet.getRange(lastRow, 1, 1, rowData.length);

  if (status === 'success') {
    range.setBackground('#d4edda'); // Light green
  } else if (status === 'failed') {
    range.setBackground('#f8d7da'); // Light red
  } else if (status === 'test') {
    range.setBackground('#d1ecf1'); // Light blue
  }
}
