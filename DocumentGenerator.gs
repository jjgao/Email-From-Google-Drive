/**
 * Document Generator
 * Creates personalized Google Docs and PDFs from templates for each recipient
 *
 * Features:
 * - Smart template-based validation (only validates fields used in template)
 * - Personalized document generation from Google Doc templates
 * - PDF export from generated documents
 * - Incremental and full regeneration support
 * - Orphan file cleanup utilities
 * - Comprehensive logging of all operations
 *
 * @fileoverview Document and PDF generation module
 */

/**
 * Extract all placeholder field names from template document
 * @param {string} templateDocId - Template document ID
 * @returns {Object} Object with requiredFields and optionalFields arrays
 */
function getTemplateFields(templateDocId) {
  try {
    const templateDoc = DocumentApp.openById(templateDocId);
    const body = templateDoc.getBody();
    const text = body.getText();

    // Find all {{FieldName}} and {{?FieldName}} patterns
    const placeholderPattern = /\{\{(\??[^}]+)\}\}/g;
    const requiredFields = new Set();
    const optionalFields = new Set();

    let match;
    while ((match = placeholderPattern.exec(text)) !== null) {
      const fieldName = match[1].trim();

      // Check if field is optional (starts with ?)
      if (fieldName.startsWith('?')) {
        const cleanFieldName = fieldName.substring(1).trim();
        optionalFields.add(cleanFieldName);
      } else {
        requiredFields.add(fieldName);
      }
    }

    return {
      requiredFields: Array.from(requiredFields),
      optionalFields: Array.from(optionalFields),
      allFields: Array.from(new Set([...requiredFields, ...optionalFields]))
    };
  } catch (error) {
    throw new Error(`Failed to read template: ${error.message}`);
  }
}

/**
 * Validate that recipient has all required fields filled (based on template)
 * Optional fields (marked with {{?FieldName}}) are not validated
 * @param {string} templateDocId - Template document ID
 * @param {Object} recipientData - Recipient data
 * @returns {Object} Object with isValid boolean and missing array
 */
function validateRecipientData(templateDocId, recipientData) {
  const templateFields = getTemplateFields(templateDocId);
  const missing = [];

  // Only validate required fields (not optional ones)
  for (const fieldName of templateFields.requiredFields) {
    const value = recipientData[fieldName];
    if (!value || value.toString().trim() === '') {
      missing.push(fieldName);
    }
  }

  return {
    isValid: missing.length === 0,
    missing: missing,
    requiredFields: templateFields.requiredFields,
    optionalFields: templateFields.optionalFields
  };
}

/**
 * Create a personalized document for a single recipient
 * @param {string} templateDocId - Template document ID
 * @param {Object} recipientData - Recipient data with placeholders
 * @param {string} folderId - Output folder ID
 * @returns {Object} Object with docId and docUrl
 */
function createPersonalizedDocument(templateDocId, recipientData, folderId) {
  try {
    // Validate that all required fields are filled (based on template)
    const validation = validateRecipientData(templateDocId, recipientData);
    if (!validation.isValid) {
      throw new Error(`Missing required fields: ${validation.missing.join(', ')}`);
    }

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
 * Supports:
 * 1. Custom "Filename" column in recipient sheet (highest priority)
 * 2. DOCUMENT_NAME_TEMPLATE config with {{placeholders}}
 * 3. Fallback to "Name - Date" format
 * @param {Object} recipientData - Recipient data
 * @returns {string} Document name
 */
function generateDocumentName(recipientData) {
  // Priority 1: Check if recipient has a custom filename specified
  if (recipientData.Filename && recipientData.Filename.toString().trim() !== '') {
    return recipientData.Filename.toString().trim();
  }

  // Priority 2: Use template from config with placeholder replacement
  try {
    const template = getConfig(CONFIG_KEYS.DOCUMENT_NAME_TEMPLATE);
    if (template && template.trim() !== '') {
      // Use replacePlaceholders from TemplateParser to support {{Field}} syntax
      const name = replacePlaceholders(template, recipientData);
      // Add timestamp if not already in template
      if (!template.includes('{{') || name.trim() === '') {
        // Template didn't have placeholders or resulted in empty string
        throw new Error('Invalid template');
      }
      return name.trim();
    }
  } catch (error) {
    // Fall through to default behavior
  }

  // Priority 3: Fallback to default format
  // Try to use First Name + Last Name, then Name, then Email
  let identifier = '';
  if (recipientData['First Name'] || recipientData['Last Name']) {
    const firstName = recipientData['First Name'] || '';
    const lastName = recipientData['Last Name'] || '';
    identifier = `${lastName}, ${firstName}`.replace(/, $/, '').replace(/^, /, '');
  } else if (recipientData.Name) {
    identifier = recipientData.Name;
  } else {
    identifier = recipientData.Email || 'Document';
  }

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return `${identifier} - ${timestamp}`;
}

/**
 * Clean up orphaned punctuation in Google Doc after placeholder replacement
 * @param {GoogleAppsScript.Document.Body} body - Document body
 */
function cleanupDocumentPunctuation(body) {
  // Remove leading commas and spaces at the start of paragraphs
  body.replaceText('^[\\s,]+', '');

  // Remove orphaned commas (duplicate commas)
  body.replaceText(',\\s*,', ',');

  // Remove leading comma+space patterns
  body.replaceText(',\\s+(?=[A-Z])', '');

  // Remove trailing commas at end of paragraphs
  body.replaceText(',\\s*$', '');

  // Collapse multiple spaces
  body.replaceText('  +', ' ');

  // Remove paragraphs that are only whitespace or punctuation
  body.replaceText('^[\\s,;.]+$', '');
}

/**
 * Replace placeholders in the entire document
 * Supports both {{FieldName}} and {{?FieldName}} (optional) syntax
 * @param {GoogleAppsScript.Document.Body} body - Document body
 * @param {Object} data - Replacement data
 */
function replaceInDocument(body, data) {
  // Replace in regular text
  for (const key in data) {
    // Skip internal/system fields
    if (key === '_rowIndex' || key === 'Status' || key === 'Doc ID' || key === 'PDF ID') {
      continue;
    }

    const value = data[key] || '';

    // Replace both {{FieldName}} and {{?FieldName}} patterns
    const requiredPlaceholder = `{{${key}}}`;
    const optionalPlaceholder = `{{?${key}}}`;

    // Escape special regex characters in the placeholders
    const escapedRequired = requiredPlaceholder.replace(/[{}]/g, '\\$&');
    const escapedOptional = optionalPlaceholder.replace(/[{}?]/g, '\\$&');

    // Replace in body text
    body.replaceText(escapedRequired, value);
    body.replaceText(escapedOptional, value);
  }

  // Clean up orphaned punctuation from empty optional fields
  cleanupDocumentPunctuation(body);
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

  const templateDocId = getConfig(CONFIG_KEYS.PDF_TEMPLATE_DOC_ID);
  const folderId = getConfig(CONFIG_KEYS.OUTPUT_FOLDER_ID);

  if (!templateDocId) {
    throw new Error('PDF Template Document ID not configured. Please set it in the Config sheet to generate documents.');
  }

  const recipients = getRecipientsForDocumentCreation();

  if (recipients.length === 0) {
    throw new Error('No recipients without documents found. All recipients already have Doc IDs.');
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

      // Clear PDF ID since document has been regenerated (PDF would be out of sync)
      clearRecipientPdfId(recipient.row);

      // Log success
      logDocumentCreation(recipient.data, result.docId, result.docUrl, 'success');

      results.success++;

    } catch (error) {
      // Log failure
      logDocumentCreation(recipient.data, null, null, 'failed', error.message);

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

  const templateDocId = getConfig(CONFIG_KEYS.PDF_TEMPLATE_DOC_ID);
  const folderId = getConfig(CONFIG_KEYS.OUTPUT_FOLDER_ID);
  const testEmail = getConfig(CONFIG_KEYS.TEST_EMAIL);

  if (!templateDocId) {
    throw new Error('PDF Template Document ID not configured. Please set it in the Config sheet to generate documents.');
  }

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

/**
 * Generate PDF from a Google Doc
 * @param {string} docId - Document ID
 * @param {string} pdfFolderId - Folder to save PDF
 * @param {string} pdfName - Name for the PDF file
 * @returns {Object} Object with pdfId and pdfUrl
 */
function generatePdfFromDoc(docId, pdfFolderId, pdfName) {
  try {
    // Get the document as a PDF blob
    const doc = DriveApp.getFileById(docId);
    const blob = doc.getAs('application/pdf');
    blob.setName(pdfName);

    // Get the PDF folder
    const folder = DriveApp.getFolderById(pdfFolderId);

    // Create the PDF file in the folder
    const pdfFile = folder.createFile(blob);

    return {
      pdfId: pdfFile.getId(),
      pdfUrl: pdfFile.getUrl(),
      pdfName: pdfFile.getName()
    };

  } catch (error) {
    throw new Error(`Failed to generate PDF: ${error.message}`);
  }
}

/**
 * Generate PDFs for all recipients with documents
 * @returns {Object} Results with success and failure counts
 */
function generateAllPdfs() {
  const pdfFolderId = getConfig(CONFIG_KEYS.PDF_FOLDER_ID);

  if (!pdfFolderId) {
    throw new Error('PDF Folder ID not configured. Please set it in the Config sheet.');
  }

  // Get all recipients with Doc IDs but no PDF IDs
  const recipients = getRecipientsNeedingPdfs();

  if (recipients.length === 0) {
    throw new Error('No recipients with documents needing PDFs');
  }

  const results = {
    total: recipients.length,
    success: 0,
    failed: 0,
    errors: []
  };

  for (const recipient of recipients) {
    try {
      const docId = recipient.data['Doc ID'];
      const pdfName = `${recipient.data.Name || recipient.data.Email} - PDF`;

      // Generate PDF from document
      const result = generatePdfFromDoc(docId, pdfFolderId, pdfName);

      // Update recipient with PDF ID
      updateRecipientPdfId(recipient.row, result.pdfId);

      // Log success
      logPdfGeneration(recipient.data, result.pdfId, result.pdfUrl, 'success');

      results.success++;

    } catch (error) {
      // Log failure
      logPdfGeneration(recipient.data, null, null, 'failed', error.message);

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
 * Get recipients that have Doc IDs but no PDF IDs
 * @returns {Array<Object>} Array of recipients with {row, data} structure
 */
function getRecipientsNeedingPdfs() {
  const allRecipients = getAllRecipients();

  return allRecipients
    .filter(recipient => {
      const docId = (recipient['Doc ID'] || '').toString().trim();
      const pdfId = (recipient['PDF ID'] || '').toString().trim();
      // Include if has doc ID but no PDF ID
      return docId !== '' && pdfId === '';
    })
    .map(recipient => {
      const rowIndex = recipient._rowIndex;
      delete recipient._rowIndex;
      return {
        row: rowIndex,
        data: recipient
      };
    });
}

/**
 * Update recipient row with PDF ID
 * @param {number} row - Row number in sheet
 * @param {string} pdfId - PDF file ID
 */
function updateRecipientPdfId(row, pdfId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getConfig(CONFIG_KEYS.RECIPIENT_SHEET_NAME);
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }

  // Find PDF ID column
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const pdfIdColumn = headers.indexOf('PDF ID') + 1;

  if (pdfIdColumn === 0) {
    throw new Error('PDF ID column not found in recipient sheet');
  }

  // Update the cell with PDF ID
  sheet.getRange(row, pdfIdColumn).setValue(pdfId);
}

/**
 * Log PDF generation activity
 * @param {Object} recipientData - Recipient data
 * @param {string} pdfId - PDF file ID (null if failed)
 * @param {string} pdfUrl - PDF URL (null if failed)
 * @param {string} status - Status (success/failed)
 * @param {string} errorMessage - Error message if failed
 */
function logPdfGeneration(recipientData, pdfId, pdfUrl, status, errorMessage = '') {
  const logSheet = getLogSheet();
  const timestamp = new Date();

  const rowData = [
    timestamp,
    'PDF Generation',
    recipientData.Email || '',
    recipientData.Name || '',
    status,
    pdfId || '',
    pdfUrl || '',
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
  }
}

/**
 * Clear PDF ID from recipient row
 * Called when document is regenerated to ensure PDF stays in sync
 * @param {number} row - Row number in sheet
 */
function clearRecipientPdfId(row) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getConfig(CONFIG_KEYS.RECIPIENT_SHEET_NAME);
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    return; // Silently return if sheet not found
  }

  // Find PDF ID column
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const pdfIdColumn = headers.indexOf('PDF ID') + 1;

  if (pdfIdColumn > 0) {
    // Clear the cell
    sheet.getRange(row, pdfIdColumn).setValue('');
  }
}

/**
 * Regenerate documents for ALL recipients (ignoring existing Doc IDs)
 * @returns {Object} Results with success and failure counts
 */
function regenerateAllDocuments() {
  const validation = validateConfig();
  if (!validation.isValid) {
    throw new Error(`Missing required configuration: ${validation.missing.join(', ')}`);
  }

  const templateDocId = getConfig(CONFIG_KEYS.PDF_TEMPLATE_DOC_ID);
  const folderId = getConfig(CONFIG_KEYS.OUTPUT_FOLDER_ID);

  if (!templateDocId) {
    throw new Error('PDF Template Document ID not configured. Please set it in the Config sheet to generate documents.');
  }

  const recipients = getAllRecipientsFormatted();

  if (recipients.length === 0) {
    throw new Error('No recipients found');
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

      // Clear PDF ID since document has been regenerated
      clearRecipientPdfId(recipient.row);

      // Log success
      logDocumentCreation(recipient.data, result.docId, result.docUrl, 'regenerated');

      results.success++;

    } catch (error) {
      // Log failure
      logDocumentCreation(recipient.data, null, null, 'failed', error.message);

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
 * Regenerate PDFs for ALL recipients with documents (ignoring existing PDF IDs)
 * @returns {Object} Results with success and failure counts
 */
function regenerateAllPdfs() {
  const pdfFolderId = getConfig(CONFIG_KEYS.PDF_FOLDER_ID);

  if (!pdfFolderId) {
    throw new Error('PDF Folder ID not configured. Please set it in the Config sheet.');
  }

  // Get all recipients with Doc IDs (ignore PDF ID status)
  const allRecipients = getAllRecipientsFormatted();
  const recipients = allRecipients.filter(r => {
    const docId = (r.data['Doc ID'] || '').toString().trim();
    return docId !== '';
  });

  if (recipients.length === 0) {
    throw new Error('No recipients with documents found');
  }

  const results = {
    total: recipients.length,
    success: 0,
    failed: 0,
    errors: []
  };

  for (const recipient of recipients) {
    try {
      const docId = recipient.data['Doc ID'];
      const pdfName = `${recipient.data.Name || recipient.data.Email} - PDF`;

      // Generate PDF from document
      const result = generatePdfFromDoc(docId, pdfFolderId, pdfName);

      // Update recipient with PDF ID
      updateRecipientPdfId(recipient.row, result.pdfId);

      // Log success
      logPdfGeneration(recipient.data, result.pdfId, result.pdfUrl, 'regenerated');

      results.success++;

    } catch (error) {
      // Log failure
      logPdfGeneration(recipient.data, null, null, 'failed', error.message);

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
 * Delete orphan documents from output folder
 * (files that don't match any current recipient)
 * @returns {Object} Results with deleted count
 */
function deleteOrphanDocuments() {
  const folderId = getConfig(CONFIG_KEYS.OUTPUT_FOLDER_ID);

  if (!folderId) {
    throw new Error('Output Folder ID not configured.');
  }

  const allRecipients = getAllRecipients();
  const validDocIds = allRecipients
    .map(r => (r['Doc ID'] || '').toString().trim())
    .filter(id => id !== '');

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();

  let deletedCount = 0;
  const deletedFiles = [];

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();

    // If file ID is not in our valid Doc IDs, it's an orphan
    if (!validDocIds.includes(fileId)) {
      deletedFiles.push({
        name: file.getName(),
        id: fileId
      });
      file.setTrashed(true);
      deletedCount++;
    }
  }

  return {
    deleted: deletedCount,
    files: deletedFiles
  };
}

/**
 * Delete orphan PDFs from PDF folder
 * (files that don't match any current recipient)
 * @returns {Object} Results with deleted count
 */
function deleteOrphanPdfs() {
  const pdfFolderId = getConfig(CONFIG_KEYS.PDF_FOLDER_ID);

  if (!pdfFolderId) {
    throw new Error('PDF Folder ID not configured.');
  }

  const allRecipients = getAllRecipients();
  const validPdfIds = allRecipients
    .map(r => (r['PDF ID'] || '').toString().trim())
    .filter(id => id !== '');

  const folder = DriveApp.getFolderById(pdfFolderId);
  const files = folder.getFiles();

  let deletedCount = 0;
  const deletedFiles = [];

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();

    // If file ID is not in our valid PDF IDs, it's an orphan
    if (!validPdfIds.includes(fileId)) {
      deletedFiles.push({
        name: file.getName(),
        id: fileId
      });
      file.setTrashed(true);
      deletedCount++;
    }
  }

  return {
    deleted: deletedCount,
    files: deletedFiles
  };
}

/**
 * Delete all orphan files (both documents and PDFs)
 * @returns {Object} Results with deleted counts
 */
function deleteAllOrphanFiles() {
  const results = {
    documents: { deleted: 0, files: [] },
    pdfs: { deleted: 0, files: [] }
  };

  // Delete orphan documents
  try {
    results.documents = deleteOrphanDocuments();
  } catch (error) {
    if (error.message !== 'Output Folder ID not configured.') {
      throw error;
    }
  }

  // Delete orphan PDFs
  try {
    results.pdfs = deleteOrphanPdfs();
  } catch (error) {
    if (error.message !== 'PDF Folder ID not configured.') {
      throw error;
    }
  }

  return results;
}
