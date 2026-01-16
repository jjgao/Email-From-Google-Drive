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
 * Get or create an output folder for the current sheet
 * Creates folder structure: OutputFolder/SheetName/subfolder/
 * @param {string} subfolder - Subfolder name ('docs' or 'pdfs')
 * @returns {string} Folder ID for the target subfolder
 */
function getOrCreateOutputFolder(subfolder) {
  const parentFolderId = getConfig(CONFIG_KEYS.OUTPUT_FOLDER_ID);

  if (!parentFolderId) {
    throw new Error('Output Folder ID not configured. Please set it in the Config sheet.');
  }

  const sheetName = getCurrentRecipientSheetName();

  if (!sheetName || sheetName === '(none)') {
    throw new Error('No recipient sheet selected.');
  }

  // Get the parent folder
  const parentFolder = DriveApp.getFolderById(parentFolderId);

  // Get or create sheet-named subfolder
  let sheetFolder;
  const sheetFolders = parentFolder.getFoldersByName(sheetName);
  if (sheetFolders.hasNext()) {
    sheetFolder = sheetFolders.next();
  } else {
    sheetFolder = parentFolder.createFolder(sheetName);
  }

  // Get or create docs/pdfs subfolder
  let targetFolder;
  const targetFolders = sheetFolder.getFoldersByName(subfolder);
  if (targetFolders.hasNext()) {
    targetFolder = targetFolders.next();
  } else {
    targetFolder = sheetFolder.createFolder(subfolder);
  }

  return targetFolder.getId();
}

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
 * 3. Fallback to "Template Name - First Last" format
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
  // Use template name + First Last name format
  let templateName = 'Document';
  try {
    const templateDocId = getConfig(CONFIG_KEYS.PDF_TEMPLATE_DOC_ID);
    if (templateDocId) {
      const templateFile = DriveApp.getFileById(templateDocId);
      templateName = templateFile.getName();
    }
  } catch (error) {
    // Use default if template not found
  }

  // Build recipient name: First Last (or just what's available)
  let recipientName = '';

  // Try to build from First Name and Last Name
  const firstName = (recipientData['First Name'] || '').toString().trim();
  const lastName = (recipientData['Last Name'] || '').toString().trim();

  if (firstName && lastName) {
    // Both names available
    recipientName = `${firstName} ${lastName}`;
  } else if (firstName) {
    // Only first name
    recipientName = firstName;
  } else if (lastName) {
    // Only last name
    recipientName = lastName;
  } else if (recipientData.Name && recipientData.Name.toString().trim()) {
    // Try Name field
    recipientName = recipientData.Name.toString().trim();
  } else if (recipientData.Email && recipientData.Email.toString().trim()) {
    // Use email as fallback
    recipientName = recipientData.Email.toString().trim();
  } else {
    // Last resort
    recipientName = 'Recipient';
  }

  return `${templateName} - ${recipientName}`;
}

/**
 * Generate default document name (Priority 3 logic only)
 * Format: "Template Name - First Last"
 * Handles missing First Name or Last Name gracefully
 * @param {Object} recipientData - Recipient data
 * @returns {string} Default document name
 */
function generateDefaultDocumentName(recipientData) {
  // Get template name
  let templateName = 'Document';
  try {
    const templateDocId = getConfig(CONFIG_KEYS.PDF_TEMPLATE_DOC_ID);
    if (templateDocId) {
      const templateFile = DriveApp.getFileById(templateDocId);
      templateName = templateFile.getName();
    }
  } catch (error) {
    // Use default if template not found
  }

  // Build recipient name: First Last (or just what's available)
  let recipientName = '';

  // Try to build from First Name and Last Name
  const firstName = (recipientData['First Name'] || '').toString().trim();
  const lastName = (recipientData['Last Name'] || '').toString().trim();

  if (firstName && lastName) {
    // Both names available
    recipientName = `${firstName} ${lastName}`;
  } else if (firstName) {
    // Only first name
    recipientName = firstName;
  } else if (lastName) {
    // Only last name
    recipientName = lastName;
  } else if (recipientData.Name && recipientData.Name.toString().trim()) {
    // Try Name field
    recipientName = recipientData.Name.toString().trim();
  } else if (recipientData.Email && recipientData.Email.toString().trim()) {
    // Use email as fallback
    recipientName = recipientData.Email.toString().trim();
  } else {
    // Last resort
    recipientName = 'Recipient';
  }

  return `${templateName} - ${recipientName}`;
}

/**
 * Fill default filenames for all recipients
 * Creates Filename column if it doesn't exist
 * @returns {Object} Results with total, filled count, and samples
 */
function fillDefaultFilenames() {
  try {
    const sheet = getRecipientSheet();

    // Get headers
    const lastColumn = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    // Find or create Filename column
    let filenameColIndex = headers.indexOf('Filename');
    if (filenameColIndex === -1) {
      // Filename column doesn't exist, add it
      // Insert after ZIP column (column 8) if possible
      const zipColIndex = headers.indexOf('ZIP');
      if (zipColIndex !== -1) {
        sheet.insertColumnAfter(zipColIndex + 1);
        filenameColIndex = zipColIndex + 1;
      } else {
        // Otherwise add at the end
        sheet.insertColumnAfter(lastColumn);
        filenameColIndex = lastColumn;
      }

      // Set header
      sheet.getRange(1, filenameColIndex + 1).setValue('Filename');

      // Format header to match others
      const headerCell = sheet.getRange(1, filenameColIndex + 1);
      headerCell.setFontWeight('bold');
      headerCell.setBackground('#4285f4');
      headerCell.setFontColor('#ffffff');

      // Set column width
      sheet.setColumnWidth(filenameColIndex + 1, 200);
    }

    // Get all recipients
    const allRecipients = getAllRecipientsFormatted();
    const results = {
      total: allRecipients.length,
      filled: 0,
      sample: []
    };

    // Generate and fill filenames
    for (let i = 0; i < allRecipients.length; i++) {
      const recipient = allRecipients[i];
      const defaultFilename = generateDefaultDocumentName(recipient.data);

      // Write to sheet (row is i+2 because row 1 is headers and array is 0-indexed)
      const rowIndex = i + 2;
      sheet.getRange(rowIndex, filenameColIndex + 1).setValue(defaultFilename);

      results.filled++;

      // Collect samples
      if (results.sample.length < 10) {
        results.sample.push({
          email: recipient.email,
          filename: defaultFilename
        });
      }
    }

    return results;

  } catch (error) {
    Logger.log(`Error filling filenames: ${error.message}`);
    throw error;
  }
}

/**
 * Create document for first recipient only (Quick Test)
 * @returns {Object} Result with email, docId, docName
 */
function createFirstDocument() {
  const validation = validateConfig();
  if (!validation.isValid) {
    throw new Error(`Missing required configuration: ${validation.missing.join(', ')}`);
  }

  // Ensure required columns exist before proceeding
  ensureRequiredColumns();

  const templateDocId = getConfig(CONFIG_KEYS.PDF_TEMPLATE_DOC_ID);

  if (!templateDocId) {
    throw new Error('PDF Template Document ID not configured. Please set it in the Config sheet to generate documents.');
  }

  // Get or create the docs folder for this sheet
  const folderId = getOrCreateOutputFolder('docs');

  const allRecipients = getAllRecipientsFormatted();

  if (allRecipients.length === 0) {
    throw new Error('No recipients found in the sheet.');
  }

  const firstRecipient = allRecipients[0];

  try {
    // Create personalized document
    const result = createPersonalizedDocument(templateDocId, firstRecipient.data, folderId);

    // Update recipient with document ID
    updateRecipientDocId(firstRecipient.row, result.docId);

    // Clear PDF ID since document has been created/regenerated
    clearRecipientPdfId(firstRecipient.row);

    // Log success
    logDocumentCreation(firstRecipient.data, result.docId, result.docUrl, 'test');

    return {
      email: firstRecipient.email,
      docId: result.docId,
      docName: result.docName
    };

  } catch (error) {
    // Log failure
    logDocumentCreation(firstRecipient.data, null, null, 'failed', error.message);
    throw error;
  }
}

/**
 * Generate PDF for first recipient only (Quick Test)
 * @returns {Object} Result with email, pdfId, pdfName
 */
function generateFirstPdf() {
  // Get or create the pdfs folder for this sheet
  const pdfFolderId = getOrCreateOutputFolder('pdfs');

  const allRecipients = getAllRecipientsFormatted();

  if (allRecipients.length === 0) {
    throw new Error('No recipients found in the sheet.');
  }

  const firstRecipient = allRecipients[0];

  // Check if first recipient has a Doc ID
  const docId = (firstRecipient.data['Doc ID'] || '').toString().trim();
  if (!docId) {
    throw new Error(`First recipient (${firstRecipient.email}) does not have a document yet. Please create a document first.`);
  }

  try {
    // Use the same name as the document
    const docFile = DriveApp.getFileById(docId);
    const pdfName = docFile.getName();

    // Generate PDF from document
    const result = generatePdfFromDoc(docId, pdfFolderId, pdfName);

    // Update recipient with PDF ID
    updateRecipientPdfId(firstRecipient.row, result.pdfId);

    // Log success
    logPdfGeneration(firstRecipient.data, result.pdfId, result.pdfUrl, 'test');

    return {
      email: firstRecipient.email,
      pdfId: result.pdfId,
      pdfName: result.pdfName
    };

  } catch (error) {
    // Log failure
    logPdfGeneration(firstRecipient.data, null, null, 'failed', error.message);
    throw error;
  }
}

/**
 * Clean up orphaned punctuation in Google Doc after placeholder replacement
 * Note: Google Apps Script's replaceText() doesn't support lookahead assertions
 * @param {GoogleAppsScript.Document.Body} body - Document body
 */
function cleanupDocumentPunctuation(body) {
  // Remove leading commas and spaces at the start of paragraphs
  body.replaceText('^[\\s,]+', '');

  // Remove orphaned commas (duplicate commas)
  body.replaceText(',\\s*,', ',');

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
    if (key === '_rowIndex' || key === 'Email Status' || key === 'Doc ID' || key === 'PDF ID') {
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

  // Ensure required columns exist before proceeding
  ensureRequiredColumns();

  const templateDocId = getConfig(CONFIG_KEYS.PDF_TEMPLATE_DOC_ID);

  if (!templateDocId) {
    throw new Error('PDF Template Document ID not configured. Please set it in the Config sheet to generate documents.');
  }

  // Get or create the docs folder for this sheet
  const folderId = getOrCreateOutputFolder('docs');

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

  // Process each recipient
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
  const testEmail = getConfig(CONFIG_KEYS.TEST_EMAIL);

  if (!templateDocId) {
    throw new Error('PDF Template Document ID not configured. Please set it in the Config sheet to generate documents.');
  }

  // Get or create the docs folder for this sheet
  const folderId = getOrCreateOutputFolder('docs');

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
  const sheet = getRecipientSheet();

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
  // Get or create the pdfs folder for this sheet
  const pdfFolderId = getOrCreateOutputFolder('pdfs');

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

  // Process each recipient
  for (const recipient of recipients) {
    const docId = (recipient.data['Doc ID'] || '').toString().trim();
    if (docId === '') {
      continue;
    }

    try {
      // Use the same name as the document
      const docFile = DriveApp.getFileById(docId);
      const pdfName = docFile.getName();

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
 * Verify that all documents and PDFs exist in Google Drive
 * Checks each recipient's Doc ID and PDF ID to ensure the files are accessible
 * @returns {Object} Verification results with counts and lists of issues
 */
function verifyDocumentsAndPdfs() {
  const allRecipients = getAllRecipients();

  const results = {
    total: allRecipients.length,
    docs: {
      expected: 0,      // Recipients with Doc ID
      verified: 0,      // Files that exist
      missing: 0,       // Files that don't exist
      missingList: []   // List of missing doc info
    },
    pdfs: {
      expected: 0,      // Recipients with PDF ID
      verified: 0,      // Files that exist
      missing: 0,       // Files that don't exist
      missingList: []   // List of missing PDF info
    },
    noDocId: 0,         // Recipients without Doc ID
    noPdfId: 0          // Recipients without PDF ID
  };

  for (const recipient of allRecipients) {
    const email = recipient.Email || '(no email)';
    const docId = (recipient['Doc ID'] || '').toString().trim();
    const pdfId = (recipient['PDF ID'] || '').toString().trim();

    // Check Doc ID
    if (docId === '') {
      results.noDocId++;
    } else {
      results.docs.expected++;
      try {
        const file = DriveApp.getFileById(docId);
        // File exists, check if it's accessible
        file.getName();
        results.docs.verified++;
      } catch (e) {
        results.docs.missing++;
        results.docs.missingList.push({
          email: email,
          docId: docId,
          error: e.message
        });
      }
    }

    // Check PDF ID
    if (pdfId === '') {
      results.noPdfId++;
    } else {
      results.pdfs.expected++;
      try {
        const file = DriveApp.getFileById(pdfId);
        // File exists, check if it's accessible
        file.getName();
        results.pdfs.verified++;
      } catch (e) {
        results.pdfs.missing++;
        results.pdfs.missingList.push({
          email: email,
          pdfId: pdfId,
          error: e.message
        });
      }
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
  const sheet = getRecipientSheet();

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
  let sheet;
  try {
    sheet = getRecipientSheet();
  } catch (error) {
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

  // Ensure required columns exist before proceeding
  ensureRequiredColumns();

  const templateDocId = getConfig(CONFIG_KEYS.PDF_TEMPLATE_DOC_ID);

  if (!templateDocId) {
    throw new Error('PDF Template Document ID not configured. Please set it in the Config sheet to generate documents.');
  }

  // Get or create the docs folder for this sheet
  const folderId = getOrCreateOutputFolder('docs');

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
  // Get or create the pdfs folder for this sheet
  const pdfFolderId = getOrCreateOutputFolder('pdfs');

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

      // Use the same name as the document
      const docFile = DriveApp.getFileById(docId);
      const pdfName = docFile.getName();

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
 * Preview orphan documents (for confirmation before deletion)
 * Scans OutputFolder/SheetName/docs/ subfolder
 * @returns {Object} Preview with orphan and protected file lists
 */
function previewOrphanDocuments() {
  // Get the docs subfolder for this sheet
  let docsFolderId;
  try {
    docsFolderId = getOrCreateOutputFolder('docs');
  } catch (error) {
    throw new Error('Output Folder ID not configured or sheet not selected.');
  }

  const allRecipients = getAllRecipients();
  const validDocIds = allRecipients
    .map(r => (r['Doc ID'] || '').toString().trim())
    .filter(id => id !== '');

  const folder = DriveApp.getFolderById(docsFolderId);
  const files = folder.getFiles();

  const orphanFiles = [];
  const protectedFiles = [];
  let totalCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    totalCount++;

    const fileInfo = {
      name: file.getName(),
      id: fileId
    };

    if (!validDocIds.includes(fileId)) {
      orphanFiles.push(fileInfo);
    } else {
      protectedFiles.push(fileInfo);
    }
  }

  return {
    total: totalCount,
    orphanCount: orphanFiles.length,
    protectedCount: protectedFiles.length,
    orphanFiles: orphanFiles,
    protectedFiles: protectedFiles,
    validDocIdsCount: validDocIds.length
  };
}

/**
 * Delete orphan documents from output folder
 * Scans OutputFolder/SheetName/docs/ subfolder
 * @returns {Object} Results with deleted count
 */
function deleteOrphanDocuments() {
  // Get the docs subfolder for this sheet
  let docsFolderId;
  try {
    docsFolderId = getOrCreateOutputFolder('docs');
  } catch (error) {
    throw new Error('Output Folder ID not configured or sheet not selected.');
  }

  const allRecipients = getAllRecipients();
  const validDocIds = allRecipients
    .map(r => (r['Doc ID'] || '').toString().trim())
    .filter(id => id !== '');

  const folder = DriveApp.getFolderById(docsFolderId);
  const files = folder.getFiles();

  let deletedCount = 0;
  const deletedFiles = [];
  let totalCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    totalCount++;

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
    total: totalCount,
    deleted: deletedCount,
    files: deletedFiles,
    validDocIdsCount: validDocIds.length
  };
}

/**
 * Preview orphan PDFs (for confirmation before deletion)
 * Scans OutputFolder/SheetName/pdfs/ subfolder
 * @returns {Object} Preview with orphan and protected file lists
 */
function previewOrphanPdfs() {
  // Get the pdfs subfolder for this sheet
  let pdfsFolderId;
  try {
    pdfsFolderId = getOrCreateOutputFolder('pdfs');
  } catch (error) {
    throw new Error('Output Folder ID not configured or sheet not selected.');
  }

  const allRecipients = getAllRecipients();
  const validPdfIds = allRecipients
    .map(r => (r['PDF ID'] || '').toString().trim())
    .filter(id => id !== '');

  const folder = DriveApp.getFolderById(pdfsFolderId);
  const files = folder.getFiles();

  const orphanFiles = [];
  const protectedFiles = [];
  let totalCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    totalCount++;

    const fileInfo = {
      name: file.getName(),
      id: fileId
    };

    if (!validPdfIds.includes(fileId)) {
      orphanFiles.push(fileInfo);
    } else {
      protectedFiles.push(fileInfo);
    }
  }

  return {
    total: totalCount,
    orphanCount: orphanFiles.length,
    protectedCount: protectedFiles.length,
    orphanFiles: orphanFiles,
    protectedFiles: protectedFiles,
    validPdfIdsCount: validPdfIds.length
  };
}

/**
 * Delete orphan PDFs from PDF folder
 * Scans OutputFolder/SheetName/pdfs/ subfolder
 * @returns {Object} Results with deleted count
 */
function deleteOrphanPdfs() {
  // Get the pdfs subfolder for this sheet
  let pdfsFolderId;
  try {
    pdfsFolderId = getOrCreateOutputFolder('pdfs');
  } catch (error) {
    throw new Error('Output Folder ID not configured or sheet not selected.');
  }

  const allRecipients = getAllRecipients();
  const validPdfIds = allRecipients
    .map(r => (r['PDF ID'] || '').toString().trim())
    .filter(id => id !== '');

  const folder = DriveApp.getFolderById(pdfsFolderId);
  const files = folder.getFiles();

  let deletedCount = 0;
  const deletedFiles = [];
  let totalCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    totalCount++;

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
    total: totalCount,
    deleted: deletedCount,
    files: deletedFiles,
    validPdfIdsCount: validPdfIds.length
  };
}

/**
 * Preview all orphan files (both documents and PDFs)
 * Scans OutputFolder/SheetName/docs/ and OutputFolder/SheetName/pdfs/ subfolders
 * @returns {Object} Preview with orphan counts and file lists
 */
function previewAllOrphanFiles() {
  const results = {
    documents: { total: 0, orphanCount: 0, protectedCount: 0, orphanFiles: [], protectedFiles: [], validDocIdsCount: 0 },
    pdfs: { total: 0, orphanCount: 0, protectedCount: 0, orphanFiles: [], protectedFiles: [], validPdfIdsCount: 0 },
    sheetName: getCurrentRecipientSheetName()
  };

  // Scan docs subfolder
  try {
    results.documents = previewOrphanDocuments();
  } catch (error) {
    // Folder may not exist yet if no documents created
    if (!error.message.includes('not configured')) {
      throw error;
    }
  }

  // Scan pdfs subfolder
  try {
    results.pdfs = previewOrphanPdfs();
  } catch (error) {
    // Folder may not exist yet if no PDFs created
    if (!error.message.includes('not configured')) {
      throw error;
    }
  }

  return results;
}

/**
 * Delete all orphan files (both documents and PDFs)
 * Scans OutputFolder/SheetName/docs/ and OutputFolder/SheetName/pdfs/ subfolders
 * @returns {Object} Results with deleted counts
 */
function deleteAllOrphanFiles() {
  const results = {
    documents: { total: 0, deleted: 0, files: [], validDocIdsCount: 0 },
    pdfs: { total: 0, deleted: 0, files: [], validPdfIdsCount: 0 },
    sheetName: getCurrentRecipientSheetName()
  };

  // Delete from docs subfolder
  try {
    results.documents = deleteOrphanDocuments();
  } catch (error) {
    // Folder may not exist yet if no documents created
    if (!error.message.includes('not configured')) {
      throw error;
    }
  }

  // Delete from pdfs subfolder
  try {
    results.pdfs = deleteOrphanPdfs();
  } catch (error) {
    // Folder may not exist yet if no PDFs created
    if (!error.message.includes('not configured')) {
      throw error;
    }
  }

  return results;
}

/**
 * Debug function: Show detailed file and recipient ID mapping
 * Helps diagnose orphan file issues
 * Scans OutputFolder/SheetName/docs/ and OutputFolder/SheetName/pdfs/ subfolders
 * @returns {Object} Detailed diagnostic information
 */
function debugOrphanFiles() {
  const result = {
    sheetName: getCurrentRecipientSheetName(),
    recipients: [],
    documentFolder: { files: [], folderUrl: '' },
    pdfFolder: { files: [], folderUrl: '' }
  };

  // Get all recipients with their IDs
  const allRecipients = getAllRecipients();
  result.recipients = allRecipients.map(r => ({
    email: r.Email || '',
    docId: (r['Doc ID'] || '').toString().trim(),
    pdfId: (r['PDF ID'] || '').toString().trim()
  }));

  // Get document folder files
  try {
    const docFolderId = getOrCreateOutputFolder('docs');
    const docFolder = DriveApp.getFolderById(docFolderId);
    result.documentFolder.folderUrl = docFolder.getUrl();

    const docFiles = docFolder.getFiles();
    while (docFiles.hasNext()) {
      const file = docFiles.next();
      result.documentFolder.files.push({
        name: file.getName(),
        id: file.getId(),
        url: file.getUrl()
      });
    }
  } catch (error) {
    result.documentFolder.error = error.message;
  }

  // Get PDF folder files
  try {
    const pdfFolderId = getOrCreateOutputFolder('pdfs');
    const pdfFolder = DriveApp.getFolderById(pdfFolderId);
    result.pdfFolder.folderUrl = pdfFolder.getUrl();

    const pdfFiles = pdfFolder.getFiles();
    while (pdfFiles.hasNext()) {
      const file = pdfFiles.next();
      result.pdfFolder.files.push({
        name: file.getName(),
        id: file.getId(),
        url: file.getUrl()
      });
    }
  } catch (error) {
    result.pdfFolder.error = error.message;
  }

  return result;
}
