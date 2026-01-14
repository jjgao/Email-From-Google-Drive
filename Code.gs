/**
 * Email From Google Drive
 * Main entry point for the Google Apps Script application
 *
 * This application sends personalized emails using templates from Google Docs
 * and recipient data from Google Sheets.
 *
 * @author Email-From-Google-Drive
 * @version 0.3.0
 */

/**
 * Application information
 */
const APP_INFO = {
  name: 'Email From Google Drive',
  version: '0.3.0',
  description: 'Send personalized emails using Google Sheets and Google Docs'
};

/**
 * Get application version
 * @returns {string} Application version
 */
function getVersion() {
  return APP_INFO.version;
}

/**
 * Get application info
 * @returns {Object} Application information
 */
function getAppInfo() {
  return APP_INFO;
}

/**
 * Initialize the application
 * Called when the spreadsheet is opened
 */
function initializeApp() {
  // Ensure log sheet exists
  try {
    getLogSheet();
  } catch (error) {
    console.log('Log sheet will be created on first use');
  }

  // Ensure status column exists if recipient sheet exists
  try {
    // Try to get the recipient sheet and ensure status column exists
    getRecipientSheet();
    ensureStatusColumn();
  } catch (error) {
    console.log('Recipient sheet not found or status column already exists');
  }
}

/**
 * Test function to verify all modules are working
 * Run this from Script Editor to test the setup
 */
function runSystemTest() {
  const ui = SpreadsheetApp.getUi();
  let report = 'SYSTEM TEST REPORT\n\n';

  // Test 1: Configuration
  try {
    const config = getAllConfig();
    report += '✅ Config Module: OK\n';
    report += `   - Config keys loaded: ${Object.keys(config).length}\n`;
  } catch (error) {
    report += '❌ Config Module: FAILED\n';
    report += `   - Error: ${error.message}\n`;
  }

  // Test 2: Recipient Manager
  try {
    const summary = getRecipientSummary();
    report += '✅ Recipient Manager: OK\n';
    report += `   - Total recipients: ${summary.total}\n`;
  } catch (error) {
    report += '⚠️ Recipient Manager: WARNING\n';
    report += `   - ${error.message}\n`;
  }

  // Test 3: Logger
  try {
    const logSheet = getLogSheet();
    const stats = getLogStats();
    report += '✅ Logger: OK\n';
    report += `   - Log entries: ${stats.total}\n`;
  } catch (error) {
    report += '❌ Logger: FAILED\n';
    report += `   - Error: ${error.message}\n`;
  }

  // Test 4: Template Parser
  try {
    const testData = { Name: 'Test', Email: 'test@example.com' };
    const result = replacePlaceholders('Hello {{Name}}', testData);
    if (result === 'Hello Test') {
      report += '✅ Template Parser: OK\n';
    } else {
      report += '❌ Template Parser: FAILED\n';
      report += `   - Expected "Hello Test", got "${result}"\n`;
    }
  } catch (error) {
    report += '❌ Template Parser: FAILED\n';
    report += `   - Error: ${error.message}\n`;
  }

  // Test 5: Email Validation
  try {
    const validEmail = isValidEmail('test@example.com');
    const invalidEmail = isValidEmail('invalid-email');
    if (validEmail && !invalidEmail) {
      report += '✅ Email Validation: OK\n';
    } else {
      report += '❌ Email Validation: FAILED\n';
    }
  } catch (error) {
    report += '❌ Email Validation: FAILED\n';
    report += `   - Error: ${error.message}\n`;
  }

  // Test 6: Quota Check
  try {
    const quota = getQuotaInfo();
    report += '✅ Quota Check: OK\n';
    report += `   - Remaining: ${quota.remaining}/${quota.total}\n`;
  } catch (error) {
    report += '❌ Quota Check: FAILED\n';
    report += `   - Error: ${error.message}\n`;
  }

  // Test 7: Configuration Validation
  try {
    const validation = validateConfig();
    if (validation.isValid) {
      report += '✅ Configuration: COMPLETE\n';
    } else {
      report += '⚠️ Configuration: INCOMPLETE\n';
      report += `   - Missing: ${validation.missing.join(', ')}\n`;
    }
  } catch (error) {
    report += '❌ Configuration Validation: FAILED\n';
    report += `   - Error: ${error.message}\n`;
  }

  // Test 8: Document Generator - Template Field Extraction
  try {
    // Note: This test requires a PDF template document to be configured
    const templateDocId = getConfig(CONFIG_KEYS.PDF_TEMPLATE_DOC_ID);
    if (templateDocId) {
      const fields = getTemplateFields(templateDocId);
      report += '✅ PDF Template Field Extraction: OK\n';
      report += `   - Required fields: ${fields.requiredFields.length > 0 ? fields.requiredFields.join(', ') : 'none'}\n`;
      report += `   - Optional fields: ${fields.optionalFields.length > 0 ? fields.optionalFields.join(', ') : 'none'}\n`;
    } else {
      report += '⚠️ PDF Template Field Extraction: SKIPPED\n';
      report += '   - No PDF template configured\n';
    }
  } catch (error) {
    report += '❌ PDF Template Field Extraction: FAILED\n';
    report += `   - Error: ${error.message}\n`;
  }

  // Test 9: Document Generator - Recipient Validation
  try {
    const templateDocId = getConfig(CONFIG_KEYS.PDF_TEMPLATE_DOC_ID);
    if (templateDocId) {
      // Test with complete data
      const completeData = { Name: 'Test User', Company: 'Test Corp', City: 'New York' };
      const validation1 = validateRecipientData(templateDocId, completeData);

      // Test with missing data
      const incompleteData = { Name: 'Test User' }; // Missing other fields
      const validation2 = validateRecipientData(templateDocId, incompleteData);

      if (validation1.isValid !== undefined && validation2.isValid !== undefined) {
        report += '✅ PDF Recipient Validation: OK\n';
        report += `   - Complete data: ${validation1.isValid ? 'Valid' : 'Invalid'}\n`;
        report += `   - Incomplete data: ${validation2.isValid ? 'Valid' : 'Invalid (expected)'}\n`;
        if (!validation2.isValid && validation2.missing.length > 0) {
          report += `   - Missing fields detected: ${validation2.missing.join(', ')}\n`;
        }
      } else {
        report += '❌ PDF Recipient Validation: FAILED\n';
      }
    } else {
      report += '⚠️ PDF Recipient Validation: SKIPPED\n';
      report += '   - No PDF template configured\n';
    }
  } catch (error) {
    report += '❌ PDF Recipient Validation: FAILED\n';
    report += `   - Error: ${error.message}\n`;
  }

  report += '\n---\n';
  report += 'System test complete. Review results above.\n';

  Logger.log(report);
  ui.alert('System Test', report, ui.ButtonSet.OK);

  return report;
}

/**
 * Create a sample recipient sheet for testing
 * Run this to set up a test environment
 */
function createSampleRecipientSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Recipients';

  // Check if sheet already exists
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Sheet Exists',
      `Sheet "${sheetName}" already exists. Replace it?`,
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      spreadsheet.deleteSheet(sheet);
      sheet = spreadsheet.insertSheet(sheetName);
    } else {
      return;
    }
  } else {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  // Add headers (Email Status will be added later by ensureRequiredColumns)
  const headers = ['Email', 'First Name', 'Last Name', 'Address1', 'Address2', 'City', 'State', 'ZIP'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  // Add sample data
  const sampleData = [
    ['john@example.com', 'John', 'Doe', '123 Main Street', 'Apt 4B', 'New York', 'NY', '10001', 'pending'],
    ['jane@example.com', 'Jane', 'Smith', '456 Oak Avenue', '', 'Los Angeles', 'CA', '90001', 'pending'],
    ['bob@example.com', 'Bob', 'Johnson', '789 Pine Road', 'Suite 200', 'Chicago', 'IL', '60601', 'pending']
  ];

  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);

  // Set column widths
  sheet.setColumnWidth(1, 200); // Email
  sheet.setColumnWidth(2, 120); // First Name
  sheet.setColumnWidth(3, 120); // Last Name
  sheet.setColumnWidth(4, 180); // Address1
  sheet.setColumnWidth(5, 100); // Address2
  sheet.setColumnWidth(6, 120); // City
  sheet.setColumnWidth(7, 60);  // State
  sheet.setColumnWidth(8, 80);  // ZIP
  sheet.setColumnWidth(9, 100); // Status

  // Freeze header row
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    'Success',
    'Sample recipient sheet created with 3 test recipients.\n\nNOTE: Please use valid test email addresses before sending!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
