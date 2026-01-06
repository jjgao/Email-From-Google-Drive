/**
 * Logger
 * Handles logging email activities to Google Sheets
 */

/**
 * Get or create the log sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Log sheet
 */
function getLogSheet() {
  const logSheetName = getConfig(CONFIG_KEYS.LOG_SHEET_NAME);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(logSheetName);

  // Create log sheet if it doesn't exist
  if (!sheet) {
    sheet = spreadsheet.insertSheet(logSheetName);
    initializeLogSheet(sheet);
  }

  return sheet;
}

/**
 * Initialize log sheet with headers
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Log sheet
 */
function initializeLogSheet(sheet) {
  const headers = ['Timestamp', 'Email', 'Name', 'Status', 'Error Message'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 150); // Timestamp
  sheet.setColumnWidth(2, 200); // Email
  sheet.setColumnWidth(3, 150); // Name
  sheet.setColumnWidth(4, 100); // Status
  sheet.setColumnWidth(5, 300); // Error Message
}

/**
 * Log an email send attempt
 * @param {string} email - Recipient email address
 * @param {string} name - Recipient name
 * @param {string} status - Status (sent, failed, test)
 * @param {string} errorMessage - Error message if failed
 */
function logEmail(email, name, status, errorMessage = '') {
  try {
    const sheet = getLogSheet();
    const timestamp = new Date();

    // Add new row at the top (after header)
    sheet.insertRowAfter(1);

    // Write log entry
    const logData = [
      timestamp,
      email,
      name,
      status,
      errorMessage
    ];

    sheet.getRange(2, 1, 1, logData.length).setValues([logData]);

    // Apply color coding based on status
    const statusCell = sheet.getRange(2, 4);
    if (status === 'sent') {
      statusCell.setBackground('#d4edda');
      statusCell.setFontColor('#155724');
    } else if (status === 'failed') {
      statusCell.setBackground('#f8d7da');
      statusCell.setFontColor('#721c24');
    } else if (status === 'test') {
      statusCell.setBackground('#d1ecf1');
      statusCell.setFontColor('#0c5460');
    }

  } catch (error) {
    // Log to console if sheet logging fails
    console.error('Failed to log email:', error.message);
  }
}

/**
 * Log campaign summary
 * @param {Object} summary - Campaign summary object
 */
function logCampaignSummary(summary) {
  try {
    const sheet = getLogSheet();
    const timestamp = new Date();

    // Add separator row
    sheet.insertRowAfter(1);
    const separatorData = [
      timestamp,
      '--- CAMPAIGN SUMMARY ---',
      '',
      '',
      `Total: ${summary.total} | Success: ${summary.success} | Failed: ${summary.failed}`
    ];

    sheet.getRange(2, 1, 1, separatorData.length).setValues([separatorData]);

    // Format summary row
    const summaryRange = sheet.getRange(2, 1, 1, separatorData.length);
    summaryRange.setFontWeight('bold');
    summaryRange.setBackground('#f0f0f0');

  } catch (error) {
    console.error('Failed to log campaign summary:', error.message);
  }
}

/**
 * Clear all logs (useful for testing)
 */
function clearLogs() {
  const sheet = getLogSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
}

/**
 * Get log statistics
 * @returns {Object} Statistics with counts by status
 */
function getLogStats() {
  const sheet = getLogSheet();
  const data = sheet.getDataRange().getValues();

  const stats = {
    total: 0,
    sent: 0,
    failed: 0,
    test: 0
  };

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const status = data[i][3]; // Status column

    if (status && !status.toString().startsWith('---')) {
      stats.total++;

      if (status === 'sent') {
        stats.sent++;
      } else if (status === 'failed') {
        stats.failed++;
      } else if (status === 'test') {
        stats.test++;
      }
    }
  }

  return stats;
}

/**
 * Export logs as CSV
 * @returns {string} CSV content
 */
function exportLogsAsCSV() {
  const sheet = getLogSheet();
  const data = sheet.getDataRange().getValues();

  let csv = '';
  for (const row of data) {
    csv += row.map(cell => `"${cell}"`).join(',') + '\n';
  }

  return csv;
}
