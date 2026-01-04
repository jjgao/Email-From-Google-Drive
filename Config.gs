/**
 * Configuration Management
 * Handles getting and setting configuration values using Script Properties
 */

const CONFIG_KEYS = {
  TEMPLATE_DOC_ID: 'TEMPLATE_DOC_ID',
  RECIPIENT_SHEET_NAME: 'RECIPIENT_SHEET_NAME',
  LOG_SHEET_NAME: 'LOG_SHEET_NAME',
  SENDER_NAME: 'SENDER_NAME',
  REPLY_TO_EMAIL: 'REPLY_TO_EMAIL',
  EMAIL_SUBJECT: 'EMAIL_SUBJECT',
  TEST_EMAIL: 'TEST_EMAIL'
};

const DEFAULT_VALUES = {
  RECIPIENT_SHEET_NAME: 'Recipients',
  LOG_SHEET_NAME: 'Email Logs',
  EMAIL_SUBJECT: 'Message from {{Name}}'
};

const CONFIG_SHEET_NAME = 'Config';

const CONFIG_LABELS = {
  TEMPLATE_DOC_ID: 'Template Document ID',
  RECIPIENT_SHEET_NAME: 'Recipient Sheet Name',
  LOG_SHEET_NAME: 'Log Sheet Name',
  SENDER_NAME: 'Sender Name',
  REPLY_TO_EMAIL: 'Reply-To Email',
  EMAIL_SUBJECT: 'Email Subject',
  TEST_EMAIL: 'Test Email'
};

const CONFIG_DESCRIPTIONS = {
  TEMPLATE_DOC_ID: 'Google Doc ID for email template (from URL)',
  RECIPIENT_SHEET_NAME: 'Name of sheet containing recipients',
  LOG_SHEET_NAME: 'Name of sheet for email logs',
  SENDER_NAME: 'Display name that appears as sender',
  REPLY_TO_EMAIL: 'Email address for replies',
  EMAIL_SUBJECT: 'Subject line (can use {{placeholders}})',
  TEST_EMAIL: 'Your email address for testing'
};

/**
 * Get a configuration value directly from Config sheet
 * @param {string} key - Configuration key
 * @returns {string|null} Configuration value or null if not set
 */
function getConfig(key) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

    if (!sheet) {
      // Return default value if sheet doesn't exist
      return DEFAULT_VALUES[key] || null;
    }

    const data = sheet.getDataRange().getValues();
    const label = CONFIG_LABELS[key];

    // Skip instruction row and header row (start at index 2)
    for (let i = 2; i < data.length; i++) {
      if (data[i][0] === label) {
        const value = data[i][1];
        if (value && value.toString().trim() !== '') {
          return value.toString().trim();
        }
        break;
      }
    }

    // Return default value if not found in sheet
    return DEFAULT_VALUES[key] || null;

  } catch (error) {
    console.log(`Error reading config for ${key}: ${error.message}`);
    return DEFAULT_VALUES[key] || null;
  }
}

/**
 * Set a configuration value in Config sheet
 * @param {string} key - Configuration key
 * @param {string} value - Configuration value
 */
function setConfig(key, value) {
  try {
    const sheet = getConfigSheet();
    const data = sheet.getDataRange().getValues();
    const label = CONFIG_LABELS[key];

    // Find and update the row
    for (let i = 2; i < data.length; i++) {
      if (data[i][0] === label) {
        sheet.getRange(i + 1, 2).setValue(value);
        return;
      }
    }
  } catch (error) {
    throw new Error(`Failed to set config ${key}: ${error.message}`);
  }
}

/**
 * Set multiple configuration values at once
 * @param {Object} configObject - Object with key-value pairs
 */
function setMultipleConfig(configObject) {
  for (const key in configObject) {
    setConfig(key, configObject[key]);
  }
}

/**
 * Get all configuration values
 * @returns {Object} All configuration values with defaults applied
 */
function getAllConfig() {
  const config = {};

  for (const key in CONFIG_KEYS) {
    config[key] = getConfig(CONFIG_KEYS[key]);
  }

  return config;
}

/**
 * Check if required configuration is set
 * @returns {Object} Object with isValid boolean and missing array
 */
function validateConfig() {
  const required = [
    CONFIG_KEYS.TEMPLATE_DOC_ID,
    CONFIG_KEYS.SENDER_NAME,
    CONFIG_KEYS.REPLY_TO_EMAIL,
    CONFIG_KEYS.TEST_EMAIL
  ];

  const missing = [];

  for (const key of required) {
    if (!getConfig(key)) {
      missing.push(key);
    }
  }

  return {
    isValid: missing.length === 0,
    missing: missing
  };
}

/**
 * Clear all configuration (useful for testing)
 */
function clearAllConfig() {
  const sheet = getConfigSheet();
  const data = sheet.getDataRange().getValues();

  // Clear all values in the Value column (skip instruction and header rows)
  for (let i = 2; i < data.length; i++) {
    sheet.getRange(i + 1, 2).setValue('');
  }
}

/**
 * Get or create the config sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Config sheet
 */
function getConfigSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG_SHEET_NAME);
    initializeConfigSheet(sheet);
  }

  return sheet;
}

/**
 * Initialize config sheet with headers and current values
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Config sheet
 */
function initializeConfigSheet(sheet) {
  // Set up headers
  const headers = ['Setting', 'Value', 'Description'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 200); // Setting
  sheet.setColumnWidth(2, 300); // Value
  sheet.setColumnWidth(3, 350); // Description

  // Add all config settings
  const data = [];
  let rowIndex = 2;

  for (const key in CONFIG_KEYS) {
    const configKey = CONFIG_KEYS[key];
    const label = CONFIG_LABELS[configKey] || configKey;
    const description = CONFIG_DESCRIPTIONS[configKey] || '';
    const currentValue = getConfig(configKey) || DEFAULT_VALUES[configKey] || '';

    data.push([label, currentValue, description]);
  }

  sheet.getRange(2, 1, data.length, 3).setValues(data);

  // Add instructions at the top
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, 3).merge();
  sheet.getRange(1, 1).setValue('📋 Email Campaign Configuration - Edit values in the "Value" column below. Changes take effect immediately.');
  sheet.getRange(1, 1).setBackground('#fff3cd');
  sheet.getRange(1, 1).setFontSize(10);
  sheet.getRange(1, 1).setWrap(true);
  sheet.setRowHeight(1, 40);

  // Protect the Setting and Description columns
  const settingRange = sheet.getRange(3, 1, data.length, 1);
  const descRange = sheet.getRange(3, 3, data.length, 1);

  const protection1 = settingRange.protect();
  protection1.setDescription('Setting names (read-only)');
  protection1.setWarningOnly(true);

  const protection2 = descRange.protect();
  protection2.setDescription('Descriptions (read-only)');
  protection2.setWarningOnly(true);

  // Highlight required fields
  const requiredKeys = [
    CONFIG_KEYS.TEMPLATE_DOC_ID,
    CONFIG_KEYS.SENDER_NAME,
    CONFIG_KEYS.REPLY_TO_EMAIL,
    CONFIG_KEYS.TEST_EMAIL
  ];

  for (let i = 0; i < data.length; i++) {
    const key = Object.keys(CONFIG_KEYS)[i];
    const configKey = CONFIG_KEYS[key];

    if (requiredKeys.includes(configKey)) {
      sheet.getRange(i + 3, 1).setBackground('#ffe6e6'); // Light red for required
    }
  }
}


/**
 * Create or update config sheet with current values
 */
function createOrUpdateConfigSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

  if (sheet) {
    // Update existing sheet
    spreadsheet.deleteSheet(sheet);
  }

  sheet = spreadsheet.insertSheet(CONFIG_SHEET_NAME);
  initializeConfigSheet(sheet);

  // Move to first position
  spreadsheet.moveActiveSheet(1);

  return sheet;
}
