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
 * Get a configuration value
 * @param {string} key - Configuration key
 * @returns {string|null} Configuration value or null if not set
 */
function getConfig(key) {
  const properties = PropertiesService.getScriptProperties();
  let value = properties.getProperty(key);

  // Return default value if not set and default exists
  if (!value && DEFAULT_VALUES[key]) {
    return DEFAULT_VALUES[key];
  }

  return value;
}

/**
 * Set a configuration value
 * @param {string} key - Configuration key
 * @param {string} value - Configuration value
 */
function setConfig(key, value) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(key, value);
}

/**
 * Set multiple configuration values at once
 * @param {Object} configObject - Object with key-value pairs
 */
function setMultipleConfig(configObject) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties(configObject);
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
  const properties = PropertiesService.getScriptProperties();
  properties.deleteAllProperties();
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
  sheet.getRange(1, 1).setValue('📋 Email Campaign Configuration - Edit values below, then use "Email Campaign → Reload Configuration" to apply changes');
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
 * Load configuration from sheet to Script Properties
 * @returns {Object} Result with success boolean and message
 */
function loadConfigFromSheet() {
  try {
    const sheet = getConfigSheet();
    const data = sheet.getDataRange().getValues();

    // Skip instruction row and header row
    const configData = {};

    for (let i = 2; i < data.length; i++) {
      const label = data[i][0];
      const value = data[i][1];

      // Find the config key for this label
      for (const key in CONFIG_KEYS) {
        const configKey = CONFIG_KEYS[key];
        if (CONFIG_LABELS[configKey] === label) {
          if (value && value.toString().trim() !== '') {
            configData[configKey] = value.toString().trim();
          }
          break;
        }
      }
    }

    // Save to Script Properties
    if (Object.keys(configData).length > 0) {
      setMultipleConfig(configData);
    }

    return {
      success: true,
      message: `Configuration loaded successfully. ${Object.keys(configData).length} settings updated.`
    };

  } catch (error) {
    return {
      success: false,
      message: `Failed to load configuration: ${error.message}`
    };
  }
}

/**
 * Save configuration from Script Properties to sheet
 * @returns {Object} Result with success boolean and message
 */
function saveConfigToSheet() {
  try {
    const sheet = getConfigSheet();
    const currentConfig = getAllConfig();
    const data = sheet.getDataRange().getValues();

    // Update values in sheet (skip instruction and header rows)
    for (let i = 2; i < data.length; i++) {
      const label = data[i][0];

      // Find the config key for this label
      for (const key in CONFIG_KEYS) {
        const configKey = CONFIG_KEYS[key];
        if (CONFIG_LABELS[configKey] === label) {
          const value = currentConfig[key] || DEFAULT_VALUES[configKey] || '';
          sheet.getRange(i + 1, 2).setValue(value);
          break;
        }
      }
    }

    return {
      success: true,
      message: 'Configuration saved to sheet successfully.'
    };

  } catch (error) {
    return {
      success: false,
      message: `Failed to save configuration: ${error.message}`
    };
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
