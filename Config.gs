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
