/**
 * User Interface
 * Custom menu and dialogs for user interaction
 */

/**
 * Create custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Email Campaign')
    .addItem('⚙️ Open Config Sheet', 'openConfigSheetUI')
    .addItem('📋 Create Sample Data', 'createSampleDataUI')
    .addSeparator()
    .addItem('📧 Send Test Email', 'sendTestEmailUI')
    .addItem('🚀 Send Campaign', 'sendCampaignUI')
    .addSeparator()
    .addItem('📊 View Logs', 'openLogSheet')
    .addItem('📈 View Statistics', 'showStatsDialog')
    .addSeparator()
    .addSubMenu(ui.createMenu('Advanced')
      .addItem('Setup via Dialog (Legacy)', 'showConfigDialog')
      .addItem('Clear Configuration', 'clearConfigUI')
      .addItem('Ensure Status Column', 'ensureStatusColumnUI'))
    .addToUi();
}

/**
 * Show configuration dialog
 */
function showConfigDialog() {
  const ui = SpreadsheetApp.getUi();
  const currentConfig = getAllConfig();

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      .field { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input { width: 100%; padding: 8px; box-sizing: border-box; }
      .hint { font-size: 11px; color: #666; margin-top: 3px; }
      button { background-color: #4285f4; color: white; padding: 10px 20px; border: none; cursor: pointer; margin-top: 10px; }
      button:hover { background-color: #357ae8; }
    </style>

    <div class="field">
      <label>Template Document ID:</label>
      <input type="text" id="templateDocId" value="${currentConfig.TEMPLATE_DOC_ID || ''}" />
      <div class="hint">Find in URL: docs.google.com/document/d/[DOCUMENT_ID]/edit</div>
    </div>

    <div class="field">
      <label>Sender Name:</label>
      <input type="text" id="senderName" value="${currentConfig.SENDER_NAME || ''}" />
      <div class="hint">Name that appears as the sender</div>
    </div>

    <div class="field">
      <label>Reply-To Email:</label>
      <input type="text" id="replyToEmail" value="${currentConfig.REPLY_TO_EMAIL || ''}" />
      <div class="hint">Email address for replies</div>
    </div>

    <div class="field">
      <label>Test Email:</label>
      <input type="text" id="testEmail" value="${currentConfig.TEST_EMAIL || ''}" />
      <div class="hint">Your email for testing</div>
    </div>

    <div class="field">
      <label>Email Subject:</label>
      <input type="text" id="emailSubject" value="${currentConfig.EMAIL_SUBJECT || 'Message from {{Name}}'}" />
      <div class="hint">Can use {{placeholders}}</div>
    </div>

    <div class="field">
      <label>Recipient Sheet Name:</label>
      <input type="text" id="recipientSheet" value="${currentConfig.RECIPIENT_SHEET_NAME || 'Recipients'}" />
    </div>

    <button onclick="saveConfig()">Save Configuration</button>

    <script>
      function saveConfig() {
        const config = {
          TEMPLATE_DOC_ID: document.getElementById('templateDocId').value,
          SENDER_NAME: document.getElementById('senderName').value,
          REPLY_TO_EMAIL: document.getElementById('replyToEmail').value,
          TEST_EMAIL: document.getElementById('testEmail').value,
          EMAIL_SUBJECT: document.getElementById('emailSubject').value,
          RECIPIENT_SHEET_NAME: document.getElementById('recipientSheet').value
        };

        google.script.run
          .withSuccessHandler(function() {
            alert('Configuration saved successfully!');
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            alert('Error saving configuration: ' + error.message);
          })
          .saveConfiguration(config);
      }
    </script>
  `)
    .setWidth(500)
    .setHeight(550);

  ui.showModalDialog(html, 'Email Campaign Configuration');
}

/**
 * Save configuration from dialog
 * @param {Object} config - Configuration object
 */
function saveConfiguration(config) {
  setMultipleConfig(config);
}

/**
 * Send test email from menu
 */
function sendTestEmailUI() {
  const ui = SpreadsheetApp.getUi();

  // Check configuration first
  const validation = validateConfig();
  if (!validation.isValid) {
    ui.alert('Configuration Required', `Please configure the following:\n\n${validation.missing.join('\n')}`, ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    'Send Test Email',
    `This will send a test email to: ${getConfig(CONFIG_KEYS.TEST_EMAIL)}\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    try {
      const result = sendTestEmail();

      if (result.success) {
        ui.alert('Success', result.message, ui.ButtonSet.OK);
      } else {
        ui.alert('Error', `Failed to send test email:\n\n${result.error}`, ui.ButtonSet.OK);
      }
    } catch (error) {
      ui.alert('Error', `Unexpected error:\n\n${error.message}`, ui.ButtonSet.OK);
    }
  }
}

/**
 * Send campaign from menu
 */
function sendCampaignUI() {
  const ui = SpreadsheetApp.getUi();

  // Check configuration first
  const validation = validateConfig();
  if (!validation.isValid) {
    ui.alert('Configuration Required', `Please configure the following:\n\n${validation.missing.join('\n')}`, ui.ButtonSet.OK);
    return;
  }

  // Get pending recipient count
  try {
    const summary = getRecipientSummary();

    if (summary.pending === 0) {
      ui.alert('No Pending Recipients', 'All emails have already been sent or there are no recipients.', ui.ButtonSet.OK);
      return;
    }

    // Check quota
    const quota = getQuotaInfo();
    if (quota.remaining < summary.pending) {
      ui.alert(
        'Quota Warning',
        `You have ${quota.remaining} emails remaining in your daily quota, but ${summary.pending} emails to send.\n\nPlease reduce recipients or wait until tomorrow.`,
        ui.ButtonSet.OK
      );
      return;
    }

    // Confirm send
    const response = ui.alert(
      'Confirm Send Campaign',
      `Ready to send to ${summary.pending} recipients.\n\nQuota remaining: ${quota.remaining}\n\nThis action cannot be undone. Continue?`,
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      // Show progress message
      SpreadsheetApp.getActiveSpreadsheet().toast('Sending emails...', 'Campaign Progress', -1);

      const results = sendCampaign();

      // Log summary
      logCampaignSummary(results);

      // Hide progress
      SpreadsheetApp.getActiveSpreadsheet().toast('Complete!', 'Campaign Progress', 3);

      // Show results
      let message = `Campaign Complete!\n\n`;
      message += `Total: ${results.total}\n`;
      message += `Success: ${results.success}\n`;
      message += `Failed: ${results.failed}\n`;

      if (results.failed > 0) {
        message += `\nFirst few errors:\n`;
        const errorPreview = results.errors.slice(0, 3);
        errorPreview.forEach(err => {
          message += `- ${err.email}: ${err.error}\n`;
        });

        if (results.errors.length > 3) {
          message += `\n(and ${results.errors.length - 3} more - check logs)`;
        }
      }

      ui.alert('Campaign Results', message, ui.ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Error occurred', 'Campaign Progress', 3);
    ui.alert('Error', `Campaign failed:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Open log sheet
 */
function openLogSheet() {
  const sheet = getLogSheet();
  SpreadsheetApp.setActiveSheet(sheet);
}

/**
 * Show statistics dialog
 */
function showStatsDialog() {
  const ui = SpreadsheetApp.getUi();

  try {
    const recipientStats = getRecipientSummary();
    const logStats = getLogStats();
    const quota = getQuotaInfo();

    let message = 'RECIPIENT STATUS\n';
    message += `Total Recipients: ${recipientStats.total}\n`;
    message += `Pending: ${recipientStats.pending}\n`;
    message += `Sent: ${recipientStats.sent}\n`;
    message += `Failed: ${recipientStats.failed}\n\n`;

    message += 'LOG STATISTICS\n';
    message += `Total Log Entries: ${logStats.total}\n`;
    message += `Sent Emails: ${logStats.sent}\n`;
    message += `Failed Emails: ${logStats.failed}\n`;
    message += `Test Emails: ${logStats.test}\n\n`;

    message += 'QUOTA INFORMATION\n';
    message += `Remaining Today: ${quota.remaining}\n`;
    message += `Estimated Total: ${quota.total}`;

    ui.alert('Email Campaign Statistics', message, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('Error', `Failed to get statistics:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Clear configuration (Advanced)
 */
function clearConfigUI() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Clear Configuration',
    'This will delete all saved configuration. Are you sure?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    clearAllConfig();
    ui.alert('Success', 'Configuration cleared.', ui.ButtonSet.OK);
  }
}

/**
 * Ensure status column exists (Advanced)
 */
function ensureStatusColumnUI() {
  try {
    ensureStatusColumn();
    SpreadsheetApp.getUi().alert('Success', 'Status column has been added or already exists.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Open or create the config sheet
 */
function openConfigSheetUI() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

    if (!sheet) {
      // Create new config sheet
      sheet = createOrUpdateConfigSheet();
      SpreadsheetApp.getUi().alert(
        'Config Sheet Created',
        'A new configuration sheet has been created.\n\n' +
        'Edit the values in the "Value" column. Changes take effect immediately.\n\n' +
        'Required fields are highlighted in light red.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      // Just open existing sheet
      spreadsheet.setActiveSheet(sheet);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', `Failed to open config sheet:\n\n${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Create sample data in the spreadsheet
 */
function createSampleDataUI() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const recipientSheetName = getConfig(CONFIG_KEYS.RECIPIENT_SHEET_NAME);

  const response = ui.alert(
    'Create Sample Data',
    `This will create sample recipient data in a sheet named "${recipientSheetName}".\n\n` +
    'If the sheet already exists, you will be asked to confirm replacement.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    // Check if sheet already exists
    let sheet = spreadsheet.getSheetByName(recipientSheetName);

    if (sheet) {
      const replaceResponse = ui.alert(
        'Sheet Exists',
        `Sheet "${recipientSheetName}" already exists. Replace it with sample data?`,
        ui.ButtonSet.YES_NO
      );

      if (replaceResponse === ui.Button.YES) {
        spreadsheet.deleteSheet(sheet);
        sheet = spreadsheet.insertSheet(recipientSheetName);
      } else {
        return;
      }
    } else {
      sheet = spreadsheet.insertSheet(recipientSheetName);
    }

    // Add headers
    const headers = ['Email', 'Name', 'Company', 'Status'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');

    // Add sample data
    const sampleData = [
      ['your-email@example.com', 'John Doe', 'Acme Inc', 'pending'],
      ['another-email@example.com', 'Jane Smith', 'Tech Corp', 'pending'],
      ['third-email@example.com', 'Bob Johnson', 'StartupXYZ', 'pending']
    ];

    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);

    // Set column widths
    sheet.setColumnWidth(1, 220); // Email
    sheet.setColumnWidth(2, 150); // Name
    sheet.setColumnWidth(3, 150); // Company
    sheet.setColumnWidth(4, 100); // Status

    // Freeze header row
    sheet.setFrozenRows(1);

    // Activate the new sheet
    spreadsheet.setActiveSheet(sheet);

    ui.alert(
      'Success',
      `Sample recipient sheet "${recipientSheetName}" created with 3 test recipients.\n\n` +
      'IMPORTANT: Please replace the sample email addresses with valid test emails before sending!\n\n' +
      'You can edit the data directly in the sheet.',
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('Error', `Failed to create sample data:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}
