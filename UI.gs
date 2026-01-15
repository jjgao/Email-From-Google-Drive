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
    // Top Level - Document & PDF Generation
    .addItem('📑 Create All Documents', 'createAllDocumentsUI')
    .addItem('📕 Generate All PDFs', 'generateAllPdfsUI')
    .addItem('🚀 Send Campaign', 'sendCampaignUI')
    .addSeparator()
    // Reports & Monitoring
    .addItem('📊 View Logs', 'openLogSheet')
    .addItem('📈 View Statistics', 'showStatsDialog')
    .addSeparator()
    // Setup Submenu
    .addSubMenu(ui.createMenu('Setup')
      .addItem('⚙️ Open Config Sheet', 'openConfigSheetUI')
      .addItem('📋 Create Sample Data', 'createSampleDataUI')
      .addItem('📧 Create Sample Email Template', 'createSampleEmailTemplateUI')
      .addItem('📄 Create Sample PDF Template', 'createSamplePdfTemplateUI')
      .addSeparator()
      .addItem('🔧 Ensure Required Columns', 'ensureRequiredColumnsUI'))
    // Testing Submenu
    .addSubMenu(ui.createMenu('Testing')
      .addItem('📧 Send Test Email', 'sendTestEmailUI')
      .addItem('📄 Create Document (First Row)', 'createFirstDocumentUI')
      .addItem('📕 Generate PDF (First Row)', 'generateFirstPdfUI'))
    // Maintenance Submenu
    .addSubMenu(ui.createMenu('Maintenance')
      .addItem('✏️ Fill Default Filenames', 'fillDefaultFilenamesUI')
      .addItem('🔄 Regenerate All Documents', 'regenerateAllDocumentsUI')
      .addItem('🔄 Regenerate All PDFs', 'regenerateAllPdfsUI')
      .addSeparator()
      .addSubMenu(ui.createMenu('Advanced')
        .addItem('🔍 Preview Orphan Files', 'previewOrphanFilesUI')
        .addItem('🗑️ Delete Orphan Files', 'deleteOrphanFilesUI')
        .addItem('🐛 Debug Orphan Files', 'debugOrphanFilesUI')
        .addSeparator()
        .addItem('Setup via Dialog (Legacy)', 'showConfigDialog')
        .addItem('Clear Configuration', 'clearConfigUI')
        .addItem('Ensure Status Column', 'ensureStatusColumnUI')))
    .addToUi();
}

/**
 * Show sheet selection dialog and store the selection
 * @returns {string|null} Selected sheet name, or null if cancelled
 */
function showSheetSelectionDialog() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  // Filter out system sheets
  const systemSheets = ['Config', 'Email Logs'];
  const dataSheets = sheets.filter(sheet => !systemSheets.includes(sheet.getName()));

  if (dataSheets.length === 0) {
    ui.alert('No Data Sheets', 'No recipient data sheets found. Please create a sheet with recipient data first.', ui.ButtonSet.OK);
    return null;
  }

  // Build list of sheets with numbers
  let sheetList = 'Available sheets:\n';
  dataSheets.forEach((sheet, index) => {
    const isActive = sheet.getName() === spreadsheet.getActiveSheet().getName();
    sheetList += `${index + 1}. ${sheet.getName()}${isActive ? ' (currently active)' : ''}\n`;
  });

  // Get current selection or default
  const currentSelection = getSelectedSheetName();
  const activeSheetName = spreadsheet.getActiveSheet().getName();
  const defaultSheet = currentSelection || (systemSheets.includes(activeSheetName) ? dataSheets[0].getName() : activeSheetName);

  const response = ui.prompt(
    'Select Recipient Sheet',
    `${sheetList}\nEnter the sheet name or number (default: ${defaultSheet}):`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;
  }

  let selectedName = response.getResponseText().trim();

  // If empty, use default
  if (selectedName === '') {
    selectedName = defaultSheet;
  }

  // If numeric, convert to sheet name
  const sheetNumber = parseInt(selectedName, 10);
  if (!isNaN(sheetNumber) && sheetNumber >= 1 && sheetNumber <= dataSheets.length) {
    selectedName = dataSheets[sheetNumber - 1].getName();
  }

  // Validate sheet exists and is not a system sheet
  const selectedSheet = spreadsheet.getSheetByName(selectedName);
  if (!selectedSheet) {
    ui.alert('Invalid Sheet', `Sheet "${selectedName}" not found.`, ui.ButtonSet.OK);
    return null;
  }

  if (systemSheets.includes(selectedName)) {
    ui.alert('Invalid Sheet', `"${selectedName}" is a system sheet. Please select a data sheet.`, ui.ButtonSet.OK);
    return null;
  }

  // Store selection
  setSelectedSheetName(selectedName);

  return selectedName;
}

/**
 * Get the current recipient sheet name for display
 * @returns {string} Current sheet name being used
 */
function getCurrentRecipientSheetName() {
  const selectedName = getSelectedSheetName();
  if (selectedName) {
    return selectedName;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  const systemSheets = ['Config', 'Email Logs'];

  if (!systemSheets.includes(activeSheet.getName())) {
    return activeSheet.getName();
  }

  // Fall back to first data sheet
  const sheets = spreadsheet.getSheets();
  for (const sheet of sheets) {
    if (!systemSheets.includes(sheet.getName())) {
      return sheet.getName();
    }
  }

  return '(none)';
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
      <label>Email Template Document ID:</label>
      <input type="text" id="emailTemplateDocId" value="${currentConfig.EMAIL_TEMPLATE_DOC_ID || ''}" />
      <div class="hint">Template for email body - Find in URL: docs.google.com/document/d/[DOCUMENT_ID]/edit</div>
    </div>

    <div class="field">
      <label>PDF Template Document ID (Optional):</label>
      <input type="text" id="pdfTemplateDocId" value="${currentConfig.PDF_TEMPLATE_DOC_ID || ''}" />
      <div class="hint">Template for generating personalized PDFs - Leave empty if not using PDF generation</div>
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

    <button onclick="saveConfig()">Save Configuration</button>

    <script>
      function saveConfig() {
        const config = {
          EMAIL_TEMPLATE_DOC_ID: document.getElementById('emailTemplateDocId').value,
          PDF_TEMPLATE_DOC_ID: document.getElementById('pdfTemplateDocId').value,
          SENDER_NAME: document.getElementById('senderName').value,
          REPLY_TO_EMAIL: document.getElementById('replyToEmail').value,
          TEST_EMAIL: document.getElementById('testEmail').value,
          EMAIL_SUBJECT: document.getElementById('emailSubject').value
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

  // Get current sheet name for display
  const sheetName = getCurrentRecipientSheetName();

  // Get pending recipient count
  try {
    const summary = getRecipientSummary();

    if (summary.pending === 0) {
      ui.alert('No Pending Recipients', `Sheet: "${sheetName}"\n\nAll emails have already been sent or there are no recipients.`, ui.ButtonSet.OK);
      return;
    }

    // Check quota
    const quota = getQuotaInfo();
    if (quota.remaining < summary.pending) {
      ui.alert(
        'Quota Warning',
        `Sheet: "${sheetName}"\n\nYou have ${quota.remaining} emails remaining in your daily quota, but ${summary.pending} emails to send.\n\nPlease reduce recipients or wait until tomorrow.`,
        ui.ButtonSet.OK
      );
      return;
    }

    // Confirm send
    const response = ui.alert(
      'Confirm Send Campaign',
      `Sheet: "${sheetName}"\n` +
      `Recipients: ${summary.pending}\n` +
      `Quota remaining: ${quota.remaining}\n\n` +
      'This action cannot be undone. Continue?\n\n' +
      '(To change the sheet, cancel and use a different active sheet)',
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
    const sheetName = getCurrentRecipientSheetName();
    const recipientStats = getRecipientSummary();
    const logStats = getLogStats();
    const quota = getQuotaInfo();

    let message = `CURRENT SHEET: "${sheetName}"\n\n`;
    message += 'RECIPIENT STATUS\n';
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
 * Ensure all required columns exist (Filename, Status, Doc ID, PDF ID) and fill defaults (UI wrapper)
 */
function ensureRequiredColumnsUI() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Confirm before proceeding
    const response = ui.alert(
      'Ensure Required Columns',
      'This will:\n\n' +
      '1. Create missing columns: Filename, Doc ID, PDF ID, Email Status\n' +
      '2. Fill empty Filename cells with default names (Template Name - First Last)\n' +
      '3. Fill empty Email Status cells with "pending"\n\n' +
      'Existing values will NOT be overwritten.\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('Processing...', 'Ensuring Required Columns', -1);

    const results = ensureRequiredColumns();

    SpreadsheetApp.getActiveSpreadsheet().toast('Done!', 'Complete', 3);

    // Build result message
    let message = '✅ Required columns ensured!\n\n';

    if (results.columnsCreated.length > 0) {
      message += `📋 Columns created: ${results.columnsCreated.join(', ')}\n`;
    } else {
      message += '📋 All required columns already exist\n';
    }

    message += `\n📝 Filenames filled: ${results.filenamesFilled}`;
    message += `\n📊 Statuses set to "pending": ${results.statusesFilled}`;

    if (results.filenamesFilled === 0 && results.statusesFilled === 0 && results.columnsCreated.length === 0) {
      message = '✅ All required columns exist and all values are already filled.\n\nNo changes were made.';
    }

    ui.alert('Complete', message, ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed', 'Error', 3);
    ui.alert('Error', `Failed to ensure required columns:\n\n${error.message}`, ui.ButtonSet.OK);
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
  const sampleSheetName = 'Recipients';

  const response = ui.alert(
    'Create Sample Data',
    `This will create sample recipient data in a sheet named "${sampleSheetName}".\n\n` +
    'If the sheet already exists, you will be asked to confirm replacement.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    // Check if sheet already exists
    let sheet = spreadsheet.getSheetByName(sampleSheetName);

    if (sheet) {
      const replaceResponse = ui.alert(
        'Sheet Exists',
        `Sheet "${sampleSheetName}" already exists. Replace it with sample data?`,
        ui.ButtonSet.YES_NO
      );

      if (replaceResponse === ui.Button.YES) {
        spreadsheet.deleteSheet(sheet);
        sheet = spreadsheet.insertSheet(sampleSheetName);
      } else {
        return;
      }
    } else {
      sheet = spreadsheet.insertSheet(sampleSheetName);
    }

    // Add headers
    const headers = ['Email', 'First Name', 'Last Name', 'Address1', 'Address2', 'City', 'State', 'ZIP', 'Filename', 'Doc ID', 'PDF ID', 'Email Status', 'Attachment IDs'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');

    // Add sample data
    // Note: Third recipient has missing City field to test validation
    const sampleData = [
      ['your-email@example.com', 'John', 'Doe', '123 Main Street', 'Apt 4B', 'New York', 'NY', '10001', 'pending', '', '', ''],
      ['another-email@example.com', 'Jane', 'Smith', '456 Oak Avenue', '', 'Los Angeles', 'CA', '90001', 'pending', '', '', ''],
      ['third-email@example.com', 'Bob', 'Johnson', '789 Pine Road', 'Suite 200', '', 'IL', '60601', 'pending', '', '', '']
    ];

    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);

    // Set column widths
    sheet.setColumnWidth(1, 220); // Email
    sheet.setColumnWidth(2, 120); // First Name
    sheet.setColumnWidth(3, 120); // Last Name
    sheet.setColumnWidth(4, 200); // Address1
    sheet.setColumnWidth(5, 100); // Address2
    sheet.setColumnWidth(6, 120); // City
    sheet.setColumnWidth(7, 60);  // State
    sheet.setColumnWidth(8, 80);  // ZIP
    sheet.setColumnWidth(9, 100); // Status
    sheet.setColumnWidth(10, 300); // Doc ID
    sheet.setColumnWidth(11, 300); // PDF ID
    sheet.setColumnWidth(12, 350); // Attachment IDs

    // Freeze header row
    sheet.setFrozenRows(1);

    // Activate the new sheet
    spreadsheet.setActiveSheet(sheet);

    ui.alert(
      'Success',
      `Sample recipient sheet "${sampleSheetName}" created with 3 test recipients.\n\n` +
      'IMPORTANT: Please replace the sample email addresses with valid test emails before sending!\n\n' +
      'NOTE: The third recipient (Bob Johnson) has a missing City field to test validation. Document generation will skip this recipient if City is used in your template.\n\n' +
      'TIP: Use the "Filename" column to customize document names per recipient (overrides template config). Add a "Filename" column if you want custom naming.\n\n' +
      'TIP: Use the "Attachment IDs" column to attach additional files to emails (comma-separated Drive file IDs). PDFs are auto-attached when PDF ID exists.\n\n' +
      'Fields included: First Name, Last Name, Address1, Address2, City, State, ZIP\n\n' +
      'You can edit the data directly in the sheet.',
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('Error', `Failed to create sample data:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Create sample email template document
 */
function createSampleEmailTemplateUI() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Create Sample Email Template',
    'This will create a sample email template Google Doc.\n\n' +
    'The document will include:\n' +
    '• Example placeholder syntax\n' +
    '• Basic HTML formatting examples\n' +
    '• Template instructions\n\n' +
    'This template is used for email body content.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    // Create a new Google Doc
    const doc = DocumentApp.create('Email Campaign Template - Sample');
    const body = doc.getBody();

    // Clear any default content
    body.clear();

    // Add title
    const title = body.appendParagraph('Sample Email Template');
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);

    body.appendParagraph('');

    // Add formal letter address block
    const namePara = body.appendParagraph('');
    namePara.appendText('{{First Name}} {{Last Name}}');

    body.appendParagraph('{{Address1}}');

    const address2Para = body.appendParagraph('{{?Address2}}');

    const cityStateZip = body.appendParagraph('{{City}}, {{State}} {{ZIP}}');

    body.appendParagraph('');

    // Add greeting with placeholder
    const greetingPara = body.appendParagraph('');
    greetingPara.appendText('Dear {{First Name}},');

    body.appendParagraph('');

    // Add body paragraphs
    const para1 = body.appendParagraph(
      "Thank you for your interest in our services. We're excited to connect with you " +
      "and share some important updates!"
    );

    body.appendParagraph('');

    // Add list items as separate paragraphs
    body.appendParagraph("Here's what you can expect from us:");
    body.appendParagraph('• Personalized updates and news');
    body.appendParagraph('• Exclusive offers and promotions');
    body.appendParagraph('• Important announcements');

    body.appendParagraph('');

    // Add link example
    const linkPara = body.appendParagraph('Visit our website: ');
    const linkText = linkPara.appendText('https://example.com');
    linkText.setLinkUrl('https://example.com');
    linkText.setUnderline(true);
    linkText.setForegroundColor('#1155cc');

    body.appendParagraph('');

    // Add closing
    body.appendParagraph('Best regards,');
    body.appendParagraph('The Team');

    // Add separator
    body.appendParagraph('');
    const separator = body.appendParagraph('─────────────────────────────────');
    separator.setForegroundColor('#cccccc');

    // Add instructions section
    body.appendParagraph('');
    const instructionsTitle = body.appendParagraph('Template Instructions');
    instructionsTitle.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    instructionsTitle.setForegroundColor('#666666');

    const instructions = [
      '• Use double braces around field names for placeholders (example: First Name becomes { {First Name} } without the spaces)',
      '• Optional fields: Add a question mark for optional fields (example: { {?Address2} } without spaces) - empty values won\'t cause errors',
      '• Field names must match your spreadsheet column headers exactly (case-sensitive)',
      '• This template uses formal letter format with name and address at the top',
      '• Sample fields: First Name, Last Name, Address1, Address2 (optional), City, State, ZIP',
      '• You can add any custom fields by adding columns to your Recipients sheet',
      '• Basic formatting supported: bold, italic, underline, links',
      '• Keep formatting simple - complex layouts may not work',
      '• DELETE THESE INSTRUCTIONS before using the template - they contain placeholder examples that will cause errors'
    ];

    instructions.forEach(instruction => {
      const para = body.appendParagraph(instruction);
      para.setForegroundColor('#666666');
      para.setFontSize(10);
    });

    // Save the document
    doc.saveAndClose();

    // Get the document ID and URL
    const docId = doc.getId();
    const docUrl = doc.getUrl();

    // Ask if user wants to auto-populate Config sheet
    const configResponse = ui.alert(
      'Sample Template Created!',
      `Your sample template has been created.\n\n` +
      `Document ID: ${docId}\n\n` +
      'Would you like to automatically add this template ID to your Config sheet?',
      ui.ButtonSet.YES_NO
    );

    if (configResponse === ui.Button.YES) {
      // Set the email template ID in config
      setConfig(CONFIG_KEYS.EMAIL_TEMPLATE_DOC_ID, docId);

      ui.alert(
        'Success',
        'Sample email template created and configured!\n\n' +
        `Document ID has been added to your Config sheet.\n\n` +
        'The template will open in a new tab. You can edit it as needed.\n\n' +
        `Document URL: ${docUrl}`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Success',
        'Sample template created!\n\n' +
        `Document ID: ${docId}\n\n` +
        'Copy this ID to the "Template Document ID" field in your Config sheet.\n\n' +
        'The template will open in a new tab. You can edit it as needed.\n\n' +
        `Document URL: ${docUrl}`,
        ui.ButtonSet.OK
      );
    }

    // Open the document in a new tab
    const htmlOutput = HtmlService.createHtmlOutput(
      `<script>window.open('${docUrl}', '_blank'); google.script.host.close();</script>`
    ).setWidth(1).setHeight(1);

    ui.showModalDialog(htmlOutput, 'Opening Template...');

  } catch (error) {
    ui.alert('Error', `Failed to create sample template:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Create sample PDF template document
 */
function createSamplePdfTemplateUI() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Create Sample PDF Template',
    'This will create a sample PDF template Google Doc.\n\n' +
    'The document will include:\n' +
    '• Formal document format (certificate/letter style)\n' +
    '• Example placeholder syntax\n' +
    '• Professional layout\n' +
    '• Template instructions\n\n' +
    'This template is used for generating personalized PDF documents.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    // Create a new Google Doc
    const doc = DocumentApp.create('PDF Document Template - Sample');
    const body = doc.getBody();

    // Clear any default content
    body.clear();

    // Set page margins for formal document
    body.setMarginTop(72);    // 1 inch
    body.setMarginBottom(72);
    body.setMarginLeft(72);
    body.setMarginRight(72);

    // Add letterhead/header section
    const header = body.appendParagraph('YOUR COMPANY NAME');
    header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    header.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    const subheader = body.appendParagraph('Business Address • City, State ZIP • Phone • Email');
    subheader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    subheader.setFontSize(10);
    subheader.setForegroundColor('#666666');

    body.appendParagraph('');
    body.appendParagraph('');

    // Add horizontal line
    const line1 = body.appendParagraph('═══════════════════════════════════════════════════');
    line1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    line1.setForegroundColor('#4285f4');

    body.appendParagraph('');

    // Add document title
    const docTitle = body.appendParagraph('CERTIFICATE OF REGISTRATION');
    docTitle.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    docTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    docTitle.setBold(true);

    body.appendParagraph('');
    body.appendParagraph('');

    // Add date (example using current date, but you could use a placeholder)
    const datePara = body.appendParagraph('Date: [Current Date]');
    datePara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    datePara.setFontSize(10);

    body.appendParagraph('');

    // Add recipient information section
    const recipientSection = body.appendParagraph('This certifies that:');
    recipientSection.setBold(true);

    body.appendParagraph('');

    // Name with emphasis
    const namePara = body.appendParagraph('{{First Name}} {{Last Name}}');
    namePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    namePara.setFontSize(16);
    namePara.setBold(true);
    namePara.setForegroundColor('#4285f4');

    body.appendParagraph('');

    // Address block (centered)
    const address1Para = body.appendParagraph('{{Address1}}');
    address1Para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    address1Para.setFontSize(11);

    const address2Para = body.appendParagraph('{{?Address2}}');
    address2Para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    address2Para.setFontSize(11);

    const cityStateZipPara = body.appendParagraph('{{City}}, {{State}} {{ZIP}}');
    cityStateZipPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    cityStateZipPara.setFontSize(11);

    body.appendParagraph('');
    body.appendParagraph('');

    // Add body content
    const bodyText = body.appendParagraph(
      'has been successfully registered with our organization and is hereby granted ' +
      'full membership privileges effective immediately. This registration confirms ' +
      'compliance with all requirements and standards.'
    );
    bodyText.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    bodyText.setLineSpacing(1.5);

    body.appendParagraph('');
    body.appendParagraph('');

    // Add benefits/details section
    const benefitsTitle = body.appendParagraph('Membership Benefits:');
    benefitsTitle.setBold(true);

    body.appendParagraph('✓ Access to exclusive resources and materials');
    body.appendParagraph('✓ Priority support and assistance');
    body.appendParagraph('✓ Invitations to special events and webinars');
    body.appendParagraph('✓ Networking opportunities with peers');

    body.appendParagraph('');
    body.appendParagraph('');

    // Add contact info section
    const contactTitle = body.appendParagraph('Questions or Assistance:');
    contactTitle.setBold(true);

    const contactText = body.appendParagraph(
      'If you have any questions about your registration or need assistance, ' +
      'please don\'t hesitate to contact us. We\'re here to help!'
    );
    contactText.setFontSize(10);

    body.appendParagraph('');
    body.appendParagraph('');

    // Add signature line
    body.appendParagraph('_________________________________');
    body.appendParagraph('Authorized Signature');

    body.appendParagraph('');

    // Add footer line
    const line2 = body.appendParagraph('═══════════════════════════════════════════════════');
    line2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    line2.setForegroundColor('#4285f4');

    body.appendParagraph('');
    body.appendParagraph('');

    // Add separator before instructions
    const separator = body.appendParagraph('─────────────────────────────────');
    separator.setForegroundColor('#cccccc');
    separator.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    // Add instructions section
    body.appendParagraph('');
    const instructionsTitle = body.appendParagraph('Template Instructions');
    instructionsTitle.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    instructionsTitle.setForegroundColor('#666666');

    const instructions = [
      '• This is a PDF template - it generates personalized documents that get converted to PDF',
      '• Use double braces around field names for placeholders (example: First Name becomes { {First Name} } without the spaces)',
      '• Optional fields: Add a question mark for optional fields (example: { {?Address2} } without spaces) - empty values won\'t cause errors',
      '• Field names must match your spreadsheet column headers exactly (case-sensitive)',
      '• Sample fields: First Name, Last Name, Address1, Address2 (optional), City, State, ZIP',
      '• This template creates a formal certificate/letter format suitable for PDFs',
      '• You can add any custom fields by adding columns to your Recipients sheet',
      '• Customize the header, body content, and layout to match your needs',
      '• This is separate from your email template - email template is for email body, PDF template is for attached documents',
      '',
      'USAGE:',
      '1. Edit this template with your company information and desired content',
      '2. Copy the Document ID from the URL',
      '3. Add it to the "PDF Template Document ID" field in your Config sheet',
      '4. Use "Create All Documents" to generate personalized docs',
      '5. Use "Generate All PDFs" to convert docs to PDFs',
      '6. PDFs will automatically attach to emails when you send campaigns',
      '',
      'IMPORTANT: DELETE THESE INSTRUCTIONS before using the template!'
    ];

    instructions.forEach(instruction => {
      const para = body.appendParagraph(instruction);
      para.setForegroundColor('#666666');
      para.setFontSize(10);
    });

    // Save the document
    doc.saveAndClose();

    // Get the document ID and URL
    const docId = doc.getId();
    const docUrl = doc.getUrl();

    // Ask if user wants to auto-populate Config sheet
    const configResponse = ui.alert(
      'Sample PDF Template Created!',
      `Your sample PDF template has been created.\n\n` +
      `Document ID: ${docId}\n\n` +
      'Would you like to automatically add this template ID to your Config sheet?',
      ui.ButtonSet.YES_NO
    );

    if (configResponse === ui.Button.YES) {
      // Set the PDF template ID in config
      setConfig(CONFIG_KEYS.PDF_TEMPLATE_DOC_ID, docId);

      ui.alert(
        'Success',
        'Sample PDF template created and configured!\n\n' +
        `Document ID has been added to your Config sheet.\n\n` +
        'The template will open in a new tab. You can edit it as needed.\n\n' +
        `Document URL: ${docUrl}`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Success',
        'Sample PDF template created!\n\n' +
        `Document ID: ${docId}\n\n` +
        'Copy this ID to the "PDF Template Document ID" field in your Config sheet.\n\n' +
        'The template will open in a new tab. You can edit it as needed.\n\n' +
        `Document URL: ${docUrl}`,
        ui.ButtonSet.OK
      );
    }

    // Open the document in a new tab
    const htmlOutput = HtmlService.createHtmlOutput(
      `<script>window.open('${docUrl}', '_blank'); google.script.host.close();</script>`
    ).setWidth(1).setHeight(1);

    ui.showModalDialog(htmlOutput, 'Opening PDF Template...');

  } catch (error) {
    ui.alert('Error', `Failed to create sample PDF template:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Create a test document (UI wrapper)
 */
function createTestDocumentUI() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Create Test Document',
    'This will create a single test document using sample data.\n\n' +
    'The document will be saved in your configured output folder.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Creating test document...', 'Processing', -1);

    const result = createTestDocument();

    SpreadsheetApp.getActiveSpreadsheet().toast('Test document created successfully!', 'Success', 3);

    ui.alert(
      'Test Document Created',
      `Test document created successfully!\n\n` +
      `Document Name: ${result.docName}\n` +
      `Document ID: ${result.docId}\n\n` +
      'The document will open in a new tab.',
      ui.ButtonSet.OK
    );

    // Open the document in a new tab
    const htmlOutput = HtmlService.createHtmlOutput(
      `<script>window.open('${result.docUrl}', '_blank'); google.script.host.close();</script>`
    ).setWidth(1).setHeight(1);

    ui.showModalDialog(htmlOutput, 'Opening Document...');

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed to create test document', 'Error', 3);
    ui.alert('Error', `Failed to create test document:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Create all documents (UI wrapper)
 */
function createAllDocumentsUI() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Get current sheet name for display
    const sheetName = getCurrentRecipientSheetName();

    // Get pending count
    const pending = getRecipientsForDocumentCreation();

    if (pending.length === 0) {
      ui.alert(
        'No Pending Recipients',
        `Sheet: "${sheetName}"\n\n` +
        'All recipients either have documents created already.\n\n' +
        'To regenerate documents, clear the "Doc ID" column.',
        ui.ButtonSet.OK
      );
      return;
    }

    const response = ui.alert(
      'Create All Documents',
      `Sheet: "${sheetName}"\n` +
      `Recipients: ${pending.length}\n\n` +
      'Documents will be saved to your configured output folder.\n\n' +
      'This may take a few moments. Continue?\n\n' +
      '(To change the sheet, cancel and use a different active sheet)',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(`Creating ${pending.length} documents...`, 'Processing', -1);

    const results = createAllDocuments();

    SpreadsheetApp.getActiveSpreadsheet().toast('Document creation complete!', 'Success', 3);

    // Show results
    let message = `Documents Created: ${results.success}\n`;
    message += `Failed: ${results.failed}\n`;
    message += `Total: ${results.total}\n\n`;

    if (results.failed > 0) {
      message += 'Errors:\n';
      results.errors.forEach(err => {
        message += `• ${err.recipient}: ${err.error}\n`;
      });
    }

    message += '\nCheck the Email Logs sheet for details.';

    ui.alert('Document Creation Complete', message, ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed to create documents', 'Error', 3);
    ui.alert('Error', `Failed to create documents:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Generate PDFs for all documents (UI wrapper)
 */
function generateAllPdfsUI() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Get current sheet name for display
    const sheetName = getCurrentRecipientSheetName();

    // Get recipients needing PDFs
    const recipients = getRecipientsNeedingPdfs();

    if (recipients.length === 0) {
      ui.alert(
        'No Documents to Convert',
        `Sheet: "${sheetName}"\n\n` +
        'All recipients either have PDFs already or don\'t have documents yet.\n\n' +
        'To regenerate PDFs, clear the "PDF ID" column.',
        ui.ButtonSet.OK
      );
      return;
    }

    const response = ui.alert(
      'Generate PDFs',
      `Sheet: "${sheetName}"\n` +
      `Recipients: ${recipients.length}\n\n` +
      'PDFs will be saved to your configured output folder.\n\n' +
      'This may take a few moments. Continue?\n\n' +
      '(To change the sheet, cancel and use a different active sheet)',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(`Generating ${recipients.length} PDFs...`, 'Processing', -1);

    const results = generateAllPdfs();

    SpreadsheetApp.getActiveSpreadsheet().toast('PDF generation complete!', 'Success', 3);

    // Show results
    let message = `PDFs Generated: ${results.success}\n`;
    message += `Failed: ${results.failed}\n`;
    message += `Total: ${results.total}\n\n`;

    if (results.failed > 0) {
      message += 'Errors:\n';
      results.errors.forEach(err => {
        message += `• ${err.recipient}: ${err.error}\n`;
      });
    }

    message += '\nCheck the Email Logs sheet for details.';

    ui.alert('PDF Generation Complete', message, ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed to generate PDFs', 'Error', 3);
    ui.alert('Error', `Failed to generate PDFs:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Regenerate all documents (UI wrapper)
 */
function regenerateAllDocumentsUI() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Get all recipients count
    const allRecipients = getAllRecipientsFormatted();

    if (allRecipients.length === 0) {
      ui.alert(
        'No Recipients',
        'No recipients found in the sheet.',
        ui.ButtonSet.OK
      );
      return;
    }

    const response = ui.alert(
      'Regenerate All Documents',
      `⚠️ WARNING: This will regenerate documents for ALL ${allRecipients.length} recipient(s), overwriting existing documents.\n\n` +
      'Existing documents will be replaced with new ones.\n' +
      'PDF IDs will be cleared (PDFs will need to be regenerated).\n\n' +
      'This action cannot be undone. Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(`Regenerating ${allRecipients.length} documents...`, 'Processing', -1);

    const results = regenerateAllDocuments();

    SpreadsheetApp.getActiveSpreadsheet().toast('Document regeneration complete!', 'Success', 3);

    // Show results
    let message = `Documents Regenerated: ${results.success}\n`;
    message += `Failed: ${results.failed}\n`;
    message += `Total: ${results.total}\n\n`;

    if (results.failed > 0) {
      message += 'Errors:\n';
      results.errors.forEach(err => {
        message += `• ${err.recipient}: ${err.error}\n`;
      });
    }

    message += '\nCheck the Email Logs sheet for details.';

    ui.alert('Document Regeneration Complete', message, ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed to regenerate documents', 'Error', 3);
    ui.alert('Error', `Failed to regenerate documents:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Regenerate all PDFs (UI wrapper)
 */
function regenerateAllPdfsUI() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Get current sheet name for display
    const sheetName = getCurrentRecipientSheetName();

    // Get recipients with documents
    const allRecipients = getAllRecipientsFormatted();
    const recipientsWithDocs = allRecipients.filter(r => {
      const docId = (r.data['Doc ID'] || '').toString().trim();
      return docId !== '';
    });

    if (recipientsWithDocs.length === 0) {
      ui.alert(
        'No Documents Found',
        `Sheet: "${sheetName}"\n\nNo recipients have documents yet. Generate documents first.`,
        ui.ButtonSet.OK
      );
      return;
    }

    const response = ui.alert(
      'Regenerate All PDFs',
      `Sheet: "${sheetName}"\n` +
      `Recipients: ${recipientsWithDocs.length}\n\n` +
      '⚠️ WARNING: This will regenerate PDFs for ALL recipients with documents, overwriting existing PDFs.\n\n' +
      'This action cannot be undone. Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(`Regenerating ${recipientsWithDocs.length} PDFs...`, 'Processing', -1);

    const results = regenerateAllPdfs();

    SpreadsheetApp.getActiveSpreadsheet().toast('PDF regeneration complete!', 'Success', 3);

    // Show results
    let message = `PDFs Regenerated: ${results.success}\n`;
    message += `Failed: ${results.failed}\n`;
    message += `Total: ${results.total}\n\n`;

    if (results.failed > 0) {
      message += 'Errors:\n';
      results.errors.forEach(err => {
        message += `• ${err.recipient}: ${err.error}\n`;
      });
    }

    message += '\nCheck the Email Logs sheet for details.';

    ui.alert('PDF Regeneration Complete', message, ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed to regenerate PDFs', 'Error', 3);
    ui.alert('Error', `Failed to regenerate PDFs:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Preview orphan files without deleting (UI wrapper)
 */
function previewOrphanFilesUI() {
  const ui = SpreadsheetApp.getUi();

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Scanning folders for orphan files...', 'Processing', -1);

    const preview = previewAllOrphanFiles();
    const totalOrphans = preview.documents.orphanCount + preview.pdfs.orphanCount;
    const totalFiles = preview.documents.total + preview.pdfs.total;
    const totalProtected = preview.documents.protectedCount + preview.pdfs.protectedCount;

    SpreadsheetApp.getActiveSpreadsheet().toast('Scan complete', 'Done', 2);

    // Build preview message
    let message = '📋 ORPHAN FILES PREVIEW\n\n';

    // Show same folder mode notice
    if (preview.sameFolderMode) {
      message += 'ℹ️ SAME FOLDER MODE DETECTED\n';
      message += 'Using one folder for both documents and PDFs.\n';
      message += 'Files are protected if they appear in EITHER "Doc ID" OR "PDF ID" column.\n\n';
    }

    message += `📊 Summary:\n`;
    message += `• Total files in folder${preview.sameFolderMode ? '' : 's'}: ${totalFiles}\n`;
    message += `• Protected (linked to recipients): ${totalProtected}\n`;
    message += `• Orphan (not linked): ${totalOrphans}\n\n`;

    // Check if no orphans found
    if (totalOrphans === 0) {
      message += '✅ All files match recipients in the sheet.\n';
      message += 'No orphan files found - no cleanup needed.';
      ui.alert('No Orphan Files Found', message, ui.ButtonSet.OK);
      return;
    }

    // Warning if large percentage
    const deletePercent = (totalOrphans / totalFiles * 100).toFixed(0);
    if (deletePercent >= 50) {
      message += `\n⚠️ WARNING: ${deletePercent}% of all files are orphans!\n`;
      message += `This may indicate:\n`;
      message += `• Recipients are missing Doc IDs/PDF IDs\n`;
      message += `• Wrong folders are configured\n\n`;
    }

    // Show sample orphan files (up to 10)
    if (preview.documents.orphanFiles && preview.documents.orphanFiles.length > 0) {
      message += '\n📄 Orphan Documents:\n';
      const docSamples = preview.documents.orphanFiles.slice(0, 10);
      docSamples.forEach(file => {
        message += `• ${file.name}\n`;
      });
      if (preview.documents.orphanFiles.length > 10) {
        message += `... and ${preview.documents.orphanFiles.length - 10} more\n`;
      }
    }

    if (!preview.sameFolderMode && preview.pdfs.orphanFiles && preview.pdfs.orphanFiles.length > 0) {
      message += '\n📕 Orphan PDFs:\n';
      const pdfSamples = preview.pdfs.orphanFiles.slice(0, 10);
      pdfSamples.forEach(file => {
        message += `• ${file.name}\n`;
      });
      if (preview.pdfs.orphanFiles.length > 10) {
        message += `... and ${preview.pdfs.orphanFiles.length - 10} more\n`;
      }
    }

    message += '\n💡 Use "Delete Orphan Files" to remove these files.';

    ui.alert('Orphan Files Preview', message, ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed to preview orphan files', 'Error', 3);
    ui.alert('Error', `Failed to preview orphan files:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Delete orphan files (UI wrapper)
 */
function deleteOrphanFilesUI() {
  const ui = SpreadsheetApp.getUi();

  try {
    // First, preview what would be deleted
    SpreadsheetApp.getActiveSpreadsheet().toast('Scanning folders for orphan files...', 'Processing', -1);

    const preview = previewAllOrphanFiles();
    const totalOrphans = preview.documents.orphanCount + preview.pdfs.orphanCount;
    const totalFiles = preview.documents.total + preview.pdfs.total;
    const totalProtected = preview.documents.protectedCount + preview.pdfs.protectedCount;

    SpreadsheetApp.getActiveSpreadsheet().toast('Scan complete', 'Done', 2);

    // Check if no orphans found
    if (totalOrphans === 0) {
      ui.alert(
        'No Orphan Files Found',
        'All files in your output folders match recipients in the sheet.\n\n' +
        `📊 Summary:\n` +
        `• Total files: ${totalFiles}\n` +
        `• Protected files: ${totalProtected}\n` +
        `• Orphan files: 0\n\n` +
        `✅ No cleanup needed.`,
        ui.ButtonSet.OK
      );
      return;
    }

    // Build preview message
    let previewMessage = '⚠️ PREVIEW: The following files will be deleted:\n\n';

    // Show same folder mode notice
    if (preview.sameFolderMode) {
      previewMessage += 'ℹ️ SAME FOLDER MODE DETECTED\n';
      previewMessage += 'Using one folder for both documents and PDFs.\n';
      previewMessage += 'Files are protected if they appear in EITHER "Doc ID" OR "PDF ID" column.\n\n';
    }

    previewMessage += `📊 Summary:\n`;
    previewMessage += `• Total files in folder${preview.sameFolderMode ? '' : 's'}: ${totalFiles}\n`;
    previewMessage += `• Protected (linked to recipients): ${totalProtected}\n`;
    previewMessage += `• Orphan (will be deleted): ${totalOrphans}\n\n`;

    // Warning if deleting a large percentage
    const deletePercent = (totalOrphans / totalFiles * 100).toFixed(0);
    if (deletePercent >= 50) {
      previewMessage += `🚨 WARNING: This will delete ${deletePercent}% of all files!\n`;
      previewMessage += `⚠️ This may indicate a problem. Please verify:\n`;
      previewMessage += `• Recipients have Doc IDs/PDF IDs filled in\n`;
      previewMessage += `• These folders are only for this email campaign\n\n`;
    }

    // Warning if no recipient IDs found
    if (preview.documents.validDocIdsCount === 0 && preview.documents.total > 0) {
      previewMessage += `⚠️ WARNING: No recipients have Doc IDs!\n`;
      previewMessage += `This will delete ALL ${preview.documents.total} documents.\n`;
      previewMessage += `Make sure you've created documents first.\n\n`;
    }
    if (preview.pdfs.validPdfIdsCount === 0 && preview.pdfs.total > 0) {
      previewMessage += `⚠️ WARNING: No recipients have PDF IDs!\n`;
      previewMessage += `This will delete ALL ${preview.pdfs.total} PDFs.\n`;
      previewMessage += `Make sure you've generated PDFs first.\n\n`;
    }

    // List files to be deleted (limit to first 10 of each type)
    if (preview.documents.orphanCount > 0) {
      const label = preview.sameFolderMode ? '📄 Files to delete' : '📄 Documents to delete';
      previewMessage += `${label} (${preview.documents.orphanCount}):\n`;
      const docsToShow = preview.documents.orphanFiles.slice(0, 10);
      docsToShow.forEach(file => {
        previewMessage += `• ${file.name}\n`;
      });
      if (preview.documents.orphanCount > 10) {
        previewMessage += `... and ${preview.documents.orphanCount - 10} more\n`;
      }
      previewMessage += '\n';
    }

    if (preview.pdfs.orphanCount > 0 && !preview.sameFolderMode) {
      previewMessage += `📕 PDFs to delete (${preview.pdfs.orphanCount}):\n`;
      const pdfsToShow = preview.pdfs.orphanFiles.slice(0, 10);
      pdfsToShow.forEach(file => {
        previewMessage += `• ${file.name}\n`;
      });
      if (preview.pdfs.orphanCount > 10) {
        previewMessage += `... and ${preview.pdfs.orphanCount - 10} more\n`;
      }
      previewMessage += '\n';
    }

    previewMessage += 'Files will be moved to trash (recoverable for 30 days).\n\n';
    previewMessage += 'Proceed with deletion?';

    // Show preview and confirm
    const response = ui.alert(
      'Delete Orphan Files - Confirmation',
      previewMessage,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      ui.alert('Cancelled', 'No files were deleted.', ui.ButtonSet.OK);
      return;
    }

    // Proceed with deletion
    SpreadsheetApp.getActiveSpreadsheet().toast('Deleting orphan files...', 'Processing', -1);

    const results = deleteAllOrphanFiles();

    SpreadsheetApp.getActiveSpreadsheet().toast('Cleanup complete!', 'Success', 3);

    // Show results
    let message = `✅ Orphan files have been moved to trash:\n\n`;
    message += `Documents deleted: ${results.documents.deleted}\n`;
    message += `PDFs deleted: ${results.pdfs.deleted}\n\n`;

    if (results.documents.deleted > 0) {
      message += 'Deleted documents:\n';
      const docsToShow = results.documents.files.slice(0, 15);
      docsToShow.forEach(file => {
        message += `• ${file.name}\n`;
      });
      if (results.documents.deleted > 15) {
        message += `... and ${results.documents.deleted - 15} more\n`;
      }
      message += '\n';
    }

    if (results.pdfs.deleted > 0) {
      message += 'Deleted PDFs:\n';
      const pdfsToShow = results.pdfs.files.slice(0, 15);
      pdfsToShow.forEach(file => {
        message += `• ${file.name}\n`;
      });
      if (results.pdfs.deleted > 15) {
        message += `... and ${results.pdfs.deleted - 15} more\n`;
      }
    }

    message += '\nFiles can be recovered from Google Drive trash for 30 days.';

    ui.alert('Cleanup Complete', message, ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed to delete orphan files', 'Error', 3);
    ui.alert('Error', `Failed to delete orphan files:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Debug orphan files - show detailed file and recipient ID mapping
 */
function debugOrphanFilesUI() {
  const ui = SpreadsheetApp.getUi();

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Gathering diagnostic information...', 'Processing', -1);

    const debug = debugOrphanFiles();

    SpreadsheetApp.getActiveSpreadsheet().toast('Analysis complete', 'Done', 2);

    // Build diagnostic message
    let message = '🔍 ORPHAN FILES DIAGNOSTIC\n\n';

    // Recipients section
    message += `📋 RECIPIENTS (${debug.recipients.length}):\n`;
    debug.recipients.forEach((r, i) => {
      message += `${i + 1}. ${r.email}\n`;
      message += `   Doc ID: ${r.docId || '(empty)'}\n`;
      message += `   PDF ID: ${r.pdfId || '(empty)'}\n`;
    });
    message += '\n';

    // Document folder section
    message += `📄 DOCUMENT FOLDER (${debug.documentFolder.files.length} files):\n`;
    if (debug.documentFolder.error) {
      message += `   Error: ${debug.documentFolder.error}\n`;
    } else if (debug.documentFolder.files.length === 0) {
      message += '   (empty)\n';
    } else {
      debug.documentFolder.files.forEach((f, i) => {
        const isProtected = debug.recipients.some(r => r.docId === f.id);
        message += `${i + 1}. ${f.name}\n`;
        message += `   ID: ${f.id}\n`;
        message += `   Status: ${isProtected ? '✅ Protected' : '⚠️ ORPHAN'}\n`;
      });
    }
    message += '\n';

    // PDF folder section
    message += `📕 PDF FOLDER (${debug.pdfFolder.files.length} files):\n`;
    if (debug.pdfFolder.error) {
      message += `   Error: ${debug.pdfFolder.error}\n`;
    } else if (debug.pdfFolder.files.length === 0) {
      message += '   (empty)\n';
    } else {
      debug.pdfFolder.files.forEach((f, i) => {
        const isProtected = debug.recipients.some(r => r.pdfId === f.id);
        message += `${i + 1}. ${f.name}\n`;
        message += `   ID: ${f.id}\n`;
        message += `   Status: ${isProtected ? '✅ Protected' : '⚠️ ORPHAN'}\n`;
      });
    }

    // Summary
    const docOrphans = debug.documentFolder.files.filter(f =>
      !debug.recipients.some(r => r.docId === f.id)
    ).length;
    const pdfOrphans = debug.pdfFolder.files.filter(f =>
      !debug.recipients.some(r => r.pdfId === f.id)
    ).length;

    message += '\n📊 SUMMARY:\n';
    message += `• Document orphans: ${docOrphans}\n`;
    message += `• PDF orphans: ${pdfOrphans}\n`;
    message += `• Total orphans: ${docOrphans + pdfOrphans}\n`;

    // Log to console for copy/paste
    Logger.log(message);
    Logger.log('Full debug data:');
    Logger.log(JSON.stringify(debug, null, 2));

    ui.alert('Orphan Files Diagnostic', message, ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed to run diagnostic', 'Error', 3);
    ui.alert('Error', `Failed to run diagnostic:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Fill default filenames for all recipients (UI wrapper)
 */
function fillDefaultFilenamesUI() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Get all recipients
    const allRecipients = getAllRecipientsFormatted();

    if (allRecipients.length === 0) {
      ui.alert(
        'No Recipients',
        'No recipients found in the sheet.',
        ui.ButtonSet.OK
      );
      return;
    }

    const response = ui.alert(
      'Fill Default Filenames',
      `This will generate default filenames for all ${allRecipients.length} recipient(s).\n\n` +
      'Default format: "Template Name - First Last"\n' +
      'Example: "Sample PDF Template - John Doe"\n\n' +
      'If the "Filename" column doesn\'t exist, it will be created.\n' +
      'Existing custom filenames will be overwritten.\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('Generating default filenames...', 'Processing', -1);

    const results = fillDefaultFilenames();

    SpreadsheetApp.getActiveSpreadsheet().toast('Filenames filled!', 'Success', 3);

    // Show results
    let message = `✅ Default filenames have been generated:\n\n`;
    message += `Recipients processed: ${results.total}\n`;
    message += `Filenames filled: ${results.filled}\n\n`;

    if (results.sample && results.sample.length > 0) {
      message += 'Sample filenames:\n';
      const samplesToShow = results.sample.slice(0, 5);
      samplesToShow.forEach(s => {
        message += `• ${s.email}: "${s.filename}"\n`;
      });
      if (results.sample.length > 5) {
        message += `... and ${results.sample.length - 5} more\n`;
      }
    }

    ui.alert('Filenames Generated', message, ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed to fill filenames', 'Error', 3);
    ui.alert('Error', `Failed to fill filenames:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Create document for first row only (UI wrapper)
 */
function createFirstDocumentUI() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Get current sheet name for display
    const sheetName = getCurrentRecipientSheetName();

    // Get first recipient
    const allRecipients = getAllRecipientsFormatted();

    if (allRecipients.length === 0) {
      ui.alert(
        'No Recipients',
        `Sheet: "${sheetName}"\n\nNo recipients found in the sheet.`,
        ui.ButtonSet.OK
      );
      return;
    }

    const firstRecipient = allRecipients[0];
    const recipientEmail = firstRecipient.data.Email || '(no email)';

    // Check if document already exists
    if (firstRecipient.data['Doc ID'] && firstRecipient.data['Doc ID'].toString().trim() !== '') {
      const response = ui.alert(
        'Document Already Exists',
        `Sheet: "${sheetName}"\n` +
        `First recipient: ${recipientEmail}\n\n` +
        'This will create a new document and replace the existing Doc ID.\n\n' +
        'Continue?',
        ui.ButtonSet.YES_NO
      );

      if (response !== ui.Button.YES) {
        return;
      }
    } else {
      const response = ui.alert(
        'Create Document (First Row)',
        `Sheet: "${sheetName}"\n` +
        `First recipient: ${recipientEmail}\n\n` +
        'Continue?',
        ui.ButtonSet.YES_NO
      );

      if (response !== ui.Button.YES) {
        return;
      }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('Creating document...', 'Processing', -1);

    const result = createFirstDocument();

    SpreadsheetApp.getActiveSpreadsheet().toast('Document created!', 'Success', 3);

    // Show results
    let message = `✅ Document created successfully!\n\n`;
    message += `Recipient: ${result.email}\n`;
    message += `Document: ${result.docName}\n`;
    message += `Document ID: ${result.docId}\n\n`;
    message += 'The document has been saved to your output folder.';

    ui.alert('Document Created', message, ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed to create document', 'Error', 3);
    ui.alert('Error', `Failed to create document:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Generate PDF for first row only (UI wrapper)
 */
function generateFirstPdfUI() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Get current sheet name for display
    const sheetName = getCurrentRecipientSheetName();

    // Get first recipient
    const allRecipients = getAllRecipientsFormatted();

    if (allRecipients.length === 0) {
      ui.alert(
        'No Recipients',
        `Sheet: "${sheetName}"\n\nNo recipients found in the sheet.`,
        ui.ButtonSet.OK
      );
      return;
    }

    const firstRecipient = allRecipients[0];
    const recipientEmail = firstRecipient.data.Email || '(no email)';

    // Check if document exists
    if (!firstRecipient.data['Doc ID'] || firstRecipient.data['Doc ID'].toString().trim() === '') {
      ui.alert(
        'No Document Found',
        `Sheet: "${sheetName}"\n\nFirst recipient does not have a document yet.\n\n` +
        'Please create a document first using "Create Document (First Row)".',
        ui.ButtonSet.OK
      );
      return;
    }

    // Check if PDF already exists
    if (firstRecipient.data['PDF ID'] && firstRecipient.data['PDF ID'].toString().trim() !== '') {
      const response = ui.alert(
        'PDF Already Exists',
        `Sheet: "${sheetName}"\n` +
        `First recipient: ${recipientEmail}\n\n` +
        'This will create a new PDF and replace the existing PDF ID.\n\n' +
        'Continue?',
        ui.ButtonSet.YES_NO
      );

      if (response !== ui.Button.YES) {
        return;
      }
    } else {
      const response = ui.alert(
        'Generate PDF (First Row)',
        `Sheet: "${sheetName}"\n` +
        `First recipient: ${recipientEmail}\n\n` +
        'Continue?',
        ui.ButtonSet.YES_NO
      );

      if (response !== ui.Button.YES) {
        return;
      }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('Generating PDF...', 'Processing', -1);

    const result = generateFirstPdf();

    SpreadsheetApp.getActiveSpreadsheet().toast('PDF generated!', 'Success', 3);

    // Show results
    let message = `✅ PDF generated successfully!\n\n`;
    message += `Recipient: ${result.email}\n`;
    message += `PDF: ${result.pdfName}\n`;
    message += `PDF ID: ${result.pdfId}\n\n`;
    message += 'The PDF has been saved to your PDF folder.';

    ui.alert('PDF Generated', message, ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed to generate PDF', 'Error', 3);
    ui.alert('Error', `Failed to generate PDF:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}
