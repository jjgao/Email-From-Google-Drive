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
    .addItem('📝 Create Sample Template', 'createSampleTemplateUI')
    .addSeparator()
    .addItem('📄 Create Test Document', 'createTestDocumentUI')
    .addItem('📑 Create All Documents', 'createAllDocumentsUI')
    .addSeparator()
    .addItem('📕 Generate All PDFs', 'generateAllPdfsUI')
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
    const headers = ['Email', 'Name', 'Company', 'Street', 'City', 'State', 'ZIP', 'Status', 'Doc ID', 'PDF ID'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');

    // Add sample data
    const sampleData = [
      ['your-email@example.com', 'John Doe', 'Acme Inc', '123 Main Street', 'New York', 'NY', '10001', 'pending', '', ''],
      ['another-email@example.com', 'Jane Smith', 'Tech Corp', '456 Oak Avenue', 'San Francisco', 'CA', '94102', 'pending', '', ''],
      ['third-email@example.com', 'Bob Johnson', 'StartupXYZ', '789 Pine Road', 'Austin', 'TX', '73301', 'pending', '', '']
    ];

    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);

    // Set column widths
    sheet.setColumnWidth(1, 220); // Email
    sheet.setColumnWidth(2, 150); // Name
    sheet.setColumnWidth(3, 150); // Company
    sheet.setColumnWidth(4, 200); // Street
    sheet.setColumnWidth(5, 120); // City
    sheet.setColumnWidth(6, 60);  // State
    sheet.setColumnWidth(7, 80);  // ZIP
    sheet.setColumnWidth(8, 100); // Status
    sheet.setColumnWidth(9, 300); // Doc ID
    sheet.setColumnWidth(10, 300); // PDF ID

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

/**
 * Create sample email template document
 */
function createSampleTemplateUI() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Create Sample Template',
    'This will create a sample email template Google Doc.\n\n' +
    'The document will include:\n' +
    '• Example placeholder syntax\n' +
    '• Basic HTML formatting examples\n' +
    '• Template instructions\n\n' +
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
    body.appendParagraph('{{Name}}');
    body.appendParagraph('{{Street}}');
    const cityStateZip = body.appendParagraph('{{City}}, {{State}} {{ZIP}}');

    body.appendParagraph('');

    // Add greeting with placeholder
    body.appendParagraph('Dear {{Name}},');

    // Add body paragraphs
    const para1 = body.appendParagraph('Thank you for being a valued customer of ');
    para1.appendText('{{Company}}').setBold(true);
    para1.appendText(". We're excited to share some updates with you!");

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
    body.appendParagraph('The {{Company}} Team');

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
      '• Use {{FieldName}} for placeholders - field names must match your spreadsheet columns exactly',
      '• Available fields from sample data: {{Name}}, {{Company}}, {{Street}}, {{City}}, {{State}}, {{ZIP}}, {{Email}}',
      '• This template uses formal letter format with name and address at the top',
      '• You can add any custom fields by adding columns to your Recipients sheet',
      '• Basic formatting supported: bold, italic, underline, links',
      '• Keep formatting simple - complex layouts may not work',
      '• Delete these instructions before using the template'
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
      // Set the template ID in config
      setConfig(CONFIG_KEYS.TEMPLATE_DOC_ID, docId);

      ui.alert(
        'Success',
        'Sample template created and configured!\n\n' +
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
    // Get pending count
    const pending = getPendingRecipients();

    if (pending.length === 0) {
      ui.alert(
        'No Pending Recipients',
        'All recipients either have documents created already or have non-pending status.\n\n' +
        'To regenerate documents, clear the "Doc ID" column and set status back to "pending".',
        ui.ButtonSet.OK
      );
      return;
    }

    const response = ui.alert(
      'Create All Documents',
      `This will create ${pending.length} personalized document(s).\n\n` +
      'Documents will be saved to your configured output folder.\n\n' +
      'This may take a few moments. Continue?',
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
    // Check if PDF folder is configured
    const pdfFolderId = getConfig(CONFIG_KEYS.PDF_FOLDER_ID);
    if (!pdfFolderId) {
      ui.alert(
        'PDF Folder Not Configured',
        'Please configure the PDF Folder ID in the Config sheet before generating PDFs.\n\n' +
        'This is where your PDF files will be saved.',
        ui.ButtonSet.OK
      );
      return;
    }

    // Get recipients needing PDFs
    const recipients = getRecipientsNeedingPdfs();

    if (recipients.length === 0) {
      ui.alert(
        'No Documents to Convert',
        'All recipients either have PDFs already or don\'t have documents yet.\n\n' +
        'To regenerate PDFs, clear the "PDF ID" column.',
        ui.ButtonSet.OK
      );
      return;
    }

    const response = ui.alert(
      'Generate PDFs',
      `This will generate ${recipients.length} PDF(s) from personalized documents.\n\n` +
      'PDFs will be saved to your configured PDF folder.\n\n' +
      'This may take a few moments. Continue?',
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
