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
    // Setup & Configuration
    .addItem('⚙️ Open Config Sheet', 'openConfigSheetUI')
    .addItem('📋 Create Sample Data', 'createSampleDataUI')
    .addItem('📧 Create Sample Email Template', 'createSampleEmailTemplateUI')
    .addItem('📄 Create Sample PDF Template', 'createSamplePdfTemplateUI')
    .addSeparator()
    // Document Generation
    .addItem('📑 Create All Documents', 'createAllDocumentsUI')
    .addItem('🔄 Regenerate All Documents', 'regenerateAllDocumentsUI')
    .addItem('📕 Generate All PDFs', 'generateAllPdfsUI')
    .addItem('🔄 Regenerate All PDFs', 'regenerateAllPdfsUI')
    .addSeparator()
    // Email Campaign
    .addItem('📧 Send Test Email', 'sendTestEmailUI')
    .addItem('🚀 Send Campaign', 'sendCampaignUI')
    .addItem('⚡ Generate PDFs & Send', 'generatePdfsAndSendCampaignUI')
    .addSeparator()
    // Reports & Analytics
    .addItem('📊 View Logs', 'openLogSheet')
    .addItem('📈 View Statistics', 'showStatsDialog')
    .addSeparator()
    .addSubMenu(ui.createMenu('Advanced')
      .addItem('Setup via Dialog (Legacy)', 'showConfigDialog')
      .addItem('Clear Configuration', 'clearConfigUI')
      .addItem('Ensure Status Column', 'ensureStatusColumnUI')
      .addItem('🗑️ Delete Orphan Files', 'deleteOrphanFilesUI'))
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

    <div class="field">
      <label>Recipient Sheet Name:</label>
      <input type="text" id="recipientSheet" value="${currentConfig.RECIPIENT_SHEET_NAME || 'Recipients'}" />
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
 * Generate PDFs and send campaign in one workflow (UI wrapper)
 */
function generatePdfsAndSendCampaignUI() {
  const ui = SpreadsheetApp.getUi();

  // Check configuration first
  const validation = validateConfig();
  if (!validation.isValid) {
    ui.alert('Configuration Required', `Please configure the following:\n\n${validation.missing.join('\n')}`, ui.ButtonSet.OK);
    return;
  }

  // Check PDF folder
  const pdfFolderId = getConfig(CONFIG_KEYS.PDF_FOLDER_ID);
  if (!pdfFolderId) {
    ui.alert(
      'PDF Folder Not Configured',
      'This workflow requires PDF Folder ID to be configured.\n\nPlease set it in the Config sheet.',
      ui.ButtonSet.OK
    );
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
      'Generate PDFs & Send Campaign',
      `This will:\n` +
      `1. Generate PDFs for ${summary.pending} recipient(s) with documents\n` +
      `2. Send campaign emails with PDFs attached\n\n` +
      `Quota remaining: ${quota.remaining}\n\n` +
      `This action cannot be undone. Continue?`,
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      // Show progress message
      SpreadsheetApp.getActiveSpreadsheet().toast('Generating PDFs and sending emails...', 'Workflow Progress', -1);

      const results = generatePdfsAndSendCampaign();

      // Log summary
      logCampaignSummary(results.campaign);

      // Hide progress
      SpreadsheetApp.getActiveSpreadsheet().toast('Complete!', 'Workflow Progress', 3);

      // Show results
      let message = `Workflow Complete!\n\n`;
      message += `PDF GENERATION:\n`;
      message += `Generated: ${results.pdfGeneration.success}\n`;
      message += `Failed: ${results.pdfGeneration.failed}\n`;
      message += `Skipped: ${results.pdfGeneration.skipped}\n\n`;

      message += `EMAIL CAMPAIGN:\n`;
      message += `Sent: ${results.campaign.success}\n`;
      message += `Failed: ${results.campaign.failed}\n`;
      message += `Total: ${results.campaign.total}\n`;

      if (results.campaign.failed > 0) {
        message += `\nFirst few errors:\n`;
        const errorPreview = results.campaign.errors.slice(0, 3);
        errorPreview.forEach(err => {
          message += `- ${err.email}: ${err.error}\n`;
        });

        if (results.campaign.errors.length > 3) {
          message += `\n(and ${results.campaign.errors.length - 3} more - check logs)`;
        }
      }

      ui.alert('Workflow Results', message, ui.ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Error occurred', 'Workflow Progress', 3);
    ui.alert('Error', `Workflow failed:\n\n${error.message}`, ui.ButtonSet.OK);
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
    const headers = ['Email', 'Name', 'Company', 'Street', 'City', 'State', 'ZIP', 'Status', 'Doc ID', 'PDF ID', 'Attachment IDs'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');

    // Add sample data
    // Note: Third recipient has missing City field to test validation
    const sampleData = [
      ['your-email@example.com', 'John Doe', 'Acme Inc', '123 Main Street', 'New York', 'NY', '10001', 'pending', '', '', ''],
      ['another-email@example.com', 'Jane Smith', 'Tech Corp', '456 Oak Avenue', 'San Francisco', 'CA', '94102', 'pending', '', '', ''],
      ['third-email@example.com', 'Bob Johnson', 'StartupXYZ', '789 Pine Road', '', 'TX', '73301', 'pending', '', '', '']
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
    sheet.setColumnWidth(11, 350); // Attachment IDs

    // Freeze header row
    sheet.setFrozenRows(1);

    // Activate the new sheet
    spreadsheet.setActiveSheet(sheet);

    ui.alert(
      'Success',
      `Sample recipient sheet "${recipientSheetName}" created with 3 test recipients.\n\n` +
      'IMPORTANT: Please replace the sample email addresses with valid test emails before sending!\n\n' +
      'NOTE: The third recipient (Bob Johnson) has a missing City field to test validation. Document generation will skip this recipient if City is used in your template.\n\n' +
      'TIP: Use the "Attachment IDs" column to attach additional files to emails (comma-separated Drive file IDs). PDFs are auto-attached when PDF ID exists.\n\n' +
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
    const namePara = body.appendParagraph('{{Name}}');
    namePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    namePara.setFontSize(16);
    namePara.setBold(true);
    namePara.setForegroundColor('#4285f4');

    body.appendParagraph('');

    // Company
    const companyPara = body.appendParagraph('{{Company}}');
    companyPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    companyPara.setFontSize(12);

    body.appendParagraph('');

    // Address block
    body.appendParagraph('{{Street}}');
    body.appendParagraph('{{City}}, {{State}} {{ZIP}}');

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
      '• Use {{FieldName}} for placeholders - field names must match your spreadsheet columns exactly',
      '• Available fields from sample data: {{Name}}, {{Company}}, {{Street}}, {{City}}, {{State}}, {{ZIP}}, {{Email}}',
      '• This template creates a formal certificate/letter format suitable for PDFs',
      '• You can add any custom fields by adding columns to your Recipients sheet',
      '• Customize the header, body content, and layout to match your needs',
      '• This is separate from your email template - email template is for email body, PDF template is for attached documents',
      '• Delete these instructions before using the template',
      '',
      'USAGE:',
      '1. Edit this template with your company information and desired content',
      '2. Copy the Document ID from the URL',
      '3. Add it to the "PDF Template Document ID" field in your Config sheet',
      '4. Use "Create All Documents" to generate personalized docs',
      '5. Use "Generate All PDFs" to convert docs to PDFs',
      '6. PDFs will automatically attach to emails when you send campaigns'
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

    // Get recipients with documents
    const allRecipients = getAllRecipientsFormatted();
    const recipientsWithDocs = allRecipients.filter(r => {
      const docId = (r.data['Doc ID'] || '').toString().trim();
      return docId !== '';
    });

    if (recipientsWithDocs.length === 0) {
      ui.alert(
        'No Documents Found',
        'No recipients have documents yet. Generate documents first.',
        ui.ButtonSet.OK
      );
      return;
    }

    const response = ui.alert(
      'Regenerate All PDFs',
      `⚠️ WARNING: This will regenerate PDFs for ALL ${recipientsWithDocs.length} recipient(s) with documents, overwriting existing PDFs.\n\n` +
      'Existing PDFs will be replaced with new ones.\n\n' +
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
 * Delete orphan files (UI wrapper)
 */
function deleteOrphanFilesUI() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Delete Orphan Files',
    '⚠️ WARNING: This will permanently delete files in your output folders that don\'t match any recipient in the sheet.\n\n' +
    'Orphan files are those with IDs that don\'t appear in the "Doc ID" or "PDF ID" columns.\n\n' +
    'Deleted files will be moved to trash and can be recovered from Google Drive trash for 30 days.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Scanning for orphan files...', 'Processing', -1);

    const results = deleteAllOrphanFiles();

    SpreadsheetApp.getActiveSpreadsheet().toast('Cleanup complete!', 'Success', 3);

    // Show results
    let message = `Orphan files have been moved to trash:\n\n`;
    message += `Documents deleted: ${results.documents.deleted}\n`;
    message += `PDFs deleted: ${results.pdfs.deleted}\n\n`;

    if (results.documents.deleted > 0) {
      message += 'Deleted documents:\n';
      results.documents.files.forEach(file => {
        message += `• ${file.name}\n`;
      });
      message += '\n';
    }

    if (results.pdfs.deleted > 0) {
      message += 'Deleted PDFs:\n';
      results.pdfs.files.forEach(file => {
        message += `• ${file.name}\n`;
      });
    }

    if (results.documents.deleted === 0 && results.pdfs.deleted === 0) {
      message = 'No orphan files found. All files match current recipients.';
    } else {
      message += '\nFiles can be recovered from Google Drive trash for 30 days.';
    }

    ui.alert('Cleanup Complete', message, ui.ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Failed to delete orphan files', 'Error', 3);
    ui.alert('Error', `Failed to delete orphan files:\n\n${error.message}`, ui.ButtonSet.OK);
  }
}
