/**
 * Email Sender
 * Core logic for sending personalized emails using GmailApp
 *
 * Features:
 * - Auto-attach PDF if PDF ID exists
 * - Support for multiple attachments via Attachment IDs column
 * - Template-based email personalization
 * - Test email capability
 * - Campaign sending with status tracking
 */

/**
 * Get attachments for a recipient
 * @param {Object} recipient - Recipient data object
 * @returns {Array<GoogleAppsScript.Base.Blob>} Array of file blobs to attach
 */
function getEmailAttachments(recipient) {
  const attachments = [];

  try {
    // 1. Auto-attach PDF if PDF ID exists
    const pdfId = (recipient['PDF ID'] || '').toString().trim();
    if (pdfId) {
      try {
        const pdfFile = DriveApp.getFileById(pdfId);
        attachments.push(pdfFile.getBlob());
      } catch (error) {
        console.log(`Warning: Could not attach PDF ${pdfId}: ${error.message}`);
      }
    }

    // 2. Attach files from Attachment IDs column (comma-separated)
    const attachmentIds = (recipient['Attachment IDs'] || '').toString().trim();
    if (attachmentIds) {
      const ids = attachmentIds.split(',').map(id => id.trim()).filter(id => id !== '');

      for (const fileId of ids) {
        try {
          const file = DriveApp.getFileById(fileId);
          attachments.push(file.getBlob());
        } catch (error) {
          console.log(`Warning: Could not attach file ${fileId}: ${error.message}`);
        }
      }
    }

  } catch (error) {
    console.log(`Warning: Error getting attachments for ${recipient.Email}: ${error.message}`);
  }

  return attachments;
}

/**
 * Send a single email to a recipient
 * @param {Object} recipient - Recipient data object
 * @param {string} templateDocId - Google Doc ID for template
 * @param {boolean} isTest - Whether this is a test email
 * @returns {Object} Result object with success boolean and error message
 */
function sendEmail(recipient, templateDocId, isTest = false) {
  try {
    // Validate recipient
    const validation = validateRecipient(recipient);
    if (!validation.isValid) {
      throw new Error(validation.error);
    }

    // Get configuration
    const senderName = getConfig(CONFIG_KEYS.SENDER_NAME);
    const replyTo = getConfig(CONFIG_KEYS.REPLY_TO_EMAIL);
    const subjectTemplate = getConfig(CONFIG_KEYS.EMAIL_SUBJECT);

    // Parse template and subject
    const emailBody = parseTemplate(templateDocId, recipient);
    let subject = parseSubject(subjectTemplate, recipient);

    // Add [TEST] prefix for test emails
    if (isTest) {
      subject = `[TEST] ${subject}`;
    }

    // Get attachments (PDF + any additional files)
    const attachments = getEmailAttachments(recipient);

    // Prepare email options
    const emailOptions = {
      htmlBody: emailBody,
      name: senderName,
      replyTo: replyTo
    };

    // Add attachments if any exist
    if (attachments.length > 0) {
      emailOptions.attachments = attachments;
    }

    // Send email using GmailApp
    GmailApp.sendEmail(
      recipient.Email,
      subject,
      '', // Plain text body (empty, using HTML)
      emailOptions
    );

    return {
      success: true,
      error: null
    };

  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Send test email to configured test address
 * Uses first recipient's data or sample data
 * @returns {Object} Result with success boolean, message, and error
 */
function sendTestEmail() {
  try {
    // Check configuration
    const validation = validateConfig();
    if (!validation.isValid) {
      throw new Error(`Configuration incomplete. Missing: ${validation.missing.join(', ')}`);
    }

    const templateDocId = getConfig(CONFIG_KEYS.TEMPLATE_DOC_ID);
    const testEmail = getConfig(CONFIG_KEYS.TEST_EMAIL);

    // Get first recipient data or use sample data
    let recipientData = getFirstRecipient();

    if (!recipientData) {
      // Use sample data if no recipients exist
      recipientData = {
        Name: 'Test User',
        Email: testEmail,
        Company: 'Sample Company'
      };
    } else {
      // Override email with test email
      recipientData.Email = testEmail;
    }

    // Send test email
    const result = sendEmail(recipientData, templateDocId, true);

    if (result.success) {
      // Log test email
      logEmail(recipientData.Email, recipientData.Name || '', 'test', '');

      return {
        success: true,
        message: `Test email sent successfully to ${testEmail}`,
        error: null
      };
    } else {
      throw new Error(result.error);
    }

  } catch (error) {
    return {
      success: false,
      message: '',
      error: error.message
    };
  }
}

/**
 * Send campaign to all pending recipients
 * @returns {Object} Summary with total, success, failed counts and errors
 */
function sendCampaign() {
  try {
    // Check configuration
    const validation = validateConfig();
    if (!validation.isValid) {
      throw new Error(`Configuration incomplete. Missing: ${validation.missing.join(', ')}`);
    }

    const templateDocId = getConfig(CONFIG_KEYS.TEMPLATE_DOC_ID);

    // Ensure status column exists
    ensureStatusColumn();

    // Get pending recipients
    const recipients = getPendingRecipients();

    if (recipients.length === 0) {
      throw new Error('No pending recipients found. All emails may have already been sent.');
    }

    const results = {
      total: recipients.length,
      success: 0,
      failed: 0,
      errors: []
    };

    // Send to each recipient
    for (const recipient of recipients) {
      const result = sendEmail(recipient, templateDocId, false);

      if (result.success) {
        results.success++;
        updateRecipientStatus(recipient._rowIndex, 'sent');
        logEmail(recipient.Email, recipient.Name || '', 'sent', '');
      } else {
        results.failed++;
        updateRecipientStatus(recipient._rowIndex, 'failed');
        logEmail(recipient.Email, recipient.Name || '', 'failed', result.error);
        results.errors.push({
          email: recipient.Email,
          error: result.error
        });
      }

      // Small delay to avoid rate limiting (100ms between emails)
      Utilities.sleep(100);
    }

    return results;

  } catch (error) {
    throw new Error(`Campaign failed: ${error.message}`);
  }
}

/**
 * Get sending quota information
 * @returns {Object} Quota information
 */
function getQuotaInfo() {
  const quotaRemaining = MailApp.getRemainingDailyQuota();

  return {
    remaining: quotaRemaining,
    total: quotaRemaining > 100 ? 1500 : 100 // Estimate based on remaining quota
  };
}

/**
 * Combined workflow: Generate PDFs and send campaign
 * @returns {Object} Combined results from both operations
 */
function generatePdfsAndSendCampaign() {
  try {
    // Check configuration
    const validation = validateConfig();
    if (!validation.isValid) {
      throw new Error(`Configuration incomplete. Missing: ${validation.missing.join(', ')}`);
    }

    const pdfFolderId = getConfig(CONFIG_KEYS.PDF_FOLDER_ID);
    if (!pdfFolderId) {
      throw new Error('PDF Folder ID not configured. Please set it in the Config sheet.');
    }

    // Get pending recipients for email
    const pendingRecipients = getPendingRecipients();

    if (pendingRecipients.length === 0) {
      throw new Error('No pending recipients found. All emails may have already been sent.');
    }

    // Check quota
    const quota = getQuotaInfo();
    if (quota.remaining < pendingRecipients.length) {
      throw new Error(`Insufficient quota: ${quota.remaining} remaining, ${pendingRecipients.length} emails needed.`);
    }

    const combinedResults = {
      pdfGeneration: {
        total: 0,
        success: 0,
        failed: 0,
        skipped: 0
      },
      campaign: {
        total: 0,
        success: 0,
        failed: 0,
        errors: []
      }
    };

    // Step 1: Generate PDFs for recipients with Doc IDs (but no PDF ID yet)
    try {
      const pdfResults = generateAllPdfs();
      combinedResults.pdfGeneration.total = pdfResults.total;
      combinedResults.pdfGeneration.success = pdfResults.success;
      combinedResults.pdfGeneration.failed = pdfResults.failed;
    } catch (error) {
      // If no PDFs needed, that's ok - continue with campaign
      if (!error.message.includes('No recipients with documents needing PDFs')) {
        throw error;
      }
      combinedResults.pdfGeneration.skipped = pendingRecipients.length;
    }

    // Step 2: Send campaign (PDFs will auto-attach if they exist)
    const campaignResults = sendCampaign();
    combinedResults.campaign = campaignResults;

    return combinedResults;

  } catch (error) {
    throw new Error(`Combined workflow failed: ${error.message}`);
  }
}
