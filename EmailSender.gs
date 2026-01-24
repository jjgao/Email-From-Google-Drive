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

    // Search for the sent message to get its ID
    // Look in sent folder for recent message to this recipient
    let messageId = '';
    try {
      const threads = GmailApp.search(`to:${recipient.Email} subject:"${subject}" in:sent`, 0, 1);
      if (threads.length > 0) {
        const messages = threads[0].getMessages();
        if (messages.length > 0) {
          messageId = messages[messages.length - 1].getId();
        }
      }
    } catch (searchError) {
      // If search fails, continue without message ID
      console.log(`Could not retrieve message ID: ${searchError.message}`);
    }

    return {
      success: true,
      error: null,
      messageId: messageId
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

    const templateDocId = getConfig(CONFIG_KEYS.EMAIL_TEMPLATE_DOC_ID);
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
 * Send campaign to pending recipients
 * @param {number} [limit] - Optional limit on number of recipients to process
 * @returns {Object} Summary with total, success, failed counts and errors
 */
function sendCampaign(limit) {
  try {
    // Check configuration
    const validation = validateConfig();
    if (!validation.isValid) {
      throw new Error(`Configuration incomplete. Missing: ${validation.missing.join(', ')}`);
    }

    const templateDocId = getConfig(CONFIG_KEYS.EMAIL_TEMPLATE_DOC_ID);

    // Ensure status column exists
    ensureStatusColumn();

    // Get pending recipients (limited if specified)
    let recipients = getPendingRecipients();
    if (limit && limit > 0 && recipients.length > limit) {
      recipients = recipients.slice(0, limit);
    }

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
        if (result.messageId) {
          updateRecipientEmailId(recipient._rowIndex, result.messageId);
        }
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
 * Check for bounced emails by examining threads for bounce messages
 * Looks for messages from mailer-daemon or postmaster in the same thread
 * @returns {Object} Results with checked, bounced counts and bounce list
 */
function checkForBounces() {
  const recipients = getAllRecipients();

  const results = {
    checked: 0,
    bounced: 0,
    noAccess: 0,
    bounceList: [],
    errors: []
  };

  // Bounce sender patterns (case-insensitive)
  const bounceSenders = [
    'mailer-daemon',
    'postmaster',
    'mail-daemon',
    'mailerdaemon'
  ];

  for (const recipient of recipients) {
    const emailId = (recipient['Email ID'] || '').toString().trim();
    const status = (recipient['Email Status'] || '').toString().trim().toLowerCase();

    // Skip if no Email ID or already marked as bounced
    if (!emailId || status === 'bounced') {
      continue;
    }

    // Only check emails that were sent
    if (status !== 'sent') {
      continue;
    }

    results.checked++;

    try {
      // Get the original message
      const message = GmailApp.getMessageById(emailId);
      if (!message) {
        continue;
      }

      // Get the thread
      const thread = message.getThread();
      const messages = thread.getMessages();

      // Check if any message in thread is from a bounce address
      let isBounced = false;
      for (const msg of messages) {
        const from = msg.getFrom().toLowerCase();

        for (const bounceSender of bounceSenders) {
          if (from.includes(bounceSender)) {
            isBounced = true;
            break;
          }
        }

        if (isBounced) break;
      }

      if (isBounced) {
        results.bounced++;
        results.bounceList.push({
          email: recipient.Email,
          name: recipient.Name || '',
          rowIndex: recipient._rowIndex
        });

        // Update status to bounced
        updateRecipientStatus(recipient._rowIndex, 'bounced');
      }

    } catch (error) {
      // Check if it's an access/permission error
      const errorMsg = error.message.toLowerCase();
      if (errorMsg.includes('access') || errorMsg.includes('permission') ||
          errorMsg.includes('not found') || errorMsg.includes('no message')) {
        results.noAccess++;
      } else {
        results.errors.push({
          email: recipient.Email,
          error: error.message
        });
      }
    }
  }

  return results;
}

