# Email-From-Google-Drive

Send personalized emails using templates from Google Docs and recipient data from Google Sheets.

## Overview

This Google Apps Script application allows you to send personalized email campaigns directly from Google Sheets. It uses Google Docs for email templates with placeholder support and provides a simple menu-based interface for non-technical users.

## Features (MVP1)

- ✅ **Template-based emails** - Use Google Docs with `{{placeholder}}` syntax
- ✅ **Personalized content** - Replace placeholders with recipient data from sheets
- ✅ **Basic HTML support** - Bold, italic, links, and headers in templates
- ✅ **Test email capability** - Send test emails before launching campaigns
- ✅ **Status tracking** - Track sent/failed status for each recipient
- ✅ **Comprehensive logging** - All email activities logged to a sheet
- ✅ **Simple UI** - Custom menu in Google Sheets for easy access
- ✅ **Quota management** - Check available quota before sending

## Installation

### 1. Set Up Google Sheet

1. Create a new Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete any existing code in Code.gs
4. Create the following files in the Apps Script editor:
   - `appsscript.json`
   - `Code.gs`
   - `Config.gs`
   - `TemplateParser.gs`
   - `RecipientManager.gs`
   - `EmailSender.gs`
   - `Logger.gs`
   - `UI.gs`

5. Copy the code from this repository into each corresponding file
6. Save the project (Ctrl+S or Cmd+S)
7. Refresh your Google Sheet

### 2. Create Recipient Sheet

Create a sheet named "Recipients" with the following columns:

| Email | Name | Company | Status |
|-------|------|---------|--------|
| john@example.com | John Doe | Acme Inc | pending |

**Or** run the `createSampleRecipientSheet()` function from the Apps Script editor to create a sample sheet automatically.

### 3. Create Email Template

1. Create a new Google Doc
2. Write your email template using HTML formatting
3. Use `{{FieldName}}` for placeholders (must match column names in your sheet)

Example template:
```
Hi {{Name}},

Thank you for being a valued customer of {{Company}}.

We appreciate your business!

Best regards,
The Team
```

4. Copy the Document ID from the URL:
   `https://docs.google.com/document/d/DOCUMENT_ID_HERE/edit`

## Configuration

1. Open your Google Sheet
2. Click **Email Campaign → Setup Configuration**
3. Fill in the required fields:
   - **Template Document ID** - ID from your Google Doc URL
   - **Sender Name** - Name that appears as sender
   - **Reply-To Email** - Email for replies
   - **Test Email** - Your email for testing
   - **Email Subject** - Subject line (can use {{placeholders}})
   - **Recipient Sheet Name** - Name of sheet with recipients (default: "Recipients")

4. Click **Save Configuration**

## Usage

### Send Test Email

1. Click **Email Campaign → Send Test Email**
2. Check your test email inbox
3. Verify the content looks correct

### Send Campaign

1. Ensure all recipients have "pending" or empty status
2. Click **Email Campaign → Send Campaign**
3. Review the confirmation dialog
4. Click **Yes** to send
5. Wait for completion (progress shown in toast)
6. Review the results dialog

### View Logs

1. Click **Email Campaign → View Logs**
2. Review all email activities with timestamps and status
3. Check error messages for failed emails

### View Statistics

1. Click **Email Campaign → View Statistics**
2. See recipient counts, log statistics, and quota information

## File Structure

```
/
├── Code.gs              # Main entry point, system tests, sample data
├── Config.gs            # Configuration management using Script Properties
├── TemplateParser.gs    # Template parsing and placeholder replacement
├── RecipientManager.gs  # Recipient list management and validation
├── EmailSender.gs       # Email sending logic using GmailApp
├── Logger.gs            # Activity logging to Google Sheets
├── UI.gs                # Custom menu and dialogs
└── appsscript.json      # Apps Script manifest with OAuth scopes
```

## How It Works

1. **Template Parsing**: Reads your Google Doc template and extracts content with basic HTML formatting
2. **Recipient Loading**: Reads recipients from your Google Sheet with all their data
3. **Personalization**: Replaces `{{placeholders}}` with actual recipient data
4. **Sending**: Uses Gmail API to send personalized emails to each recipient
5. **Tracking**: Updates status in sheet and logs all activities

## Troubleshooting

### "Authorization Required" Error
1. Click **Email Campaign** menu
2. Google will prompt for authorization
3. Click **Advanced → Go to [Project Name]**
4. Grant necessary permissions

### "Template document not found"
- Check that the Document ID is correct
- Ensure you have access to the Google Doc
- Make sure the Doc isn't in Trash

### "No recipients found"
- Ensure "Recipients" sheet exists
- Check that Email column has valid addresses
- Verify sheet name matches configuration

### "Quota exceeded"
- Gmail has daily sending limits:
  - Free accounts: 100 emails/day
  - Google Workspace: 1,500 emails/day
- Wait 24 hours for quota to reset
- Use **Email Campaign → View Statistics** to check remaining quota

## Testing

Run the system test from Apps Script editor:
1. Open Apps Script editor
2. Select `runSystemTest` function
3. Click **Run**
4. Review the test report

## Limitations (MVP1)

- ❌ No batch processing with delays (sends all at once)
- ❌ No automatic retry for failed emails
- ❌ No personalized attachments
- ❌ No scheduled sending
- ❌ Basic HTML only (complex layouts not supported)

These features will be added in future versions (MVP2, MVP3).

## Support

For issues or questions, please create an issue on GitHub.

## License

MIT License - see LICENSE file for details
