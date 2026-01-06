# Email-From-Google-Drive

Send personalized emails using templates from Google Docs and recipient data from Google Sheets.

## Overview

This Google Apps Script application allows you to send personalized email campaigns directly from Google Sheets. It uses Google Docs for email templates with placeholder support and provides a simple menu-based interface for non-technical users.

## Features (MVP1)

### Email Features
- ✅ **Template-based emails** - Use Google Docs with `{{placeholder}}` syntax
- ✅ **Personalized content** - Replace placeholders with recipient data from sheets
- ✅ **Basic HTML support** - Bold, italic, links, and headers in templates
- ✅ **Auto-attach PDFs** - Automatically attach PDFs when PDF ID exists
- ✅ **Multiple attachments** - Support for multiple file attachments per email
- ✅ **One-click workflow** - Generate PDFs and send emails in single operation
- ✅ **Test email capability** - Send test emails before launching campaigns
- ✅ **Status tracking** - Track sent/failed status for each recipient
- ✅ **Quota management** - Check available quota before sending

### Document Generation
- ✅ **Personalized documents** - Generate Google Docs for each recipient
- ✅ **Smart validation** - Only validates fields actually used in your template
- ✅ **PDF export** - Convert generated documents to PDFs
- ✅ **Incremental & full regeneration** - Create new docs or regenerate all existing ones
- ✅ **Orphan file cleanup** - Remove files not matching any recipient

### System
- ✅ **Comprehensive logging** - All activities logged to a sheet
- ✅ **Simple UI** - Custom menu in Google Sheets for easy access
- ✅ **Config sheet** - Easy configuration management

## Installation

### 1. Set Up Google Sheet

1. Create a new Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete any existing code in Code.gs
4. Create the following files in the Apps Script editor:
   - `appsscript.json`
   - `Code.gs`
   - `Config.gs`
   - `DocumentGenerator.gs`
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

**Note:** PDFs are automatically attached to emails when a recipient has a PDF ID. Additional attachments can be specified in the "Attachment IDs" column.

### One-Click Workflow: Generate PDFs & Send

The fastest way to send emails with PDF attachments:

1. Click **Email Campaign → ⚡ Generate PDFs & Send**
2. System will:
   - Generate PDFs for all recipients with documents (if not already generated)
   - Send campaign emails with PDFs automatically attached
3. Review confirmation dialog showing pending count
4. Click **Yes** to execute the workflow
5. View combined results for both PDF generation and email sending

**Benefits:**
- Single operation instead of two separate steps
- Ensures PDFs are fresh and up-to-date
- Perfect for recurring campaigns
- Saves time and reduces errors

### Adding File Attachments

To attach files to emails for specific recipients:

1. Get the Google Drive file ID(s) you want to attach
   - Open the file in Google Drive
   - Copy the ID from the URL: `drive.google.com/file/d/[FILE_ID]/view`
2. In your Recipients sheet, find the "Attachment IDs" column
3. Paste the file ID(s):
   - Single file: `1abc123xyz`
   - Multiple files: `1abc123xyz,2def456uvw,3ghi789rst` (comma-separated)
4. When the email is sent, all specified files will be attached

**Automatic PDF Attachment:**
- If a recipient has a PDF ID, the PDF is automatically attached
- No need to add the PDF ID to "Attachment IDs" column
- "Attachment IDs" is for additional files beyond the main PDF

### View Logs

1. Click **Email Campaign → View Logs**
2. Review all email activities with timestamps and status
3. Check error messages for failed emails

### View Statistics

1. Click **Email Campaign → View Statistics**
2. See recipient counts, log statistics, and quota information

### Generate Personalized Documents

1. Click **Email Campaign → Create All Documents**
2. System will create a personalized Google Doc for each recipient without a Doc ID
3. Only recipients with all required template fields filled will be processed
4. Documents are saved to your configured Output Folder
5. Doc IDs are automatically tracked in the Recipients sheet

**Note:** The system automatically detects which fields are required by scanning your template document. Only fields that appear as `{{FieldName}}` in the template need to be filled.

### Regenerate All Documents

1. Click **Email Campaign → Regenerate All Documents**
2. Recreates documents for ALL recipients, overwriting existing ones
3. Useful after updating your template
4. PDF IDs are automatically cleared to maintain synchronization

### Generate PDFs

1. Click **Email Campaign → Generate All PDFs**
2. Converts documents to PDF format for recipients with Doc IDs
3. PDFs are saved to your configured PDF Folder
4. PDF IDs are tracked in the Recipients sheet

### Clean Up Orphan Files

1. Click **Email Campaign → Advanced → Delete Orphan Files**
2. Removes files from Drive folders that don't match any recipient
3. Files are moved to trash (recoverable for 30 days)
4. Useful for cleaning up after data changes

## File Structure

```
/
├── Code.gs              # Main entry point, system tests, sample data
├── Config.gs            # Configuration management (Config sheet)
├── DocumentGenerator.gs # Document & PDF generation with validation
├── TemplateParser.gs    # Template parsing and placeholder replacement
├── RecipientManager.gs  # Recipient list management and validation
├── EmailSender.gs       # Email sending logic using GmailApp
├── Logger.gs            # Activity logging to Google Sheets
├── UI.gs                # Custom menu and dialogs
└── appsscript.json      # Apps Script manifest with OAuth scopes
```

## How It Works

### Document Generation
1. **Template Analysis**: Scans your Google Doc template for `{{FieldName}}` placeholders
2. **Smart Validation**: Only validates fields that appear in your template (optional fields can be empty)
3. **Document Creation**: Creates a personalized copy of the template for each recipient
4. **Placeholder Replacement**: Replaces all `{{placeholders}}` with actual recipient data
5. **PDF Export**: Optionally converts documents to PDF format
6. **Tracking**: Updates Doc ID and PDF ID columns, logs all activities

### Email Campaign
1. **Template Parsing**: Reads your Google Doc template and extracts content with HTML formatting
2. **Recipient Loading**: Reads recipients from your Google Sheet with all their data
3. **Personalization**: Replaces `{{placeholders}}` with actual recipient data
4. **Sending**: Uses Gmail API to send personalized emails to each recipient
5. **Status Tracking**: Updates status column and logs all email activities

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

### "Missing required fields" Error
- The system scans your template for `{{FieldName}}` placeholders
- Only fields that appear in the template are required
- Check which fields are mentioned in your template document
- Ensure those exact fields are filled in your Recipients sheet
- Field names are case-sensitive (e.g., `{{Name}}` ≠ `{{name}}`)

### Document Generation Issues
- **Output Folder not configured**: Set Output Folder ID in Config sheet
- **PDF Folder not configured**: Set PDF Folder ID in Config sheet (optional for PDFs)
- **No pending recipients**: All recipients already have Doc IDs - use "Regenerate All" or clear Doc IDs
- **Permission denied**: Ensure you have write access to the output folders

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
