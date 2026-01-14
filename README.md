# Email-From-Google-Drive

Send personalized emails using templates from Google Docs and recipient data from Google Sheets.

## Overview

This Google Apps Script application allows you to send personalized email campaigns directly from Google Sheets. It uses Google Docs for email templates with placeholder support and provides a simple menu-based interface for non-technical users.

**Current Version: 0.3.0** - Multi-Sheet Support & Improved Folder Organization

## Features (MVP1.1)

### Email Features
- ✅ **Template-based emails** - Use Google Docs with `{{placeholder}}` syntax
- ✅ **Personalized content** - Replace placeholders with recipient data from sheets
- ✅ **Basic HTML support** - Bold, italic, links, and headers in templates
- ✅ **Auto-attach PDFs** - Automatically attach PDFs when PDF ID exists
- ✅ **Multiple attachments** - Support for multiple file attachments per email
- ✅ **Test email capability** - Send test emails before launching campaigns
- ✅ **Status tracking** - Track sent/failed status for each recipient
- ✅ **Quota management** - Check available quota before sending

### Document Generation
- ✅ **Personalized documents** - Generate Google Docs for each recipient
- ✅ **Smart validation** - Only validates fields actually used in your template
- ✅ **PDF export** - Convert generated documents to PDFs
- ✅ **Incremental & full regeneration** - Create new docs or regenerate all existing ones
- ✅ **Orphan file cleanup** - Remove files not matching any recipient
- ✅ **Multi-sheet support** - Work with multiple recipient sheets in the same spreadsheet
- ✅ **Auto-organized folders** - Files automatically organized in `OutputFolder/SheetName/docs/` and `OutputFolder/SheetName/pdfs/`

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

Create a sheet with recipient data. Any sheet name works (except "Config" and "Email Logs" which are system sheets). The system will use the currently active sheet as the recipient sheet.

Minimum required columns:

| Email | First Name | Last Name |
|-------|------------|-----------|
| john@example.com | John | Doe |

Additional columns are auto-created when needed:
- **Filename** - Custom document name (optional, auto-generated if empty)
- **Doc ID** - Tracks generated Google Docs
- **PDF ID** - Tracks generated PDFs
- **Email Status** - Tracks email send status (pending/sent/failed)

**Or** use **Email Campaign → Setup → Create Sample Data** to create a sample sheet automatically.

### 3. Create Templates

#### Email Template (Required)

1. Create a new Google Doc for your email body
2. Write your email content using basic HTML formatting
3. Use `{{FieldName}}` for placeholders (must match column names in your sheet)

Example email template:
```
Hi {{Name}},

Thank you for being a valued customer of {{Company}}.

We appreciate your business!

Best regards,
The Team
```

4. Copy the Document ID from the URL:
   `https://docs.google.com/document/d/EMAIL_TEMPLATE_ID/edit`

#### PDF Template (Optional)

If you want to generate personalized PDF documents to attach to emails:

1. Create a separate Google Doc for your PDF content
2. Format it as you want the PDF to appear (letterhead, contract, invoice, etc.)
3. Use `{{FieldName}}` for placeholders

Example PDF template:
```
Dear {{Name}},

This letter confirms your registration with {{Company}}.

Address: {{Street}}, {{City}}, {{State}} {{ZIP}}

Thank you for choosing us!

Sincerely,
Management
```

4. Copy the Document ID from the URL:
   `https://docs.google.com/document/d/PDF_TEMPLATE_ID/edit`

**Note:** The email template and PDF template are completely separate. The email template is used for the email body, while the PDF template is used to generate personalized PDFs that get attached to the emails.

## Configuration

1. Open your Google Sheet
2. Click **Email Campaign → Setup → Setup via Config Sheet**
3. Fill in the required fields in the Config sheet:
   - **Email Template Document ID** - ID for email body template (required for emails)
   - **PDF Template Document ID** - ID for PDF generation template (required for document generation)
   - **Output Folder ID** - Google Drive folder ID (required for document generation)
   - **Sender Name** - Name that appears as sender (required for emails)
   - **Reply-To Email** - Email for replies (required for emails)
   - **Test Email** - Your email for testing (required)
   - **Email Subject** - Subject line (can use {{placeholders}})
   - **Document Name Template** - Custom filename format (optional)

**Note:** The system automatically creates subfolders within your Output Folder:
- `OutputFolder/SheetName/docs/` - For generated Google Docs
- `OutputFolder/SheetName/pdfs/` - For generated PDFs

This keeps files organized when working with multiple recipient sheets.

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

### Document Generation (Optional - uses PDF Template)
1. **Template Analysis**: Scans your PDF template Google Doc for `{{FieldName}}` placeholders
2. **Smart Validation**: Only validates fields that appear in your PDF template (optional fields can be empty)
3. **Document Creation**: Creates a personalized copy of the PDF template for each recipient
4. **Placeholder Replacement**: Replaces all `{{placeholders}}` with actual recipient data
5. **PDF Export**: Converts generated documents to PDF format
6. **Tracking**: Updates Doc ID and PDF ID columns, logs all activities
7. **Auto-Attachment**: PDFs are automatically attached to emails when PDF ID exists

### Email Campaign (uses Email Template)
1. **Template Parsing**: Reads your email template Google Doc and extracts content with HTML formatting
2. **Recipient Loading**: Reads recipients from your Google Sheet with all their data
3. **Personalization**: Replaces `{{placeholders}}` in the email body with actual recipient data
4. **Attachment Handling**: Attaches PDFs (if generated) and any additional files specified
5. **Sending**: Uses Gmail API to send personalized emails to each recipient
6. **Status Tracking**: Updates status column and logs all email activities

**Key Point**: Email template is for the email body content, PDF template is for generating attached documents. They are completely independent.

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
- Ensure you have a data sheet with recipients (not Config or Email Logs)
- Make sure the active sheet contains recipient data
- Check that Email column has valid addresses

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
- **No pending recipients**: All recipients already have Doc IDs - use "Regenerate All" or clear Doc IDs
- **Permission denied**: Ensure you have write access to the output folder
- **Wrong sheet selected**: Make sure the active sheet contains your recipient data (confirmation dialogs show which sheet will be used)

## Testing

Run the system test from Apps Script editor:
1. Open Apps Script editor
2. Select `runSystemTest` function
3. Click **Run**
4. Review the test report

## Limitations (MVP1.1)

- ❌ No batch processing with delays (sends all at once)
- ❌ No automatic retry for failed emails
- ❌ No personalized attachments
- ❌ No scheduled sending
- ❌ Basic HTML only (complex layouts not supported)

These features will be added in future versions (MVP2, MVP3).

## Version History

### v0.3.0 (2026-01-14) - Multi-Sheet Support & Improved Folder Organization
**New Features:**
- **Multi-sheet support**: Work with multiple recipient sheets in the same spreadsheet
- **Auto-organized folders**: Files automatically organized in `OutputFolder/SheetName/docs/` and `OutputFolder/SheetName/pdfs/`
- **Sheet selection display**: All confirmation dialogs now show which sheet will be used
- **Ensure Required Columns**: New menu item to auto-create Filename, Doc ID, PDF ID, and Email Status columns

**Changes:**
- Renamed "Status" column to "Email Status" for clarity
- Removed "Recipient Sheet Name" config setting (now uses active sheet)
- Removed "PDF Folder ID" config setting (now auto-created as subfolder)
- Removed combined "Generate PDFs & Send" menu item (use separate operations for more control)
- Reorganized menu structure with Setup, Testing, and Maintenance submenus

**Improvements:**
- Better menu organization with logical groupings
- Confirmation dialogs show recipient count and sheet name
- Simplified configuration with fewer required settings
- Better folder organization for multi-campaign workflows

### v0.2.0 (2026-01-12) - Enhanced Filename Generation & Bug Fixes
**New Features:**
- **Fill Default Filenames**: New menu item to auto-generate filenames for all recipients
- **Improved filename generation**: Better handling of missing First Name or Last Name
- **Template name in filenames**: Default format now includes template name (e.g., "Template Name - First Last")

**Bug Fixes:**
- Fixed orphan file detection when using same folder for both documents and PDFs
- Fixed undefined `logError` and `logInfo` function calls
- Fixed PDF naming to match source document names
- Fixed default filename format to include template name

**Improvements:**
- Removed Filename column from sample data (now optional)
- Changed default `DOCUMENT_NAME_TEMPLATE` to empty (uses smart defaults)
- Enhanced filename fallback logic for missing name fields
- Added same-folder mode detection with improved diagnostics

### MVP1.1 - Separate Email and PDF Templates
**Key Changes:**
- **Split template configuration**: Separated single template into EMAIL_TEMPLATE_DOC_ID (required) and PDF_TEMPLATE_DOC_ID (optional)
- **Independent templates**: Email template is used only for email body content, PDF template is used only for document/PDF generation
- **Improved sample templates**: Added separate sample template creators for email and PDF templates
- **Better recipient filtering**: Separated recipient selection logic for document creation vs email sending
- **Enhanced validation**: PDF template validation is now optional - only required when generating documents

**Migration from MVP1:**
- Old `TEMPLATE_DOC_ID` needs to be split into two separate IDs
- Email Template Document ID is required for sending emails
- PDF Template Document ID is optional and only needed if generating personalized PDFs
- Update configuration via **Email Campaign → Setup Configuration**

### MVP1.0 - Core Email Sending
- Initial release with basic email campaign functionality
- Template-based personalization with {{placeholder}} syntax
- Document and PDF generation capabilities
- Status tracking and logging

## Support

For issues or questions, please create an issue on GitHub.

## License

MIT License - see LICENSE file for details
