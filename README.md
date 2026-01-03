# Invoice Automation (Google Sheets)

Apps Script workflow to generate invoice rows, render PDFs, and email them from Google Sheets.

## Highlights
- Generates transaction rows with formulas and lookups.
- Creates invoice PDFs from a Docs template.
- Sends invoices using Gmail draft templates.

## Tech
Apps Script, Google Sheets, Gmail, Google Drive

## Scripts
- `Initialize_sheets.js` - setup notes and formulas for sheet structure.
- `txn_generator.js` - transaction generation helpers.
- `pdf_generator.js` - PDF generation for invoices.
- `mail_merge.js` - email invoices via Gmail draft.

## Setup
1. Open the Google Sheet and go to Extensions -> Apps Script.
2. Paste the scripts into the project.
3. Review sheet/tab names referenced in the scripts and update as needed.
4. Create a Gmail draft used by `mail_merge.js`.

## Usage
- Run initialization steps to create headers and formulas.
- Generate transactions, build PDFs, then send emails.

## Notes
- Gmail quotas apply.
