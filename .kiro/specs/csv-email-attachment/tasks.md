# Implementation Plan

- [x] 1. Create CSV conversion utility function
  - Implement `convertRowsToCSV(rows)` function that converts 2D array to CSV string format
  - Handle proper CSV escaping for commas, quotes, and newlines in cell data
  - _Requirements: 3.3_

- [x] 2. Create CSV attachment generation function
  - Implement `generateCsvAttachments(rowsOrders, rowsReturns, baseFilename)` function
  - Generate Orders CSV attachment when orders data exists (more than just headers)
  - Generate Returns CSV attachment when returns data exists (more than just headers)
  - Create proper attachment objects with filename, content blob, and MIME type
  - Use filename format: `{baseFilename}_Orders.csv` and `{baseFilename}_Returns.csv`
  - _Requirements: 1.3, 1.4, 3.1, 3.2_

- [x] 3. Modify main export function to generate CSV attachments
  - Update `exportGstCsvForMonth(year, month)` function to call CSV attachment generation
  - Add CSV attachment generation after Google Sheet creation and before email sending
  - Implement error handling with try-catch around CSV generation to ensure graceful fallback
  - _Requirements: 1.1, 3.4_

- [x] 4. Update email sending logic to include attachments
  - Modify the email sending section in `exportGstCsvForMonth` to use `MailApp.sendEmail` with attachments parameter
  - Update email body template to mention attached CSV files with filenames and record counts
  - Include Google Sheets link as backup reference in email body
  - _Requirements: 1.2, 2.1, 2.2, 2.3_

- [x] 5. Add error handling and fallback mechanism
  - Implement fallback to original link-only email if CSV generation or attachment fails
  - Add appropriate error logging for CSV generation failures
  - Update toast notifications to indicate when CSV attachments are included vs link-only
  - _Requirements: 3.4_