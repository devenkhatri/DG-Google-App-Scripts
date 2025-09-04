# Requirements Document

## Introduction

The GST Invoices export feature currently sends an email with a Google Sheets link to the generated report. The customer wants to receive the actual CSV file as an email attachment instead of just a link to the Google Sheet. This enhancement will modify the existing email functionality to export the Google Sheet data as CSV format and attach it directly to the email.

## Requirements

### Requirement 1

**User Story:** As a business user receiving GST invoice reports, I want to receive the CSV file as an email attachment, so that I can directly access the data without needing to click links or access Google Drive.

#### Acceptance Criteria

1. WHEN the GST export process completes THEN the system SHALL generate a CSV file from the Google Sheet data
2. WHEN sending the notification email THEN the system SHALL attach the CSV file to the email
3. WHEN the CSV file is attached THEN the system SHALL use the same filename format as the Google Sheet: `GST_Invoices_{YYYYMM}_{MONTH}-{YYYYMMDD}.csv`
4. WHEN the email is sent THEN the system SHALL include both Orders and Returns data in separate CSV attachments if both contain data

### Requirement 2

**User Story:** As a business user, I want the email to still contain summary information about the export, so that I can quickly understand what data is included without opening the attachments.

#### Acceptance Criteria

1. WHEN the email is sent with CSV attachments THEN the system SHALL include summary information (month, order count, returns count) in the email body
2. WHEN CSV files are attached THEN the email body SHALL mention the attachment names
3. WHEN the export completes THEN the system SHALL still create the Google Sheet in Drive for backup/reference purposes

### Requirement 3

**User Story:** As a system administrator, I want the CSV export to handle both normal orders and returns properly, so that the data integrity is maintained across different file formats.

#### Acceptance Criteria

1. WHEN there are both orders and returns THEN the system SHALL create separate CSV attachments for "Orders" and "Returns"
2. WHEN there are only orders (no returns) THEN the system SHALL attach only the Orders CSV file
3. WHEN converting Google Sheet to CSV THEN the system SHALL preserve all column formatting and data accuracy
4. IF the CSV generation fails THEN the system SHALL fallback to sending the Google Sheet link as before