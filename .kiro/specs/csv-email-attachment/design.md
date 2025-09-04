# Design Document

## Overview

This design modifies the existing GST Invoices export functionality to send CSV file attachments via email instead of just Google Sheets links. The solution maintains backward compatibility by still creating Google Sheets in Drive while adding CSV export and email attachment capabilities.

## Architecture

The enhancement follows the existing architecture pattern with minimal changes to the core data processing flow. The main modification occurs in the email notification section where we add CSV generation and attachment functionality.

### Current Flow
1. Fetch Shopify orders → Filter → Split into Orders/Returns → Build sheet rows → Save Google Sheet → Send email with link

### Enhanced Flow  
1. Fetch Shopify orders → Filter → Split into Orders/Returns → Build sheet rows → Save Google Sheet → **Generate CSV files → Send email with CSV attachments**

## Components and Interfaces

### New Functions

#### `convertRowsToCSV(rows)`
- **Purpose**: Convert 2D array of sheet rows to CSV format string
- **Input**: `rows` - Array of arrays representing sheet data
- **Output**: String in CSV format with proper escaping
- **Logic**: Handle comma escaping, quote wrapping, and newline formatting

#### `generateCsvAttachments(rowsOrders, rowsReturns, baseFilename)`
- **Purpose**: Create CSV blob attachments for email
- **Input**: 
  - `rowsOrders` - Orders data rows
  - `rowsReturns` - Returns data rows  
  - `baseFilename` - Base filename without extension
- **Output**: Array of attachment objects for MailApp.sendEmail
- **Logic**: 
  - Generate Orders CSV if orders exist
  - Generate Returns CSV if returns exist (more than just headers)
  - Create proper filenames: `{baseFilename}_Orders.csv`, `{baseFilename}_Returns.csv`

### Modified Functions

#### `exportGstCsvForMonth(year, month)`
- **Changes**: Replace simple email sending with enhanced email that includes CSV attachments
- **New Logic**: 
  - Generate CSV attachments after Google Sheet creation
  - Update email body to mention attachments
  - Use `MailApp.sendEmail` with attachments parameter

## Data Models

### CSV Attachment Object
```javascript
{
  fileName: string,     // e.g., "GST_Invoices_202501_JANUARY-202501091234_Orders.csv"
  content: Blob,        // CSV content as blob
  mimeType: string      // "text/csv"
}
```

### Email Structure
```javascript
{
  to: string,           // recipient email
  subject: string,      // existing subject format
  body: string,         // enhanced body mentioning attachments
  attachments: Array    // array of attachment objects
}
```

## Error Handling

### CSV Generation Failures
- **Strategy**: Graceful degradation to existing Google Sheets link behavior
- **Implementation**: Wrap CSV generation in try-catch, log errors, continue with link-only email
- **User Experience**: Toast notification indicates partial success

### Email Attachment Size Limits
- **Constraint**: Google Apps Script email attachments limited to 25MB total
- **Mitigation**: CSV files are typically small (< 1MB for thousands of orders)
- **Fallback**: If attachment fails due to size, send link-only email with error message

### Missing Data Scenarios
- **Empty Orders**: Send email with Returns CSV only (if returns exist)
- **Empty Returns**: Send email with Orders CSV only
- **No Data**: Send email with summary but no attachments

## Testing Strategy

### Unit Testing Approach
1. **CSV Conversion Testing**
   - Test `convertRowsToCSV` with various data scenarios
   - Verify proper comma/quote escaping
   - Test empty data handling

2. **Attachment Generation Testing**
   - Test `generateCsvAttachments` with different data combinations
   - Verify filename generation
   - Test blob creation and MIME types

3. **Integration Testing**
   - Test full export flow with CSV attachments
   - Verify email delivery with attachments
   - Test fallback to link-only behavior

### Test Data Scenarios
- Normal orders only
- Returns only  
- Both orders and returns
- Empty datasets
- Large datasets (performance testing)
- Special characters in data (CSV escaping)

## Implementation Details

### CSV Format Specifications
- **Delimiter**: Comma (`,`)
- **Text Qualifier**: Double quotes (`"`) when needed
- **Escape Method**: Double quotes escaped as `""`
- **Line Endings**: `\n` (Unix style)
- **Encoding**: UTF-8

### Filename Convention
- **Orders**: `{baseFilename}_Orders.csv`
- **Returns**: `{baseFilename}_Returns.csv`
- **Base Format**: `GST_Invoices_{YYYYMM}_{MONTH}-{YYYYMMDD}`

### Email Body Enhancement
```
Hello Team,

The GST Invoices export files have been generated successfully and are attached to this email.

Month: JANUARY 2025
Orders Exported: 150
Returns Exported: 5

Attached Files:
- GST_Invoices_202501_JANUARY-202501091234_Orders.csv (150 orders)
- GST_Invoices_202501_JANUARY-202501091234_Returns.csv (5 returns)

Google Sheets Link (backup): [link]

Regards,
Kumud Creations
```

## Performance Considerations

### Memory Usage
- **Impact**: CSV generation creates additional string data in memory
- **Mitigation**: Process data in single pass, avoid duplicate storage
- **Monitoring**: Log memory usage for large datasets

### Execution Time
- **Additional Overhead**: ~100-500ms for CSV generation
- **Acceptable Impact**: Minimal compared to Shopify API calls (seconds)
- **Optimization**: Efficient string concatenation for CSV building

## Security Considerations

### Data Handling
- **Principle**: Same security level as existing Google Sheets export
- **CSV Content**: Contains same sensitive business data as sheets
- **Email Security**: Relies on Gmail/Google Workspace security

### Access Control
- **Unchanged**: Same recipient resolution logic
- **Attachment Access**: Recipients can save/forward CSV files
- **Recommendation**: Ensure recipient email addresses are secure

## Backward Compatibility

### Existing Functionality
- **Preserved**: All existing Google Sheets creation and Drive storage
- **Unchanged**: Menu items, user interface, configuration
- **Enhanced**: Email notifications now include attachments

### Configuration Impact
- **No Changes**: All existing configuration variables remain the same
- **New Behavior**: Automatic CSV attachment generation
- **Fallback**: Link-only emails if CSV generation fails