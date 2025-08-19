# KC-Qikink Integration

A Google Apps Script project that integrates Kumud Creations' Shopify store with Qikink for order management and GST reporting.

## Core Features

### 1. Order Management
- Fetches orders from Qikink API
- Updates order statuses between Qikink and Shopify
- Syncs tracking information for shipments
- Handles COD payment status updates
- Maintains customer information sync

### 2. GST Management
- Generates GST reports for specified months
- Supports both last month and custom month exports
- Calculates GST breakdowns (IGST, CGST, SGST)
- Handles different tax rates based on order value
- Creates organized Google Sheets for GST reporting

### 3. Tracking & Fulfillment
- Syncs tracking numbers between platforms
- Updates shipment status automatically
- Handles courier partner information
- Manages fulfillment status in Shopify
- Processes returns and exceptions

### 4. Logging & Monitoring
- Maintains sync logs for all operations
- Tracks changes in order status
- Records fulfillment updates
- Monitors exceptions and RTO cases
- Provides audit trail for all system actions

## File Structure

### [Code.js](Code.js)
- Main entry point for the application
- Handles menu creation and UI interactions
- Contains core API integration functions
- Manages authentication with Qikink

### [GSTInvoices.gs.js](GSTInvoices.gs.js)
- Handles all GST-related functionality
- Generates GST reports
- Calculates tax breakdowns
- Manages export to Google Sheets

### [ShopifyUtils.js](ShopifyUtils.js)
- Contains Shopify-specific utilities
- Manages order fulfillment
- Handles tracking updates
- Processes Shopify API interactions

### [appsscript.json](appsscript.json)
- Project configuration file
- Defines time zone settings
- Specifies script dependencies
- Sets runtime version

## Configuration

The script requires several configuration values in a Google Sheet named "Automation":
- Shopify store details (H5)
- API access tokens (H6)
- API version (H7)
- Email settings (H8)
- Drive folder ID (H9)

## Usage

1. The script adds a menu item "🛠️ KC Tools" to your Google Sheet with options:
   - Update Statuses
   - Export GST CSV (Last Month)
   - Export GST CSV (Pick Month...)

2. GST exports are saved to:
   - Configured Google Drive folder
   - Format: `GST_Invoices_{YYYYMM}_{MONTH}-{YYYYMMDD}.csv`

3. Order sync runs automatically to:
   - Update tracking information
   - Sync delivery statuses
   - Process returns and exceptions
   - Update payment statuses for COD orders

## Dependencies

- Google Apps Script
- Shopify Admin API
- Qikink API
- Google Sheets API
- Google Drive API

## Notes

- Timezone is set to Asia/Kolkata
- Uses V8 runtime
- Requires appropriate API permissions
- Assumes specific sheet structure for operation