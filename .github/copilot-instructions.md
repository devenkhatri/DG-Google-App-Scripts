# AI Agent Instructions for DG-Google-App-Scripts

## Project Overview
This repository contains Google Apps Script projects for automating Google Sheets workflows, with a focus on e-commerce integration between Shopify and Qikink platforms.

## Key Architecture Components

### KC-Qikink Integration
- **Core Logic** (`Code.js`): Entry point that handles UI menus and core API integration
- **GST Processing** (`GSTInvoices.gs.js`): GST report generation and tax calculations
- **Shopify Integration** (`ShopifyUtils.js`): Shopify API interactions and order management

### Data Flow
1. Orders flow from Shopify -> Google Sheets
2. Status updates sync between Qikink and Shopify
3. GST data is processed and exported to CSVs in Google Drive

## Development Workflow

### Local Development with clasp
```bash
clasp login        # Authenticate with Google
clasp clone <ID>   # Get script files locally
clasp pull         # Update local files
clasp push         # Deploy changes
```

### Critical Configuration
- Project settings live in `appsscript.json`
- Runtime configuration stored in "Automation" Google Sheet:
  - Shopify credentials (H5-H7)
  - Email settings (H8)
  - Drive folder ID (H9)

## Project-Specific Patterns

### Google Apps Script Conventions
- Files use `.gs.js` extension for Apps Script files
- UI elements added via custom menu "🛠️ KC Tools"
- Timezone locked to Asia/Kolkata
- V8 runtime required

### Integration Patterns
- Shopify Admin API used for order management
- Qikink API for fulfillment and tracking
- Google Drive API for GST report storage
- Sheet-based configuration for credentials

### Error Handling
- Failed API calls logged to dedicated sheet
- Rate limiting handled via exponential backoff
- Status updates verified before marking complete

## Common Tasks

### Adding New Features
1. Add menu item in `Code.js`
2. Implement handler function
3. Update sheet configuration if needed

### GST Processing
- Reports generated monthly
- Naming convention: `GST_Invoices_{YYYYMM}_{MONTH}-{YYYYMMDD}.csv`
- Tax calculations in `GSTInvoices.gs.js`

### Order Sync
- Status updates run on time trigger
- COD payment status handled separately
- RTO cases need special exception handling

## Key Files Reference
- `KC-Qikink-Integration/Code.js` - Main entry point and UI
- `KC-Qikink-Integration/GSTInvoices.gs.js` - GST logic
- `KC-Qikink-Integration/ShopifyUtils.js` - Shopify integration
- `KC-Qikink-Integration/appsscript.json` - Project config
