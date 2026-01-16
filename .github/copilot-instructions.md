# Google Apps Script - Order Management System

## Project Overview
This workspace contains a Google Apps Script project that manages order data through Google Sheets integration.

## Important Constraints

### Protected Sheets
The following sheets are **READ-ONLY** and must **NEVER** be modified by the script:
- `TEMPLATE` - Contains template structures for order data
- `CONFIG` - Contains configuration settings

**CRITICAL**: These sheets should only be modified manually by the sheet owner. Any code changes must ensure these sheets are only read from, never written to.

### System Sheets
The following sheets are for system use and should be **EXCLUDED** from user-facing dropdowns and selections:
- `ORDER` - Stores generated invoice data
- `LOG` - Contains system logs
- `INVOICE` - Invoice template for document generation
- `TEMPLATE` - Template structures (also protected)
- `CONFIG` - Configuration settings (also protected)

**When displaying available sheets to users** (e.g., for event selection), these sheets must be filtered out.

## Development Guidelines

### Google Apps Script Best Practices
- Use `SpreadsheetApp.getActiveSpreadsheet()` to access the spreadsheet
- Use `getSheetByName()` to access specific sheets
- Always validate sheet names before operations
- Implement proper error handling for sheet operations

### Code Structure
- Keep functions modular and focused on single responsibilities
- Document all functions with clear JSDoc comments
- Use meaningful variable names that reflect the data they hold

### Sheet Operations
- **Reading from TEMPLATE/CONFIG**: ✅ Allowed
  ```javascript
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
  const configData = configSheet.getDataRange().getValues();
  ```

- **Writing to TEMPLATE/CONFIG**: ❌ Strictly Forbidden
  ```javascript
  // NEVER do this with TEMPLATE or CONFIG sheets:
  configSheet.getRange('A1').setValue('new value'); // ❌
  ```

- **Working with other sheets**: ✅ Full read/write access for order management

### Order Data Management
- Focus on creating, reading, updating, and deleting order records
- Maintain data integrity across operations
- Implement validation before writing data
- Use the TEMPLATE sheet as a reference for data structure
- Use the CONFIG sheet for retrieving configuration values

## Testing
- Always test changes with a copy of the spreadsheet first
- Verify that TEMPLATE and CONFIG sheets remain unchanged after script execution
- Test edge cases for order data operations

## Deployment
- Use `clasp` for deployment (already configured)
- Test in development environment before pushing to production
- Document any changes to sheet structure or script behavior
