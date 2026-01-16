/**
 * InputOrder Class
 * Handles order input operations for Google Sheets
 */
class InputOrder {
  
  /**
   * Serves the HTML form for order input
   */
  static doGet() {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Input Order')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  /**
   * Gets all available sheets except TEMPLATE and CONFIG
   * @returns {Array} Array of sheet names
   */
  static getSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const excludedSheets = ['TEMPLATE', 'CONFIG'];
    
    return sheets
      .filter(sheet => !excludedSheets.includes(sheet.getName()))
      .map(sheet => sheet.getName());
  }
  
  /**
   * Gets names with their row ranges from a specific sheet
   * @param {string} sheetName - Name of the sheet
   * @returns {Array} Array of objects with name, startRow, and endRow
   */
  static getNames(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('Sheet not found: ' + sheetName);
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return []; // No data, only header
    }
    
    const nameColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const names = [];
    let currentName = null;
    let startRow = null;
    
    for (let i = 0; i < nameColumn.length; i++) {
      const cellValue = nameColumn[i][0];
      const rowNum = i + 2; // +2 because array is 0-indexed and we start from row 2
      
      if (cellValue && cellValue.toString().trim() !== '') {
        // If we were tracking a previous name, save it
        if (currentName !== null) {
          names.push({
            name: currentName,
            startRow: startRow,
            endRow: rowNum - 1
          });
        }
        
        // Start tracking new name
        currentName = cellValue.toString().trim();
        startRow = rowNum;
      }
    }
    
    // Don't forget the last name
    if (currentName !== null) {
      names.push({
        name: currentName,
        startRow: startRow,
        endRow: lastRow
      });
    }
    
    return names;
  }
  
  /**
   * Submits a new order to the sheet
   * @param {Object} data - Order data containing sheetName, name, item, quantity, price
   * @returns {Object} Result object with success status and message
   */
  static submitOrder(data) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(data.sheetName);
      
      if (!sheet) {
        throw new Error('Sheet not found: ' + data.sheetName);
      }
      
      const names = InputOrder.getNames(data.sheetName);
      const existingName = names.find(n => n.name === data.name);
      
      if (existingName) {
        // Insert after existing name's last row
        InputOrder._insertRowForExistingName(sheet, existingName.endRow, data);
      } else {
        // Append at the end for new name
        InputOrder._appendRowForNewName(sheet, data);
      }
      
      return {
        success: true,
        message: 'Order berhasil ditambahkan!'
      };
      
    } catch (error) {
      return {
        success: false,
        message: 'Error: ' + error.message
      };
    }
  }
  
  /**
   * Inserts a row after an existing name's entries
   * @param {Sheet} sheet - The target sheet
   * @param {number} endRow - The last row of the existing name
   * @param {Object} data - Order data
   * @private
   */
  static _insertRowForExistingName(sheet, endRow, data) {
    // Insert a new row after the last row of this name
    sheet.insertRowAfter(endRow);
    const newRow = endRow + 1;
    
    // Leave column A (Name) empty for continuation rows
    sheet.getRange(newRow, 1).setValue(''); // Name column empty
    sheet.getRange(newRow, 2).setValue(data.item);
    sheet.getRange(newRow, 3).setValue(data.quantity);
    sheet.getRange(newRow, 4).setValue(data.price);
  }
  
  /**
   * Appends a row at the end of the sheet for a new name
   * @param {Sheet} sheet - The target sheet
   * @param {Object} data - Order data
   * @private
   */
  static _appendRowForNewName(sheet, data) {
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    // For new name, fill the name column
    sheet.getRange(newRow, 1).setValue(data.name);
    sheet.getRange(newRow, 2).setValue(data.item);
    sheet.getRange(newRow, 3).setValue(data.quantity);
    sheet.getRange(newRow, 4).setValue(data.price);
  }
}

/**
 * Entry point for web app
 */
function doGet() {
  return InputOrder.doGet();
}
