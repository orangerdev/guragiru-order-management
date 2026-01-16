/**
 * InputOrder Class
 * Handles order input operations for Google Sheets
 */
class InputOrder {
  /**
   * Serves the HTML form for order input
   */
  static doGet() {
    return HtmlService.createHtmlOutputFromFile("Index")
      .setTitle("Input Order")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  /**
   * Gets all available sheets except TEMPLATE and CONFIG
   * @returns {Array} Array of sheet names
   */
  static getSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const excludedSheets = ["TEMPLATE", "CONFIG"];

    return sheets
      .filter((sheet) => !excludedSheets.includes(sheet.getName()))
      .map((sheet) => sheet.getName());
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
      throw new Error("Sheet not found: " + sheetName);
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

      if (cellValue && cellValue.toString().trim() !== "") {
        // If we were tracking a previous name, save it
        if (currentName !== null) {
          names.push({
            name: currentName,
            startRow: startRow,
            endRow: rowNum - 1,
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
        endRow: lastRow,
      });
    }

    return names;
  }

  /**
   * Submits a new order to the sheet
   * @param {Object} data - Order data containing sheetName, name, and items array
   * @returns {Object} Result object with success status and message
   */
  static submitOrder(data) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(data.sheetName);

      if (!sheet) {
        throw new Error("Sheet not found: " + data.sheetName);
      }

      if (!data.items || data.items.length === 0) {
        throw new Error("No items to submit");
      }

      const names = InputOrder.getNames(data.sheetName);
      const existingName = names.find((n) => n.name === data.name);

      if (existingName) {
        // Insert after existing name's last row
        InputOrder._insertRowsForExistingName(sheet, existingName.endRow, data);
      } else {
        // Append at the end for new name
        InputOrder._appendRowsForNewName(sheet, data);
      }

      const itemCount = data.items.length;
      const itemText = itemCount === 1 ? "item" : "items";

      return {
        success: true,
        message: `${itemCount} ${itemText} berhasil ditambahkan!`,
      };
    } catch (error) {
      return {
        success: false,
        message: "Error: " + error.message,
      };
    }
  }

  /**
   * Inserts multiple rows after an existing name's entries
   * @param {Sheet} sheet - The target sheet
   * @param {number} endRow - The last row of the existing name
   * @param {Object} data - Order data with items array
   * @private
   */
  static _insertRowsForExistingName(sheet, endRow, data) {
    // Insert rows for each item
    for (let i = 0; i < data.items.length; i++) {
      const item = data.items[i];
      const insertPosition = endRow + i;

      // Insert a new row after the last row
      sheet.insertRowAfter(insertPosition);
      const newRow = insertPosition + 1;

      // Leave column A (Name) empty for continuation rows
      sheet.getRange(newRow, 1).setValue(""); // Name column empty
      sheet.getRange(newRow, 2).setValue(item.item);
      sheet.getRange(newRow, 3).setValue(item.quantity);
      sheet.getRange(newRow, 4).setValue(item.price);
    }
  }

  /**
   * Appends multiple rows at the end of the sheet for a new name
   * @param {Sheet} sheet - The target sheet
   * @param {Object} data - Order data with items array
   * @private
   */
  static _appendRowsForNewName(sheet, data) {
    const lastRow = sheet.getLastRow();

    for (let i = 0; i < data.items.length; i++) {
      const item = data.items[i];
      const newRow = lastRow + 1 + i;

      // For new name, only fill name column on first row
      if (i === 0) {
        sheet.getRange(newRow, 1).setValue(data.name);
      } else {
        sheet.getRange(newRow, 1).setValue(""); // Empty for subsequent rows
      }

      sheet.getRange(newRow, 2).setValue(item.item);
      sheet.getRange(newRow, 3).setValue(item.quantity);
      sheet.getRange(newRow, 4).setValue(item.price);
    }
  }
}

/**
 * Entry point for web app
 */
function doGet() {
  return InputOrder.doGet();
}

/**
 * Global wrapper functions for google.script.run
 * These are required because google.script.run cannot call static class methods
 */

function getSheets() {
  return InputOrder.getSheets();
}

function getNames(sheetName) {
  return InputOrder.getNames(sheetName);
}

function submitOrder(data) {
  return InputOrder.submitOrder(data);
}
