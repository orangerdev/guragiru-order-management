/**
 * CreateInvoice Class
 * Handles invoice creation from existing event sheet data
 */
class CreateInvoice {
  /**
   * Gets all available sheets for invoice creation
   * Excludes system sheets: TEMPLATE, CONFIG, ORDER, LOG, INVOICE
   * @returns {Array} Array of sheet names
   */
  static getAvailableSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const excludedSheets = ["TEMPLATE", "CONFIG", "ORDER", "LOG", "INVOICE"];

    return sheets
      .filter((sheet) => !excludedSheets.includes(sheet.getName()))
      .map((sheet) => sheet.getName())
      .sort();
  }

  /**
   * Gets unique customer names from a specific sheet
   * @param {string} sheetName - Name of the sheet
   * @returns {Array} Array of unique customer names
   */
  static getCustomersFromSheet(sheetName) {
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
    const uniqueNames = [];

    for (let i = 0; i < nameColumn.length; i++) {
      const cellValue = nameColumn[i][0];
      if (cellValue && cellValue.toString().trim() !== "") {
        const name = cellValue.toString().trim();
        if (!uniqueNames.includes(name)) {
          uniqueNames.push(name);
        }
      }
    }

    return uniqueNames.sort();
  }

  /**
   * Gets all items for a specific customer from a sheet
   * @param {string} sheetName - Name of the sheet
   * @param {string} customerName - Customer name
   * @returns {Array} Array of item objects with: item, quantity, price
   */
  static getCustomerItems(sheetName, customerName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error("Sheet not found: " + sheetName);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }

    // Get all data (columns: Name, Item, Quantity, Price)
    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const items = [];
    let isCustomerSection = false;

    for (let i = 0; i < data.length; i++) {
      const [name, item, quantity, price] = data[i];

      // Check if this row starts a new customer section
      if (name && name.toString().trim() !== "") {
        const currentName = name.toString().trim();
        isCustomerSection = currentName === customerName;
      }

      // If we're in the correct customer section and row has item data
      if (isCustomerSection && item && item.toString().trim() !== "") {
        items.push({
          item: item.toString().trim(),
          quantity: Number(quantity) || 0,
          price: Number(price) || 0,
        });
      }

      // Stop if we hit a new customer name (and we were already collecting)
      if (
        isCustomerSection &&
        name &&
        name.toString().trim() !== "" &&
        name.toString().trim() !== customerName
      ) {
        break;
      }
    }

    return items;
  }

  /**
   * Generate invoice document & PDF from selected items
   * @param {Object} data - Invoice data
   * {
   *   sheetName: string,
   *   customerName: string,
   *   phoneNumber: string,   *   discount: number (optional),
   *   shipping: number (optional),   *   selectedItems: [{item, quantity, price}, ...]
   * }
   * @returns {string} PDF/PNG URL
   */
  static generateInvoiceFromSheet(data) {
    try {
      Logger.log(
        "generateInvoiceFromSheet called with data:",
        JSON.stringify(data)
      );

      // Debug logging
      Logger.log("data.selectedItems:", data.selectedItems);
      Logger.log("data.selectedItems type:", typeof data.selectedItems);
      Logger.log(
        "data.selectedItems length:",
        data.selectedItems ? data.selectedItems.length : "undefined"
      );

      if (!data.selectedItems || data.selectedItems.length === 0) {
        throw new Error("No items selected for invoice");
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const orderSheet = ss.getSheetByName(SHEET_ORDER);

      // Generate Invoice ID
      const invoiceId = CreateInvoice._createInvoiceId();
      const currentDate = new Date();

      // Save to ORDER sheet - one row per item with same Invoice ID
      data.selectedItems.forEach((item, index) => {
        const subtotal = item.quantity * item.price;
        const newRowIndex = orderSheet.getLastRow() + 1;

        orderSheet.appendRow([
          currentDate, // A: Date
          invoiceId, // B: Invoice ID
          data.customerName, // C: Name
          data.phoneNumber, // D: Phone
          item.item, // E: Item
          item.quantity, // F: Qty
          item.price, // G: Unit Price
          subtotal, // H: SubTotal
        ]);

        // Format phone number column as text to preserve leading zero
        const phoneCell = orderSheet.getRange(newRowIndex, 4); // Column D
        phoneCell.setNumberFormat("@");
        phoneCell.setValue(data.phoneNumber);
      });

      // Generate invoice document
      const template = ss.getSheetByName(SHEET_INVOICE);
      const tempSheet = template
        .copyTo(ss)
        .setName(SHEET_TEMP_INVOICE + "_" + new Date().getTime());

      // Fill invoice template
      tempSheet.getRange("G7").setValue(currentDate); // Date
      tempSheet.getRange("G9").setValue(invoiceId); // Invoice ID
      tempSheet.getRange("B14").setValue(data.customerName); // Customer name

      // Clear previous data
      tempSheet.getRange("B18:G28").clearContent();

      // Fill items starting from row 18
      let startRow = 18;
      let subtotal = 0;

      data.selectedItems.forEach((item, i) => {
        let row = startRow + i;
        let itemSubtotal = item.quantity * item.price;
        subtotal += itemSubtotal;

        tempSheet.getRange(row, 2).setValue(item.item); // Column B: Item
        tempSheet.getRange(row, 5).setValue(item.quantity); // Column E: Qty
        tempSheet.getRange(row, 6).setValue(item.price); // Column F: Price
        tempSheet.getRange(row, 7).setValue(itemSubtotal); // Column G: Subtotal
      });

      // Get discount and shipping values (optional)
      const discount = Number(data.discount) || 0;
      const shipping = Number(data.shipping) || 0;

      // Calculate final total
      const finalTotal = subtotal - discount + shipping;

      // Add Subtotal label and value
      tempSheet.getRange(30, 7).setValue(subtotal);

      // Add Discount if exists
      if (discount > 0) {
        orderSheet.appendRow([
          currentDate, // A: Date
          invoiceId, // B: Invoice ID
          data.customerName, // C: Name
          data.phoneNumber, // D: Phone
          "Diskon", // E: Item
          1, // F: Qty
          discount, // G: Unit Price
          -discount, // H: SubTotal
        ]);
        tempSheet.getRange(31, 7).setValue(-discount);
      }

      // Add Shipping if exists
      if (shipping > 0) {
        orderSheet.appendRow([
          currentDate, // A: Date
          invoiceId, // B: Invoice ID
          data.customerName, // C: Name
          data.phoneNumber, // D: Phone
          "Ongkir", // E: Item
          1, // F: Qty
          shipping, // G: Unit Price
          shipping, // H: SubTotal
        ]);
        tempSheet.getRange(29, 7).setValue(shipping);
      }

      // Add Total
      tempSheet.getRange(32, 7).setValue(finalTotal);

      SpreadsheetApp.flush();

      // Export as PNG or PDF
      const folderId = OUTPUT_FOLDER_ID;
      const folder = DriveApp.getFolderById(folderId);
      const spreadsheet = tempSheet.getParent();
      const sheetId = tempSheet.getSheetId();

      // Try PNG export first
      const pngUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=png&gid=${sheetId}&scale=2&fzr=false&fzc=false`;

      const pngResponse = UrlFetchApp.fetch(pngUrl, {
        headers: {
          Authorization: "Bearer " + ScriptApp.getOAuthToken(),
        },
        muteHttpExceptions: true,
      });

      let finalFileUrl, finalMimeType, finalFileName;

      if (pngResponse.getResponseCode() === 200) {
        // PNG export successful
        const pngBlob = pngResponse.getBlob();
        const pngFileName = `${invoiceId}_${data.customerName.replace(
          /[^a-zA-Z0-9]/g,
          "_"
        )}.png`;
        pngBlob.setName(pngFileName);

        const pngFile = folder.createFile(pngBlob);
        finalFileUrl = pngFile.getUrl();
        finalMimeType = pngBlob.getContentType();
        finalFileName = pngFileName;
      } else {
        // Fallback to PDF
        const pdfUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf&gid=${sheetId}&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;

        const pdfResponse = UrlFetchApp.fetch(pdfUrl, {
          headers: {
            Authorization: "Bearer " + ScriptApp.getOAuthToken(),
          },
          muteHttpExceptions: true,
        });

        if (pdfResponse.getResponseCode() !== 200) {
          throw new Error("Failed to export invoice document");
        }

        const pdfBlob = pdfResponse.getBlob();
        const pdfFileName = `${invoiceId}_${data.customerName.replace(
          /[^a-zA-Z0-9]/g,
          "_"
        )}.pdf`;
        pdfBlob.setName(pdfFileName);

        const pdfFile = folder.createFile(pdfBlob);
        finalFileUrl = pdfFile.getUrl();
        finalMimeType = pdfBlob.getContentType();
        finalFileName = pdfFileName;
      }

      Utilities.sleep(2000);

      // add shipping and discount to items array
      if (discount > 0) {
        data.selectedItems.push({
          item: "Diskon",
          quantity: 1,
          price: -discount,
        });
      }
      if (shipping > 0) {
        data.selectedItems.push({
          item: "Ongkir",
          quantity: 1,
          price: shipping,
        });
      }

      // Generate secure payment token (Doku link is created on-demand when customer clicks)
      let paymentUrl = "";
      try {
        paymentUrl = CreateInvoice._generatePaymentToken(invoiceId, {
          customerName: data.customerName,
          phone: CreateInvoice._normalizePhoneNumber(data.phoneNumber),
          amount: finalTotal,
          items: data.selectedItems.map((item) => ({
            name: item.item,
            quantity: item.quantity,
            price: item.price,
          })),
        });
        Logger.log("Payment token URL generated: " + paymentUrl);
      } catch (tokenError) {
        Logger.log("Payment token generation error: " + tokenError);
      }

      // Send webhook notification
      try {
        CreateInvoice._sendWebhookNotification(
          finalFileUrl,
          data.customerName,
          data.phoneNumber,
          finalTotal,
          invoiceId,
          finalMimeType,
          finalFileName,
          data.selectedItems,
          paymentUrl
        );
      } catch (webhookError) {
        Logger.log("Webhook failed but invoice was created:", webhookError);
      }

      // Clean up temporary sheet
      ss.deleteSheet(tempSheet);

      return finalFileUrl;
    } catch (error) {
      Logger.log("Error generating invoice:", error);
      throw new Error("Gagal membuat invoice: " + error.message);
    }
  }

  /**
   * Generate simple InvoiceID, e.g. INV-20250116-0001
   * @private
   */
  static _createInvoiceId() {
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyyMMdd"
    );
    let counterKey = "counter_" + today;
    let counter = Number(props.getProperty(counterKey) || "0");
    counter += 1;
    props.setProperty(counterKey, String(counter));
    const formatted = ("0000" + counter).slice(-4);
    return `INV-${today}-${formatted}`;
  }

  /**
   * Normalize Indonesian phone number to international format
   * @private
   */
  static _normalizePhoneNumber(phoneNumber) {
    if (!phoneNumber) return "";

    let cleanNumber = phoneNumber.replace(/[^\d+]/g, "");
    cleanNumber = cleanNumber.replace(/^\+/, "");

    if (cleanNumber.startsWith("0")) {
      cleanNumber = "62" + cleanNumber.substring(1);
    }

    if (!cleanNumber.startsWith("62")) {
      cleanNumber = "62" + cleanNumber;
    }

    return cleanNumber;
  }

  /**
   * Send webhook notification after invoice is generated
   * @private
   */
  static _sendWebhookNotification(
    imageFileUrl,
    customerName,
    phoneNumber,
    totalAmount,
    invoiceId,
    mimeType,
    fileName,
    items,
    paymentUrl
  ) {
    const normalizedPhone = CreateInvoice._normalizePhoneNumber(phoneNumber);

    let formattedItems = "";
    if (items && Array.isArray(items) && items.length > 0) {
      formattedItems = items
        .map((item) => {
          const totalHarga = CreateInvoice._formatCurrency(
            item.quantity * item.price
          );
          return `- ${item.item} x ${item.quantity}, ${totalHarga}`;
        })
        .join("\n");
    }

    const payload = {
      file_url: imageFileUrl,
      customer_name: customerName,
      phone_number: normalizedPhone,
      total_amount: CreateInvoice._formatCurrency(totalAmount),
      invoice_id: invoiceId,
      mime_type: mimeType,
      file_name: fileName,
      items: formattedItems,
      payment_url: paymentUrl || "",
    };

    const options = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      payload: JSON.stringify(payload),
    };

    const response = UrlFetchApp.fetch(WEBHOOK_URL, options);
    const responseCode = response.getResponseCode();

    if (responseCode < 200 || responseCode >= 300) {
      throw new Error(`Webhook failed with status ${responseCode}`);
    }

    return true;
  }

  /**
   * Format number to currency (Rp)
   * @private
   */
  static _formatCurrency(number) {
    number = Number(number) || 0;
    const parts = number
      .toFixed(0)
      .toString()
      .replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    return "Rp " + parts;
  }

  /**
   * Generates a UUID token, stores invoice data in PAYMENT_TOKENS sheet,
   * and returns the web app payment URL with the token as parameter.
   * @private
   */
  static _generatePaymentToken(invoiceId, invoiceData) {
    const token = CreateInvoice._createUUID();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let tokenSheet = ss.getSheetByName("PAYMENT_TOKENS");
    if (!tokenSheet) {
      tokenSheet = ss.insertSheet("PAYMENT_TOKENS");
      tokenSheet.appendRow([
        "Token",
        "InvoiceID",
        "CustomerName",
        "Phone",
        "Amount",
        "Items",
        "CreatedAt",
      ]);
    }

    tokenSheet.appendRow([
      token,
      invoiceId,
      invoiceData.customerName,
      invoiceData.phone,
      invoiceData.amount,
      JSON.stringify(invoiceData.items),
      new Date().toISOString(),
    ]);

    const webAppUrl = CreateInvoice._getWebAppUrl();
    return `${webAppUrl}?action=pay&t=${token}`;
  }

  /**
   * Looks up a payment token in the PAYMENT_TOKENS sheet.
   * Returns the associated invoice data object, or null if not found.
   * @private
   */
  static _lookupPaymentToken(token) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tokenSheet = ss.getSheetByName("PAYMENT_TOKENS");
    if (!tokenSheet) return null;

    const lastRow = tokenSheet.getLastRow();
    if (lastRow <= 1) return null;

    const data = tokenSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (const row of data) {
      if (row[0] === token) {
        return {
          token: row[0],
          invoiceId: row[1],
          customerName: row[2],
          phone: row[3],
          amount: row[4],
          items: row[5],
          createdAt: row[6],
        };
      }
    }
    return null;
  }

  /**
   * Returns the web app base URL from CONFIG sheet cell B5.
   * @private
   */
  static _getWebAppUrl() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = ss.getSheetByName(SHEET_CONFIG);
    const url = config.getRange("B5").getValue();
    if (!url) throw new Error("Web app URL not configured in CONFIG B5");
    return url.toString().trim();
  }

  /**
   * Generates a random UUID v4.
   * @private
   */
  static _createUUID() {
    return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, function (c) {
      const r = (Math.random() * 16) | 0;
      const v = c === "x" ? r : (r & 0x3) | 0x8;
      return v.toString(16);
    });
  }
}

/**
 * Global wrapper functions for google.script.run
 */
function getAvailableSheets() {
  return CreateInvoice.getAvailableSheets();
}

function getCustomersFromSheet(sheetName) {
  return CreateInvoice.getCustomersFromSheet(sheetName);
}

function getCustomerItems(sheetName, customerName) {
  return CreateInvoice.getCustomerItems(sheetName, customerName);
}

function generateInvoiceFromSheet(data) {
  return CreateInvoice.generateInvoiceFromSheet(data);
}
