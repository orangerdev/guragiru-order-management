/* ============================ */

/**
 * Serve the HTML form
 */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile("index");
  t.timestamp = new Date();
  return t
    .evaluate()
    .setTitle("Input Invoice — Web App")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME) // legacy but ok
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Include helper to import partial HTML files (if any)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function saveOrder(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(SHEET_ORDER);

    // Save order data - one row per item with same Invoice ID
    data.items.forEach((item, index) => {
      const subtotal = item.qty * item.harga;
      const newRowIndex = orderSheet.getLastRow() + 1;

      orderSheet.appendRow([
        new Date(), // A: Date
        data.invoiceId, // B: Invoice ID
        data.nama, // C: Name
        data.hp, // D: Phone
        item.nama, // E: Item
        item.qty, // F: Qty
        item.harga, // G: Unit Price
        subtotal, // H: SubTotal
      ]);
    });

    return true;
  } catch (error) {
    Logger.log("Error saving order:", error);
    throw new Error("Gagal menyimpan data: " + error.message);
  }
}

/**
 * Generate invoice document & PDF directly from form data.
 * Expects data object:
 * {
 *   nama: string,
 *   hp: string,
 *   items: [ { nama, qty, harga }, ... ]
 * }
 * Returns PDF URL
 */
function generateInvoice(data) {
  try {
    Logger.log("generateInvoice called with data:", JSON.stringify(data)); // Debug log

    // First save the data
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(SHEET_ORDER);

    // Generate Invoice ID first
    const invoiceId = createInvoiceId();
    const currentDate = new Date();

    // Save order data - one row per item with same Invoice ID
    data.items.forEach((item, index) => {
      const subtotal = item.qty * item.harga;
      const newRowIndex = orderSheet.getLastRow() + 1;

      orderSheet.appendRow([
        currentDate, // A: Date
        invoiceId, // B: Invoice ID
        data.nama, // C: Name
        data.hp, // D: Phone
        item.nama, // E: Item
        item.qty, // F: Qty
        item.harga, // G: Unit Price
        subtotal, // H: SubTotal
      ]);

      // Format phone number column as text to preserve leading zero
      const phoneCell = orderSheet.getRange(newRowIndex, 4); // Column D (Phone)
      phoneCell.setNumberFormat("@"); // @ means text format
      phoneCell.setValue(data.hp);
    });

    // Generate invoice
    const template = ss.getSheetByName(SHEET_INVOICE);
    const tempSheet = template
      .copyTo(ss)
      .setName(SHEET_TEMP_INVOICE + "_" + new Date().getTime());

    // 1. G7 untuk data tanggal
    tempSheet.getRange("G7").setValue(currentDate);

    // 2. G9 untuk nomor invoice
    tempSheet.getRange("G9").setValue(invoiceId);

    // 3. B14 untuk nama pemesan
    tempSheet.getRange("B14").setValue(data.nama);

    // 4. Clear semua data dari B18 hingga G28
    tempSheet.getRange("B18:G28").clearContent();

    // 5. Isi data item dimulai dari B18
    let startRow = 18;
    let total = 0;

    Logger.log("Data items:", JSON.stringify(data.items)); // Debug log

    data.items.forEach((item, i) => {
      let row = startRow + i;
      let subtotal = item.qty * item.harga;
      total += subtotal;

      Logger.log(
        `Item ${i}: ${item.nama}, Qty: ${item.qty}, Harga: ${item.harga}, Row: ${row}`
      ); // Debug log

      // Kolom B untuk nama produk
      tempSheet.getRange(row, 2).setValue(item.nama);
      // Kolom E untuk kuantiti
      tempSheet.getRange(row, 5).setValue(item.qty);
      // Kolom F untuk harga satuan
      tempSheet.getRange(row, 6).setValue(item.harga);
      // Kolom G untuk harga subtotal per line product
      tempSheet.getRange(row, 7).setValue(subtotal);
    });

    // Force flush all pending operations
    SpreadsheetApp.flush();

    // Try PNG export first with proper parameters
    const folderId = OUTPUT_FOLDER_ID;
    const folder = DriveApp.getFolderById(folderId);

    const spreadsheet = tempSheet.getParent();
    const sheetId = tempSheet.getSheetId();

    // Method 1: Try PNG export with optimized parameters
    const pngUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=png&gid=${sheetId}&scale=2&fzr=false&fzc=false`;

    Logger.log("Attempting PNG export with URL:", pngUrl);

    const pngResponse = UrlFetchApp.fetch(pngUrl, {
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken(),
      },
      muteHttpExceptions: true,
    });

    Logger.log("PNG Export response code:", pngResponse.getResponseCode());

    if (pngResponse.getResponseCode() === 200) {
      // PNG export successful
      const pngBlob = pngResponse.getBlob();
      const pngFileName = `${invoiceId}_${data.nama.replace(
        /[^a-zA-Z0-9]/g,
        "_"
      )}.png`;
      pngBlob.setName(pngFileName);

      const pngFile = folder.createFile(pngBlob);
      const pngFileUrl = pngFile.getUrl();
      const pngMimeType = pngBlob.getContentType();

      Logger.log("PNG export successful");

      var finalFileUrl = pngFileUrl;
      var finalMimeType = pngMimeType;
      var finalFileName = pngFileName;
    } else {
      // Fallback to PDF export
      Logger.log("PNG export failed, falling back to PDF");

      const pdfUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf&gid=${sheetId}&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;

      const pdfResponse = UrlFetchApp.fetch(pdfUrl, {
        headers: {
          Authorization: "Bearer " + ScriptApp.getOAuthToken(),
        },
        muteHttpExceptions: true,
      });

      if (pdfResponse.getResponseCode() !== 200) {
        throw new Error(
          `Export failed. PNG: ${pngResponse.getResponseCode()}, PDF: ${pdfResponse.getResponseCode()}`
        );
      }

      const pdfBlob = pdfResponse.getBlob();
      const pdfFileName = `${invoiceId}_${data.nama.replace(
        /[^a-zA-Z0-9]/g,
        "_"
      )}.pdf`;
      pdfBlob.setName(pdfFileName);

      const pdfFile = folder.createFile(pdfBlob);
      const pdfFileUrl = pdfFile.getUrl();
      const pdfMimeType = pdfBlob.getContentType();

      var finalFileUrl = pdfFileUrl;
      var finalMimeType = pdfMimeType;
      var finalFileName = pdfFileName;
    }

    // Wait a bit before cleanup to ensure file is fully created
    Utilities.sleep(2000);

    // Generate DOKU payment URL
    let paymentUrl = "";
    try {
      const doku = new DokuPayment(
        CONFIG_DOKU_CLIENT_ID,
        CONFIG_DOKU_SECRET_KEY,
        CONFIG_DOKU_ENVIRONMENT
      );

      const normalizedPhone = normalizePhoneNumber(data.hp);
      const dokuResult = doku.generatePaymentUrl({
        invoiceNumber: invoiceId,
        amount: total,
        customerName: data.nama,
        customerPhone: normalizedPhone,
        items: data.items.map((item) => ({
          name: item.nama,
          quantity: item.qty,
          price: item.harga,
        })),
        paymentDueDate: 60,
      });

      if (dokuResult.success) {
        paymentUrl = dokuResult.paymentUrl;
        Logger.log("DOKU Payment URL generated: " + paymentUrl);
      } else {
        Logger.log("Failed to generate DOKU payment URL: " + dokuResult.error);
      }
    } catch (dokuError) {
      Logger.log("DOKU payment URL generation error: " + dokuError);
      // Continue without payment URL
    }

    // Send webhook notification
    try {
      sendWebhookNotification(
        finalFileUrl,
        data.nama,
        data.hp,
        total,
        invoiceId,
        finalMimeType,
        finalFileName,
        data.items, // pass items array for formatting
        paymentUrl
      );
    } catch (webhookError) {
      console.error(
        "Warning: Webhook failed but invoice was created:",
        webhookError
      );
      // Don't throw error here, invoice was successfully created
    }

    // Clean up temporary sheet
    ss.deleteSheet(tempSheet);

    return finalFileUrl;
  } catch (error) {
    console.error("Error generating invoice:", error);
    throw new Error("Gagal membuat invoice: " + error.message);
  }
}

/**
 * Generate simple InvoiceID, e.g. INV-20250916-0001
 * Uses a counter in script properties to keep uniqueness across runs.
 */
function createInvoiceId() {
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
 * Examples:
 * 08112411915 -> 628112411915
 * +628112411915 -> 628112411915
 * 628112411915 -> 628112411915
 */
function normalizePhoneNumber(phoneNumber) {
  if (!phoneNumber) return "";

  // Remove all non-numeric characters except +
  let cleanNumber = phoneNumber.replace(/[^\d+]/g, "");

  // Remove + if exists
  cleanNumber = cleanNumber.replace(/^\+/, "");

  // If starts with 0, replace with 62
  if (cleanNumber.startsWith("0")) {
    cleanNumber = "62" + cleanNumber.substring(1);
  }

  // If doesn't start with 62, add 62 prefix (assuming Indonesian number)
  if (!cleanNumber.startsWith("62")) {
    cleanNumber = "62" + cleanNumber;
  }

  return cleanNumber;
}

/**
 * Send webhook notification after invoice is generated
 */
function sendWebhookNotification(
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
  try {
    const normalizedPhone = normalizePhoneNumber(phoneNumber);

    // Format item list string
    let formattedItems = "";
    if (items && Array.isArray(items) && items.length > 0) {
      formattedItems = items
        .map((item) => {
          const totalHarga = formatCurrency(item.qty * item.harga);
          return `- ${item.nama} x ${item.qty}, ${totalHarga}`;
        })
        .join("\n");
    }

    const payload = {
      file_url: imageFileUrl,
      customer_name: customerName,
      phone_number: normalizedPhone,
      total_amount: formatCurrency(totalAmount),
      invoice_id: invoiceId,
      mime_type: mimeType,
      file_name: fileName,
      items: formattedItems,
      payment_url: paymentUrl || "",
    };

    console.log("Sending webhook with payload:", JSON.stringify(payload));

    const options = {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      payload: JSON.stringify(payload),
    };

    const response = UrlFetchApp.fetch(WEBHOOK_URL, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    console.log(`Webhook response: ${responseCode} - ${responseText}`);

    if (responseCode < 200 || responseCode >= 300) {
      throw new Error(
        `Webhook failed with status ${responseCode}: ${responseText}`
      );
    }

    return true;
  } catch (error) {
    console.error("Webhook error:", error);
    throw error;
  }
}

/**
 * Format number to currency (Rp) - server side
 */
function formatCurrency(number) {
  number = Number(number) || 0;
  // Format with thousand separators
  const parts = number
    .toFixed(0)
    .toString()
    .replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return "Rp " + parts;
}

/* ========== Utilities for client-side include ========== */
function getScriptConfig() {
  return {
    sheetName: SHEET_ORDER,
    templateId: TEMPLATE_DOC_ID,
    outputFolderId: OUTPUT_FOLDER_ID,
  };
}
