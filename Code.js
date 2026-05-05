/* ========== CONFIG ========== */
const WEBHOOK_URL =
  "https://n8n-w3fobgi1llim.cica.sumopod.my.id/webhook/65169f52-53ec-4323-8c4f-26adf05d3370";
const SHEET_ORDER = "ORDER"; // nama sheet untuk menyimpan data
const SHEET_INVOICE = "INVOICE"; // nama sheet untuk template invoice
const SHEET_TEMP_INVOICE = "TEMP_INVOICE"; // nama sheet untuk template invoice
const SHEET_CONFIG = "CONFIG"; // nama sheet untuk menyimpan data
const OUTPUT_FOLDER_ID = "1I48VLvw1PbMfkQa3OQwHYS5iWEvyMLSu"; // ganti dengan folder ID untuk menyimpan hasil (PDF & doc copy)

/* ========== WEB APP ROUTING ========== */

/**
 * Main entry point for web app
 * Routes: ?action=pay&t={token} → payment redirect
 * Default → Order Management UI
 */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === "pay") {
    return handlePaymentRedirect(e.parameter.t);
  }

  return HtmlService.createHtmlOutputFromFile("MainAppSimple")
    .setTitle("Order Management System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handles payment link clicks: looks up token, generates fresh Doku link, redirects customer
 * @param {string} token - UUID token from payment URL
 */
function handlePaymentRedirect(token) {
  if (!token) {
    return HtmlService.createHtmlOutput(
      "<p>Link pembayaran tidak valid.</p>"
    );
  }

  try {
    const tokenData = CreateInvoice._lookupPaymentToken(token);
    if (!tokenData) {
      return HtmlService.createHtmlOutput(
        "<p>Link pembayaran tidak ditemukan atau sudah kedaluwarsa.</p>"
      );
    }

    const doku = new DokuPayment(
      CONFIG_DOKU_CLIENT_ID,
      CONFIG_DOKU_SECRET_KEY,
      CONFIG_DOKU_ENVIRONMENT
    );
    const dokuResult = doku.generatePaymentUrl({
      invoiceNumber: tokenData.invoiceId,
      amount: tokenData.amount,
      customerName: tokenData.customerName,
      customerPhone: tokenData.phone,
      items: JSON.parse(tokenData.items),
      paymentDueDate: 60,
    });

    if (!dokuResult.success) {
      return HtmlService.createHtmlOutput(
        "<p>Gagal membuat link pembayaran. Silakan coba lagi.</p>"
      );
    }

    const template = HtmlService.createTemplateFromFile("PaymentRedirect");
    template.paymentUrl = dokuResult.paymentUrl;
    return template.evaluate().setTitle("Redirect ke Pembayaran");
  } catch (err) {
    Logger.log("Payment redirect error: " + err);
    return HtmlService.createHtmlOutput(
      "<p>Terjadi kesalahan. Silakan hubungi admin.</p>"
    );
  }
}

/**
 * Include helper to import HTML files
 * Used by MainApp to load tab contents
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Test function to verify getSheets is working
 * Can be run directly from Apps Script editor
 */
function testGetSheets() {
  const sheets = getSheets();
  Logger.log("Available sheets: " + JSON.stringify(sheets));
  return sheets;
}

/* ========== INPUT ORDER WRAPPER FUNCTIONS ========== */

/**
 * Gets all available sheets for input order (excludes system sheets)
 * Wrapper for InputOrder.getSheets()
 */
function getSheets() {
  return InputOrder.getSheets();
}

/**
 * Gets names with their row ranges from a specific sheet
 * Wrapper for InputOrder.getNames()
 */
function getNames(sheetName) {
  return InputOrder.getNames(sheetName);
}

/**
 * Submits a new order to the sheet
 * Wrapper for InputOrder.submitOrder()
 */
function submitOrder(data) {
  return InputOrder.submitOrder(data);
}

function tempAuthorizeDriveAccess() {
  DriveApp.getFolders(); // This line will trigger the authorization prompt
}

function tempAuthorizeFetch(e) {
  var response = UrlFetchApp.fetch("https://guragiru.com/");
  Logger.log(response.getContentText());
}

function tempAuthorizeCreateDriveFile() {
  var folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  var file = folder.createFile("Test File", "Hello World!");
  Logger.log(file.getUrl());
}
