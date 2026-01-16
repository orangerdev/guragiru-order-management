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
 * Serves the main app with tab navigation
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("MainAppSimple")
    .setTitle("Order Management System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
