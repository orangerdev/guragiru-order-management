const SHEET = SpreadsheetApp.getActiveSpreadsheet();
const CONFIG_SHEET = SHEET.getSheetByName(SHEET_CONFIG);
const CONFIG_DOKU_CLIENT_ID = CONFIG_SHEET.getRange("B2").getValue();
const CONFIG_DOKU_SECRET_KEY = CONFIG_SHEET.getRange("B3").getValue();
const CONFIG_DOKU_ENVIRONMENT = CONFIG_SHEET.getRange("B4").getValue();

/**
 * Class untuk handle DOKU Payment API
 *
 * @example
 * const doku = new DokuPayment(CONFIG_DOKU_CLIENT_ID, CONFIG_DOKU_SECRET_KEY, CONFIG_DOKU_ENVIRONMENT);
 * const result = doku.generatePaymentUrl({
 *   invoiceNumber: "INV-20251122-0001",
 *   amount: 200000,
 *   customerName: "John Doe",
 *   customerPhone: "628112411915",
 *   items: [
 *     { name: "Product A", quantity: 2, price: 50000 },
 *     { name: "Product B", quantity: 1, price: 100000 }
 *   ],
 *   paymentDueDate: 60, // optional, default 60 minutes
 *   paymentMethodTypes: [] // optional, empty = all methods
 * });
 */
class DokuPayment {
  constructor(clientId, secretKey, environment = "") {
    this.clientId = clientId;
    this.secretKey = secretKey;
    this.isProduction =
      environment && environment.toLowerCase() === "production";
    this.baseUrl = this.isProduction
      ? "https://api.doku.com"
      : "https://api-sandbox.doku.com";
    this.endpoint = "/checkout/v1/payment";
  }

  /**
   * Generate UUID v4 untuk Request-Id
   * @returns {string} UUID
   */
  generateRequestId() {
    return Utilities.getUuid();
  }

  /**
   * Generate timestamp dalam format ISO 8601
   * @returns {string} Timestamp (e.g., 2020-08-11T08:45:42Z)
   */
  generateTimestamp() {
    const date = new Date();
    return Utilities.formatDate(date, "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  }

  /**
   * Generate Digest (SHA-256 hash dari request body)
   * @param {Object} body - Request body object
   * @returns {string} Base64 encoded SHA-256 hash
   */
  generateDigest(body) {
    const bodyString = JSON.stringify(body);
    const hash = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      bodyString,
      Utilities.Charset.UTF_8
    );
    return Utilities.base64Encode(hash);
  }

  /**
   * Generate Signature untuk request header
   * @param {string} clientId
   * @param {string} requestId
   * @param {string} timestamp
   * @param {string} digest
   * @returns {string} Signature dengan format HMACSHA256=xxx
   */
  generateSignature(clientId, requestId, timestamp, digest) {
    // Build signature components
    const components = [
      `Client-Id:${clientId}`,
      `Request-Id:${requestId}`,
      `Request-Timestamp:${timestamp}`,
      `Request-Target:${this.endpoint}`,
      `Digest:${digest}`,
    ];

    const signatureString = components.join("\n");

    // Calculate HMAC-SHA256
    const signature = Utilities.computeHmacSha256Signature(
      signatureString,
      this.secretKey
    );

    const base64Signature = Utilities.base64Encode(signature);

    return `HMACSHA256=${base64Signature}`;
  }

  /**
   * Generate Payment URL dari DOKU API
   * @param {Object} orderData - Data order
   * @param {string} orderData.invoiceNumber - Invoice number
   * @param {number} orderData.amount - Total amount
   * @param {string} orderData.customerName - Customer name
   * @param {string} orderData.customerPhone - Customer phone (format: 628xxx)
   * @param {Array} orderData.items - Array of items [{ name, quantity, price }]
   * @param {number} [orderData.paymentDueDate=60] - Payment due date in minutes
   * @param {Array} [orderData.paymentMethodTypes=[]] - Array of payment method types
   * @returns {Object} Result object { success, paymentUrl, tokenId, sessionId, expiredDate, error }
   */
  generatePaymentUrl(orderData) {
    try {
      // Validate required fields
      if (!orderData.invoiceNumber || !orderData.amount) {
        throw new Error("Invoice number and amount are required");
      }

      if (!orderData.customerName || !orderData.customerPhone) {
        throw new Error("Customer name and phone are required");
      }

      // Build request body
      const requestBody = {
        order: {
          amount: orderData.amount,
          invoice_number: orderData.invoiceNumber,
          currency: "IDR",
        },
        payment: {
          payment_due_date: orderData.paymentDueDate || 60,
        },
        customer: {
          name: orderData.customerName,
          phone: orderData.customerPhone,
        },
      };

      // Add line items if provided
      if (
        orderData.items &&
        Array.isArray(orderData.items) &&
        orderData.items.length > 0
      ) {
        requestBody.order.line_items = orderData.items.map((item) => ({
          name: item.name || item.nama,
          quantity: item.quantity || item.qty,
          price: item.price || item.harga,
        }));
      }

      // Add payment method types if provided
      if (
        orderData.paymentMethodTypes &&
        Array.isArray(orderData.paymentMethodTypes) &&
        orderData.paymentMethodTypes.length > 0
      ) {
        requestBody.payment.payment_method_types = orderData.paymentMethodTypes;
      }

      // Generate request headers
      const requestId = this.generateRequestId();
      const timestamp = this.generateTimestamp();
      const digest = this.generateDigest(requestBody);
      const signature = this.generateSignature(
        this.clientId,
        requestId,
        timestamp,
        digest
      );

      // Build request options
      const options = {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Client-Id": this.clientId,
          "Request-Id": requestId,
          "Request-Timestamp": timestamp,
          Signature: signature,
        },
        payload: JSON.stringify(requestBody),
        muteHttpExceptions: true,
      };

      // Log request for debugging (remove in production)
      Logger.log("DOKU Request URL: " + this.baseUrl + this.endpoint);
      Logger.log("DOKU Request Headers: " + JSON.stringify(options.headers));
      Logger.log("DOKU Request Body: " + JSON.stringify(requestBody));

      // Make API call
      const response = UrlFetchApp.fetch(this.baseUrl + this.endpoint, options);
      const responseCode = response.getResponseCode();
      const responseBody = JSON.parse(response.getContentText());

      Logger.log("DOKU Response Code: " + responseCode);
      Logger.log("DOKU Response Body: " + JSON.stringify(responseBody));

      // Handle response
      if (responseCode === 200 && responseBody.response) {
        return {
          success: true,
          paymentUrl: responseBody.response.payment.url,
          tokenId: responseBody.response.payment.token_id,
          sessionId: responseBody.response.order.session_id,
          expiredDate: responseBody.response.payment.expired_date,
          response: responseBody.response,
        };
      } else {
        // Handle error response
        const errorMessages = responseBody.error_messages ||
          responseBody.message || ["Unknown error"];
        return {
          success: false,
          error: Array.isArray(errorMessages)
            ? errorMessages.join(", ")
            : errorMessages,
          responseCode: responseCode,
          response: responseBody,
        };
      }
    } catch (error) {
      Logger.log("DOKU Error: " + error.toString());
      return {
        success: false,
        error: error.toString(),
      };
    }
  }
}
