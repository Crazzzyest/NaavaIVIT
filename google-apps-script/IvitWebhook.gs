/**
 * iVit Webhook — Google Apps Script
 *
 * Sends an address to the iVit scraper running on Sliplane,
 * receives befaring/oppdrag data back.
 *
 * SETUP:
 * 1. Set your Sliplane webhook URL below (WEBHOOK_URL)
 * 2. Optionally set WEBHOOK_SECRET if you configured one on the server
 * 3. Call fetchIvitData("Smithsgata 5, 6100 VOLDA") from your script or a trigger
 */

// ===== CONFIGURATION =====
const WEBHOOK_URL = 'https://YOUR-APP.sliplane.app/webhook'; // <-- Replace with your Sliplane URL
const WEBHOOK_SECRET = ''; // <-- Set if you configured WEBHOOK_SECRET on the server

/**
 * Fetch oppdrag data from iVit for a given address.
 *
 * @param {string} address - The property address to look up
 * @returns {Object} The scraped data or error info
 */
function fetchIvitData(address) {
  if (!address || address.trim().length === 0) {
    throw new Error('Address is required');
  }

  const payload = JSON.stringify({ address: address.trim() });

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: payload,
    muteHttpExceptions: true,
    // Scraping can take time — set a generous timeout (3 minutes)
    // Note: Apps Script has a hard limit of ~6 minutes for URL fetches
  };

  // Add authorization header if secret is configured
  if (WEBHOOK_SECRET) {
    options.headers = {
      'Authorization': 'Bearer ' + WEBHOOK_SECRET,
    };
  }

  try {
    const response = UrlFetchApp.fetch(WEBHOOK_URL, options);
    const statusCode = response.getResponseCode();
    const body = JSON.parse(response.getContentText());

    if (statusCode === 200 && body.success) {
      Logger.log('iVit data received for: ' + address);
      Logger.log(JSON.stringify(body.data, null, 2));
      return body;
    } else {
      Logger.log('iVit error: ' + (body.error || 'Unknown error'));
      return body;
    }
  } catch (e) {
    Logger.log('Request failed: ' + e.message);
    return {
      success: false,
      error: 'Request failed: ' + e.message,
      address: address,
    };
  }
}

/**
 * Example: Write iVit data to a Google Sheet.
 * Call this from a trigger or manually.
 *
 * Expects a sheet named "Oppdrag" with addresses in column A.
 * Writes results to columns B-E.
 */
function processOppdragSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Oppdrag');
  if (!sheet) {
    Logger.log('Sheet "Oppdrag" not found');
    return;
  }

  const lastRow = sheet.getLastRow();
  // Headers: A=Address, B=Befaring Dato, C=Befaring Klokkeslett, D=Fakturareferanse, E=Med Markedsverdi, F=Status

  for (let row = 2; row <= lastRow; row++) {
    const address = sheet.getRange(row, 1).getValue();
    const status = sheet.getRange(row, 6).getValue();

    // Skip already processed rows
    if (!address || status === 'OK' || status === 'Error') continue;

    Logger.log('Processing row ' + row + ': ' + address);
    sheet.getRange(row, 6).setValue('Processing...');
    SpreadsheetApp.flush();

    const result = fetchIvitData(address);

    if (result.success) {
      const d = result.data;
      sheet.getRange(row, 2).setValue(d.befaring_dato || '');
      sheet.getRange(row, 3).setValue(d.befaring_klokkeslett || '');
      sheet.getRange(row, 4).setValue(d.fakturareferanse || '');
      sheet.getRange(row, 5).setValue(d.med_markedsverdi || '');
      sheet.getRange(row, 6).setValue('OK');
    } else {
      sheet.getRange(row, 6).setValue('Error: ' + (result.error || 'Unknown'));
    }

    // Small delay between requests to avoid overwhelming the server
    Utilities.sleep(2000);
  }

  Logger.log('Done processing all rows.');
}

/**
 * Test function — run this to verify the connection works.
 */
function testIvitConnection() {
  const result = fetchIvitData('Smithsgata 5, 6100 VOLDA');
  Logger.log('Test result: ' + JSON.stringify(result, null, 2));
}
