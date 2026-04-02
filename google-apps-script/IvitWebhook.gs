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
 * Process new IVIT rows in the Oppdrag sheet.
 *
 * Column layout:
 *  A=Oppdragsnr  B=Dato mottatt  C=Kilde  D=Oppdragstype  E=Adresse
 *  F=Oppdragsgiver  G=Selger  H=Selger tlf  I=Selger e-post
 *  J=Megler/bestiller  K=Megler e-post  L=Faktura ref  M=Faktura sendes til
 *  N=Fakturamotaker  O=Boligtype  P=Areal (m²)  Q=Antall tilleggsbygg
 *  R=Rapporttype  S=Med markedsverdi  T=Timer  U=Pris inkl. mva
 *  V=Pris eks. mva  W=MVA-beløp  X=Avstand (km t/r)  Y=Reisekostnad eks mva
 *  Z=Reisekostnad inkl mva  AA=Sum (ferge/bom)  AB=Antall personer som deler reisekostnad
 *  AC=Status  AD=Befaring dato  AE=Befaring klokkeslett  AF=Dato statusendring
 *  AG=Timestamp (intern)  AH=Link til mappe  AI=Notater
 *
 * Only processes rows where:
 *  - C (Kilde) = "IVIT"
 *  - AD (Befaring dato) is empty (not yet scraped)
 */
function processOppdragSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Oppdrag');
  if (!sheet) {
    Logger.log('Sheet "Oppdrag" not found');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Column indices (1-based)
  const COL = {
    DATO_MOTTATT: 2,     // B
    KILDE: 3,            // C
    ADRESSE: 5,          // E
    FAKTURA_REF: 12,     // L
    MED_MARKEDSVERDI: 19,// S
    BEFARING_DATO: 30,   // AD
    BEFARING_KLOKKE: 31, // AE
    TIMESTAMP: 33,       // AG
    NOTATER: 35,         // AI
  };

  const CUTOFF = new Date('2026-03-17T08:23:00');
  let processed = 0;

  for (let row = 2; row <= lastRow; row++) {
    const kilde = sheet.getRange(row, COL.KILDE).getValue().toString().trim();
    if (kilde !== 'IVIT') continue;

    const datoMottatt = sheet.getRange(row, COL.DATO_MOTTATT).getValue();
    if (!datoMottatt) continue;

    // Parse date — handles both Date objects and strings like "17.03.2026 08:24"
    let mottattDate;
    if (datoMottatt instanceof Date) {
      mottattDate = datoMottatt;
    } else {
      const parts = datoMottatt.toString().match(/(\d{2})\.(\d{2})\.(\d{4})\s*(\d{2}):(\d{2})/);
      if (!parts) continue;
      mottattDate = new Date(parts[3], parts[2] - 1, parts[1], parts[4], parts[5]);
    }

    // Skip legacy rows before cutoff
    if (mottattDate <= CUTOFF) continue;

    // Skip already processed rows
    const timestamp = sheet.getRange(row, COL.TIMESTAMP).getValue();
    if (timestamp) continue;

    const address = sheet.getRange(row, COL.ADRESSE).getValue().toString().trim();
    if (!address) continue;

    Logger.log('Row ' + row + ': scraping "' + address + '"');

    const result = fetchIvitData(address);

    if (result.success) {
      const d = result.data;
      sheet.getRange(row, COL.BEFARING_DATO).setValue(d.befaring_dato || '');
      sheet.getRange(row, COL.BEFARING_KLOKKE).setValue(d.befaring_klokkeslett || '');
      sheet.getRange(row, COL.FAKTURA_REF).setValue(d.fakturareferanse || '');
      sheet.getRange(row, COL.MED_MARKEDSVERDI).setValue(d.med_markedsverdi || '');
      sheet.getRange(row, COL.TIMESTAMP).setValue(new Date());
      Logger.log('Row ' + row + ': OK');
      processed++;
    } else {
      // Write error to Notater column so it's visible but doesn't block re-processing
      sheet.getRange(row, COL.NOTATER).setValue('iVit feil: ' + (result.error || 'Ukjent feil'));
      Logger.log('Row ' + row + ': Error — ' + result.error);
    }

    SpreadsheetApp.flush();
    Utilities.sleep(2000);
  }

  Logger.log('Done. Processed ' + processed + ' row(s).');
}

/**
 * Test function — run this to verify the connection works.
 */
function testIvitConnection() {
  const result = fetchIvitData('Smithsgata 5, 6100 VOLDA');
  Logger.log('Test result: ' + JSON.stringify(result, null, 2));
}
