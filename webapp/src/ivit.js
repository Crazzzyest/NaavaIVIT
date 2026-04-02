const config = require('./config');
const { COL } = require('./columns');
const google = require('./google');
const { extractAddressFromHyperlink, mapBoligtype } = require('./utils');

async function fetchIvitData(address) {
  if (!address || address.trim().length === 0) {
    return { success: false, error: 'Tom adresse' };
  }

  const options = {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ address: address.trim() }),
  };

  if (config.ivit.webhookSecret) {
    options.headers['Authorization'] = `Bearer ${config.ivit.webhookSecret}`;
  }

  try {
    const res = await fetch(config.ivit.webhookUrl, options);
    const body = await res.json();
    if (res.ok && body.success) return body;
    return { success: false, error: body.error || 'Ukjent feil fra Webhook' };
  } catch (e) {
    return { success: false, error: 'Forespørsel feilet: ' + e.message };
  }
}

async function processIVITScraping() {
  console.log('START: processIVITScraping');

  const data = await google.getSheetData(config.sheet.name);
  if (!data || data.length < 2) {
    console.log('No data to process');
    return { processed: 0 };
  }

  let processed = 0;

  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1;
    const row = data[i];

    const scanIvitRaw = row[COL.SCAN_IVIT - 1];
    const scanIvit = scanIvitRaw === true || scanIvitRaw === 'TRUE' || scanIvitRaw === 'true' || scanIvitRaw === 1;

    if (!scanIvit) continue;

    let adresse = extractAddressFromHyperlink(row[COL.ADRESSE - 1]);
    let notater = String(row[COL.NOTATER - 1] || '');

    console.log(`  Row ${rowNum}: address="${adresse}"`);

    if (!adresse || adresse.trim() === '') {
      const newNotat = notater + (notater ? ' | ' : '') + '[iVit Feil: Tom/manglende adresse]';
      await google.updateCells(config.sheet.name, rowNum, [
        { col: COL.NOTATER, value: newNotat },
        { col: COL.SCAN_IVIT, value: false },
      ]);
      continue;
    }

    const result = await fetchIvitData(adresse);
    console.log(`  fetchIvitData result: success=${result.success}`);

    if (result.success) {
      const d = result.data;
      const updates = [];

      if (d.fakturareferanse) updates.push({ col: COL.FAKTURA_REF, value: d.fakturareferanse });
      if (d.befaring_dato) updates.push({ col: COL.BEFARING_DATO, value: d.befaring_dato });
      if (d.befaring_klokkeslett) updates.push({ col: COL.BEFARING_KL, value: d.befaring_klokkeslett });
      if (d.selger) updates.push({ col: COL.SELGER, value: d.selger });
      if (d.selger_tlf) updates.push({ col: COL.SELGER_TLF, value: d.selger_tlf });
      if (d.selger_epost) updates.push({ col: COL.SELGER_EPOST, value: d.selger_epost });

      if (d.boligtype) {
        updates.push({ col: COL.BOLIGTYPE, value: mapBoligtype(d.boligtype) });
      }
      if (d.areal_bra != null) {
        updates.push({ col: COL.AREAL, value: d.areal_bra });
      }
      if (d.antall_bygninger != null) {
        updates.push({ col: COL.ANTALL_TILLEGGSBYGG, value: Math.max(0, d.antall_bygninger - 1) });
      }
      if (d.med_markedsverdi === true || d.med_markedsverdi === 'true') {
        updates.push({ col: COL.MED_MARKEDSVERDI, value: true });
      }

      updates.push({ col: COL.SCAN_IVIT, value: false });

      await google.updateCells(config.sheet.name, rowNum, updates);

      // Trigger price calculation
      const { calculatePriceForRow } = require('./oppdrag');
      await calculatePriceForRow(rowNum);

      processed++;
      console.log(`  Success for row ${rowNum}`);
    } else {
      const errorMsg = result.error || 'Ukjent feil';
      const newNotat = notater + (notater ? ' | ' : '') + `[iVit Feil: ${errorMsg}]`;
      await google.updateCells(config.sheet.name, rowNum, [
        { col: COL.NOTATER, value: newNotat },
        { col: COL.SCAN_IVIT, value: false },
      ]);
    }
  }

  console.log(`DONE: Processed ${processed} rows`);
  return { processed };
}

module.exports = { fetchIvitData, processIVITScraping };
