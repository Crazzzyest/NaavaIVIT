const config = require('./config');
const google = require('./google');
const { COL } = require('./columns');
const { formatDate, extractAddressFromHyperlink, parseDateString } = require('./utils');

// Oppdrag counted as "invoiced" — i.e. a faktura has been sent to regnskap.
const INVOICED_STATUSES = ['Fakturert', 'Oppdrag fullført'];

// ============================================================
// Fakturalogg is RECONSTRUCTED from the sheet, not from a send-log.
//
// Why: the faktura email can be sent by either the webapp OR the Apps Script
// (which runs the automation in production). A send-log would only capture
// webapp sends and miss everything Apps Script did. Instead we derive a
// readable faktura for every oppdrag that has reached an invoiced status.
// This is always complete and shows the actual data (including the comment
// to regnskap) as plain, readable text.
// ============================================================

function reconstruct(row, rowNum) {
  const { buildFakturaText } = require('./emails');
  return {
    id: rowNum,
    oppdragsnr: row[COL.OPPDRAGSNR - 1] || '',
    adresse: extractAddressFromHyperlink(row[COL.ADRESSE - 1]),
    status: row[COL.STATUS - 1] || '',
    tidspunkt: row[COL.DATO_STATUSENDRING - 1] || row[COL.DATO_MOTTATT - 1] || '',
    mottaker: config.email.accountantEmail,
    kommentar: String(row[COL.KOMMENTAR_REGNSKAP - 1] || '').trim(),
    text: buildFakturaText(row),
  };
}

async function list(limit = 300) {
  if (config.demoMode) {
    const { buildFakturaText } = require('./emails');
    const demoRow = new Array(43).fill('');
    demoRow[COL.OPPDRAGSNR - 1] = 'NT-202606-001';
    demoRow[COL.ADRESSE - 1] = 'Demogata 1, 0001 Demo';
    demoRow[COL.STATUS - 1] = 'Fakturert';
    demoRow[COL.DATO_STATUSENDRING - 1] = formatDate(new Date(), 'dd.MM.yyyy HH:mm');
    demoRow[COL.SELGER - 1] = 'Demo Selger';
    demoRow[COL.PRIS_EKS - 1] = 8000;
    demoRow[COL.PRODUKTNUMMER - 1] = '5';
    demoRow[COL.RAPPORTTYPE - 1] = 'Tilstandsrapport';
    demoRow[COL.KOMMENTAR_REGNSKAP - 1] = 'Faktureres samlet med naboeiendommen.';
    return [reconstruct(demoRow, 2)];
  }

  const data = await google.getSheetData(config.sheet.name);
  if (!data || data.length < 2) return [];

  const out = [];
  for (let i = 1; i < data.length; i++) {
    const status = data[i][COL.STATUS - 1];
    if (!INVOICED_STATUSES.includes(status)) continue;
    out.push(reconstruct(data[i], i + 1));
  }

  // Newest first, by status-change date
  out.sort((a, b) => {
    const da = parseDateString(a.tidspunkt);
    const db = parseDateString(b.tidspunkt);
    return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
  });

  return out.slice(0, limit);
}

async function getById(id) {
  if (config.demoMode) {
    const all = await list();
    return all.find(e => e.id === Number(id)) || null;
  }
  const data = await google.getSheetData(config.sheet.name);
  const row = data[Number(id) - 1];
  if (!row) return null;
  return reconstruct(row, Number(id));
}

// ============================================================
// Re-send a faktura as an UPDATE that overwrites the previous one.
// Rebuilds from the current row data and prepends a banner telling regnskap
// this replaces the earlier mail.
// ============================================================
async function resend(id) {
  if (config.demoMode) return { demoMode: true };

  const data = await google.getSheetData(config.sheet.name);
  const row = data[Number(id) - 1];
  if (!row) throw new Error('Fant ikke oppdraget');

  const oppdragsnr = row[COL.OPPDRAGSNR - 1] || '';
  const adresse = extractAddressFromHyperlink(row[COL.ADRESSE - 1]);
  const tidspunkt = row[COL.DATO_STATUSENDRING - 1] || '';

  const { buildFakturaEmail } = require('./emails');
  const datoStr = formatDate(new Date(), 'dd.MM.yyyy HH:mm');

  const banner =
    '<div style="background:#fff3cd; border:2px solid #e0a800; padding:16px; ' +
    'margin-bottom:20px; border-radius:6px; color:#856404;">' +
    '<strong style="font-size:16px; display:block; margin-bottom:6px;">⚠️ OPPDATERT FAKTURA — ERSTATTER TIDLIGERE</strong>' +
    `Denne e-posten <strong>erstatter</strong> fakturaen som ble sendt tidligere${tidspunkt ? ' (' + tidspunkt + ')' : ''}. ` +
    'Bruk opplysningene i <strong>denne</strong> e-posten. Forrige versjon skal forkastes.' +
    '</div>';

  const html = banner + buildFakturaEmail(row, datoStr);
  const subject = `[OPPDATERT] Klar til fakturering: ${adresse} (${oppdragsnr})`;

  await google.sendEmail(config.email.accountantEmail, subject, html);

  return { resent: true, oppdragsnr };
}

module.exports = { list, getById, resend };
