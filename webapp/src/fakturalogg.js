const config = require('./config');
const google = require('./google');
const { COL } = require('./columns');
const { formatDate } = require('./utils');

const SHEET = 'Fakturalogg';
const HEADER = ['Tidspunkt', 'Timestamp (ISO)', 'Oppdragsnr', 'Adresse', 'Mottaker', 'Emne', 'Versjon', 'HTML'];

// ============================================================
// Append one sent-faktura record to the Fakturalogg sheet.
// Stores full HTML so the admin can review exactly what was sent.
// ============================================================
async function log({ oppdragsnr, adresse, mottaker, emne, html, versjon }) {
  if (config.demoMode) return { demoMode: true };

  const now = new Date();
  const row = [
    formatDate(now, 'dd.MM.yyyy HH:mm'),
    now.toISOString(),
    String(oppdragsnr || ''),
    String(adresse || ''),
    String(mottaker || ''),
    String(emne || ''),
    String(versjon || 1),
    String(html || ''),
  ];

  try {
    // Ensure header exists on first write
    const existing = await google.getSheetData(SHEET).catch(() => []);
    if (!existing || existing.length === 0) {
      await google.appendRow(SHEET, HEADER);
    }
    await google.appendRow(SHEET, row);
  } catch (e) {
    console.error('Fakturalogg.log failed:', e.message);
  }
  return { logged: true };
}

// ============================================================
// List recent faktura records (newest first).
// ============================================================
async function list(limit = 100) {
  if (config.demoMode) {
    return [{
      id: 2,
      tidspunkt: formatDate(new Date(), 'dd.MM.yyyy HH:mm'),
      oppdragsnr: 'DEMO-001',
      adresse: 'Demogata 1, 0001 Demo',
      mottaker: config.email.accountantEmail,
      emne: 'Klar til fakturering: Demogata 1 (DEMO-001)',
      versjon: 1,
      html: '<div style="font-family:Arial">Demo faktura-innhold</div>',
    }];
  }

  let rows;
  try {
    rows = await google.getSheetData(SHEET);
  } catch {
    return [];
  }
  if (!rows || rows.length < 2) return [];

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i] || [];
    out.push({
      id: i + 1, // sheet row number
      tidspunkt: r[0] || '',
      iso: r[1] || '',
      oppdragsnr: r[2] || '',
      adresse: r[3] || '',
      mottaker: r[4] || '',
      emne: r[5] || '',
      versjon: Number(r[6]) || 1,
      html: r[7] || '',
    });
  }
  // Newest first
  out.reverse();
  return out.slice(0, limit);
}

async function getById(id) {
  const all = await list(1000);
  return all.find(e => e.id === Number(id)) || null;
}

// ============================================================
// Re-send a faktura as an UPDATE that overwrites the previous one.
// Rebuilds from the CURRENT row data (so any fixes are reflected) and
// prepends a banner telling regnskap this replaces the earlier mail.
// ============================================================
async function resend(id) {
  if (config.demoMode) return { demoMode: true };

  const entry = await getById(id);
  if (!entry) throw new Error('Fakturalogg-oppføring ikke funnet');

  // Find the current row for this oppdrag (Oppdragslogg first, then arkiv)
  let rowData = await findRowByOppdragsnr(config.sheet.name, entry.oppdragsnr);
  if (!rowData) rowData = await findRowByOppdragsnr('arkiv', entry.oppdragsnr);
  if (!rowData) throw new Error(`Fant ikke oppdrag ${entry.oppdragsnr} i Oppdragslogg eller arkiv`);

  const { buildFakturaEmail } = require('./emails');
  const datoStr = formatDate(new Date(), 'dd.MM.yyyy HH:mm');

  const banner =
    '<div style="background:#fff3cd; border:2px solid #e0a800; padding:16px; ' +
    'margin-bottom:20px; border-radius:6px; color:#856404;">' +
    '<strong style="font-size:16px; display:block; margin-bottom:6px;">⚠️ OPPDATERT FAKTURA — ERSTATTER TIDLIGERE</strong>' +
    `Denne e-posten <strong>erstatter</strong> fakturaen som ble sendt ${entry.tidspunkt}. ` +
    'Bruk opplysningene i <strong>denne</strong> e-posten. Forrige versjon skal forkastes.' +
    '</div>';

  const html = banner + buildFakturaEmail(rowData, datoStr);
  const subject = '[OPPDATERT] ' + entry.emne.replace(/^\[OPPDATERT\]\s*/, '');

  await google.sendEmail(config.email.accountantEmail, subject, html);

  const nyVersjon = (entry.versjon || 1) + 1;
  await log({
    oppdragsnr: entry.oppdragsnr,
    adresse: entry.adresse,
    mottaker: config.email.accountantEmail,
    emne: subject,
    html,
    versjon: nyVersjon,
  });

  return { resent: true, versjon: nyVersjon };
}

async function findRowByOppdragsnr(sheetName, oppdragsnr) {
  try {
    const data = await google.getSheetData(sheetName);
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][COL.OPPDRAGSNR - 1] || '') === String(oppdragsnr)) {
        return data[i];
      }
    }
  } catch { /* sheet may not exist */ }
  return null;
}

module.exports = { log, list, getById, resend, SHEET };
