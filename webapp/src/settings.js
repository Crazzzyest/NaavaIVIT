const config = require('./config');
const google = require('./google');

const SETTINGS_SHEET = 'Innstillinger';
const CACHE_TTL_MS = 30000;

// ============================================================
// DEFAULT SETTINGS
// ============================================================
// All keys live as flat strings (e.g. "cron.scanEmails"). Values are strings
// in the sheet, coerced to their true type by `coerce()` when read.
// ============================================================

const DEFAULTS = {
  // Cron schedules (node-cron syntax)
  'cron.scanEmails':                '*/5 * * * *',
  'cron.processIvit':               '*/15 * * * *',
  'cron.checkReminders':            '0 * * * *',
  'cron.updateDashboard':           '30 * * * *',
  'cron.weeklyReport':              '0 16 * * 5',

  // Email enables (which emails the system sends)
  'email.sendBefaringConfirmation': false,  // Apps Script default: OFF (commented out)
  'email.sendNewOppdragNotification': false, // Apps Script default: OFF (commented out)
  'email.sendFakturaToAccountant':  true,
  'email.sendUrgentReminder':       true,

  // Email subjects (single line; placeholders like {{adresse}} are replaced)
  'email.befaringSubject':          '✅ Befaring avtalt — {{adresse}} ({{befaringDato}})',
  'email.nyttOppdragSubject':       '📋 Nytt takstoppdrag: {{adresse}} ({{oppdragstype}})',
  'email.fakturaSubject':           'Klar til fakturering: {{adresse}} ({{oppdragsnr}})',
  'email.ukerapportSubject':        '📊 Naava Takst — Ukerapport uke {{ukenr}}',
  'email.urgentReminderSubject':    '⚠️ Urgent: oppdrag mer enn 24 t i Mottatt — {{adresse}}',

  // Email signature (HTML block appended to all outgoing emails to customers)
  'email.signatureHtml':
    '<strong>Jacob Engholm Holen</strong><br>' +
    'Takstingeniør<br>' +
    '+47 469 49 615<br>' +
    'jacob@naava.no<br>' +
    'www.naava.no<br><br>' +
    '<em>Medlem av Norsk Takst og NITO</em>',

  // Email "intro paragraph" — leading text customers see before the structured data.
  // Body templates with full HTML for befaring/nytt-oppdrag/urgent — long strings
  'email.befaringIntro':
    'Vi bekrefter herved avtalt befaring på følgende eiendom. Se sjekkliste under for hvordan du forbereder deg.',
  'email.nyttOppdragIntro':
    'Et nytt takstoppdrag er registrert i systemet.',
  'email.urgentReminderBody':
    'Det er nå over 24 timer siden oppdraget kom inn og det står fortsatt i status "Mottatt". Vennligst sjekk om det krever oppfølging.',

  // Thresholds
  'threshold.urgentHours':          24,
  'threshold.reminderHours':        2,

  // Travel cost
  'travel.baseAddress':             'Postveien 15, 6018 Ålesund',
  'travel.satsKr':                  10,
  'travel.inkludertKm':             50,

  // MVA
  'mva.rate':                       0.25,

  // Recipients
  'recipients.ownerEmail':          'jacob@naava.no',
  'recipients.accountantEmail':     'regnskap@naava.no',
  'recipients.testSender':          'edsongreistad99@gmail.com',

  // IVIT credentials — sent to scraper-service per request, replaces env vars there
  'ivit.username':                  '',
  'ivit.password':                  '',
};

// Field type metadata for coercion + UI hints
const SCHEMA = {
  // cron
  'cron.scanEmails':                { type: 'cron', label: 'E-post-skanning' },
  'cron.processIvit':               { type: 'cron', label: 'IVIT-henting' },
  'cron.checkReminders':            { type: 'cron', label: 'Påminnelser' },
  'cron.updateDashboard':           { type: 'cron', label: 'Dashboard-oppdatering' },
  'cron.weeklyReport':              { type: 'cron', label: 'Ukerapport' },

  'email.sendBefaringConfirmation': { type: 'boolean', label: 'Send befaringsbekreftelse til kunde' },
  'email.sendNewOppdragNotification': { type: 'boolean', label: 'Send "nytt oppdrag"-varsel til eier' },
  'email.sendFakturaToAccountant':  { type: 'boolean', label: 'Send faktura-varsel til regnskap' },
  'email.sendUrgentReminder':       { type: 'boolean', label: 'Send påminnelse på gamle "Mottatt"-oppdrag' },

  'email.befaringSubject':          { type: 'text', label: 'Befaring — emne' },
  'email.nyttOppdragSubject':       { type: 'text', label: 'Nytt oppdrag — emne' },
  'email.fakturaSubject':           { type: 'text', label: 'Faktura — emne' },
  'email.ukerapportSubject':        { type: 'text', label: 'Ukerapport — emne' },
  'email.urgentReminderSubject':    { type: 'text', label: 'Urgent påminnelse — emne' },

  'email.signatureHtml':            { type: 'textarea', label: 'Signatur (HTML)' },
  'email.befaringIntro':            { type: 'textarea', label: 'Befaring — innledning' },
  'email.nyttOppdragIntro':         { type: 'textarea', label: 'Nytt oppdrag — innledning' },
  'email.urgentReminderBody':       { type: 'textarea', label: 'Urgent påminnelse — tekst' },

  'threshold.urgentHours':          { type: 'number', label: 'Timer før "urgent" påminnelse' },
  'threshold.reminderHours':        { type: 'number', label: 'Timer mellom påminnelser' },

  'travel.baseAddress':             { type: 'text', label: 'Base-adresse (start for reiseberegning)' },
  'travel.satsKr':                  { type: 'number', label: 'Reise-sats (kr/km eks mva)' },
  'travel.inkludertKm':             { type: 'number', label: 'Inkluderte km t/r' },

  'mva.rate':                       { type: 'number', label: 'MVA-rate (0.25 = 25%)' },

  'recipients.ownerEmail':          { type: 'email', label: 'Eier e-post' },
  'recipients.accountantEmail':     { type: 'email', label: 'Regnskap e-post' },
  'recipients.testSender':          { type: 'email', label: 'Test-avsender (kun TEST_MODE)' },

  'ivit.username':                  { type: 'text', label: 'IVIT brukernavn' },
  'ivit.password':                  { type: 'password', label: 'IVIT passord' },
};

// ============================================================
// CACHE
// ============================================================
let cache = { ts: 0, data: null };

function coerce(rawValue, type) {
  if (rawValue === null || rawValue === undefined) return null;
  if (type === 'boolean') {
    return rawValue === true || rawValue === 'true' || rawValue === 'TRUE' || rawValue === 1;
  }
  if (type === 'number') {
    const n = parseFloat(String(rawValue).replace(',', '.'));
    return Number.isFinite(n) ? n : null;
  }
  return String(rawValue);
}

function stringify(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'boolean') return value ? 'true' : 'false';
  return String(value);
}

// ============================================================
// READ
// ============================================================

async function loadFromSheet() {
  // Demo mode: don't touch sheet, return defaults
  if (config.demoMode) {
    return { ...DEFAULTS };
  }

  let rows;
  try {
    rows = await google.getSheetData(SETTINGS_SHEET);
  } catch (e) {
    // Sheet doesn't exist yet — return defaults; first patch will create the sheet
    console.warn(`Settings sheet "${SETTINGS_SHEET}" not found. Using defaults. Error: ${e.message}`);
    return { ...DEFAULTS };
  }

  const merged = { ...DEFAULTS };
  if (rows && rows.length >= 2) {
    for (let i = 1; i < rows.length; i++) {
      const key = String(rows[i][0] || '').trim();
      if (!key || !(key in DEFAULTS)) continue;
      const raw = rows[i][1];
      if (raw === '' || raw === null || raw === undefined) continue;
      const spec = SCHEMA[key];
      const coerced = coerce(raw, spec ? spec.type : 'text');
      if (coerced !== null) merged[key] = coerced;
    }
  }
  return merged;
}

async function get() {
  const now = Date.now();
  if (cache.data && now - cache.ts < CACHE_TTL_MS) return cache.data;
  const data = await loadFromSheet();
  cache = { ts: now, data };
  return data;
}

function getCached() {
  // Synchronous access — returns last loaded or defaults if not yet loaded
  return cache.data || { ...DEFAULTS };
}

function bustCache() {
  cache = { ts: 0, data: null };
}

// ============================================================
// WRITE
// ============================================================

async function ensureSheetExists() {
  // Try to read 1 cell; if it fails, the sheet probably doesn't exist
  try {
    await google.getSheetData(SETTINGS_SHEET);
    return;
  } catch (e) {
    // Need to create — google.js doesn't have createSheet helper yet, so use raw API
    const { google: gApi } = require('googleapis');
    const auth = require('./google');
    // Use Sheets API directly via google.js auth (we don't expose it cleanly yet,
    // so for simplicity require user to create the sheet manually on first run).
    throw new Error(
      `Innstillinger-sheet finnes ikke. Opprett et faneblad kalt "${SETTINGS_SHEET}" i Google Sheet manuelt (kolonne A: nøkkel, kolonne B: verdi).`,
    );
  }
}

async function patch(updates) {
  if (config.demoMode) {
    // Demo mode: don't actually save, just simulate
    return { demoMode: true, updated: updates };
  }

  await ensureSheetExists();

  // Read existing rows to find matching keys (we update if present, append if not)
  const rows = (await google.getSheetData(SETTINGS_SHEET)) || [];
  const keyToRow = {};
  for (let i = 1; i < rows.length; i++) {
    const k = String(rows[i][0] || '').trim();
    if (k) keyToRow[k] = i + 1; // 1-indexed
  }

  // Header row if missing
  if (rows.length === 0) {
    await google.appendRow(SETTINGS_SHEET, ['Nøkkel', 'Verdi']);
  }

  for (const [key, value] of Object.entries(updates || {})) {
    if (!(key in DEFAULTS)) continue; // ignore unknown keys
    const stringVal = stringify(value);
    if (keyToRow[key]) {
      await google.updateCell(SETTINGS_SHEET, keyToRow[key], 2, stringVal);
    } else {
      await google.appendRow(SETTINGS_SHEET, [key, stringVal]);
    }
  }

  bustCache();
  return { updated: updates };
}

// ============================================================
// EXPORT
// ============================================================

module.exports = {
  get,
  getCached,
  patch,
  bustCache,
  DEFAULTS,
  SCHEMA,
  SETTINGS_SHEET,
};
