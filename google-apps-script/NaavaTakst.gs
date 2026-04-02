// ============================================================
// NAAVA TAKST — OPPDRAGSHÅNDTERING v4.0 (KOMBINERT MED IVIT)
// ============================================================
// Endringer fra v3.1:
//   - IVIT webhook-scraping integrert (var IvitWebhook.gs)
//   - Ny kolonne AJ (36) = "Scan IVIT" checkbox
//   - Prisberegning: +1250 inkl per tilleggsbygg, +2000 inkl for markedsverdi
//   - Reisekostnad: inkluderer ferge/bom, deling mellom personer
//   - onEdit trigger for Q, AA, AB
// ============================================================

// ---- KONFIGURASJON ----
const CONFIG = {
  // ======= TESTMODUS =======
  TEST_MODE: false,
  // ==========================
  SHEET_ID: '1VEzaCNEkvbWYZf0UOrj6IG5PUQB3hL1enB24PvsQ1MI',
  OWNER_EMAIL: 'jacob@naava.no',
  ACCOUNTANT_EMAIL: 'regnskap@naava.no',
  ROOT_FOLDER_NAME: 'Naava',
  ROOT_FOLDER_ID: '1nDgJrHWnEdkkG1OR90vGmHby4CTYA7J_',
  SHEET_NAME: 'Oppdragslogg',
  DASHBOARD_SHEET_NAME: 'Dashboard',
  CHAT_WEBHOOK_URL: 'SETT_INN_WEBHOOK_URL_HER',
  BASE_ADDRESS: 'Postveien 15, 6018 Ålesund',
  REISE_SATS_EKS_MVA: 10,
  REISE_INKLUDERT_KM: 50,
  OPENAI_API_KEY: '', // Set in Apps Script project properties
  OPENAI_MODEL: 'gpt-5.2',
  TEST_SENDER: 'edsongreistad99@gmail.com',
  TRIGGER_KEYWORDS: [
    'tilstandsrapport', 'takst', 'befaring', 'verdivurdering',
    'markedsverdi', 'boligsalgsrapport', 'ordre i ivit',
    'skaderapport', 'reklamasjon', 'skadetakst', 'vurderingsoppdrag',
    'overtagelse', 'bistand'
  ],
  REMINDER_HOURS: 2,
  URGENT_HOURS: 24,
  MVA_RATE: 0.25,

  // ======= IVIT WEBHOOK =======
  IVIT_WEBHOOK_URL: 'https://naavaivit.sliplane.app/webhook',
  IVIT_WEBHOOK_SECRET: 'b9a863c2f947e4d54b6feda001cb15c0ad5ec49ce2450e6f591c1528d85b85b8',
  IVIT_CUTOFF_DATE: new Date('2026-03-17T08:23:00'),
};

// ============================================================
// KOLONNEINDEKSER (1-basert) — 36 kolonner
// ============================================================
const COL = {
  OPPDRAGSNR:          1,   // A
  DATO_MOTTATT:        2,   // B
  KILDE:               3,   // C
  OPPDRAGSTYPE:        4,   // D
  ADRESSE:             5,   // E
  OPPDRAGSGIVER:       6,   // F
  SELGER:              7,   // G
  SELGER_TLF:          8,   // H
  SELGER_EPOST:        9,   // I
  MEGLER:              10,  // J
  MEGLER_EPOST:        11,  // K
  FAKTURA_REF:         12,  // L
  FAKTURA_SENDES_TIL:  13,  // M
  FAKTURAMOTAKER:      14,  // N
  BOLIGTYPE:           15,  // O
  AREAL:               16,  // P
  ANTALL_TILLEGGSBYGG: 17,  // Q
  RAPPORTTYPE:         18,  // R
  MED_MARKEDSVERDI:    19,  // S
  TIMER:               20,  // T
  PRIS_INKL:           21,  // U
  PRIS_EKS:            22,  // V
  MVA_BELOP:           23,  // W
  AVSTAND_KM:          24,  // X
  REISE_EKS:           25,  // Y
  REISE_INKL:          26,  // Z
  SUM_FERGE_BOM:       27,  // AA
  ANTALL_DELE_REISE:   28,  // AB
  STATUS:              29,  // AC
  BEFARING_DATO:       30,  // AD
  BEFARING_KL:         31,  // AE
  DATO_STATUSENDRING:  32,  // AF
  TIMESTAMP:           33,  // AG
  LINK_MAPPE:          34,  // AH
  NOTATER:             35,  // AI
  SCAN_IVIT:           36,  // AJ  ← NY
};
const NUM_COLS = 36;
const STATUS_COL_LETTER = 'AC';

// ---- PRISLISTER (inkl mva) ----
const PRISLISTE = {
  'Leilighet': [
    { maxAreal: 80, pris: 10000 },
    { maxAreal: Infinity, pris: 12000 }
  ],
  'Rekkehus/leilighet 2-4-mannsbolig': [
    { maxAreal: 80, pris: 12000 },
    { maxAreal: Infinity, pris: 14000 }
  ],
  'Enebolig/fritidsbolig': [
    { maxAreal: 150, pris: 16000 },
    { maxAreal: 250, pris: 18000 },
    { maxAreal: Infinity, pris: 20000 }
  ],
  'Frittstående bygg': [
    { maxAreal: Infinity, pris: 1250 }
  ]
};

// Tillegg (inkl mva)
const TILLEGG_MARKEDSVERDI_INKL = 2000;
const TILLEGG_PER_TILLEGGSBYGG_INKL = 1250;

// ============================================================
// FELLES HJELPEFUNKSJONER
// ============================================================
function getSpreadsheet_() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  return SpreadsheetApp.openById(CONFIG.SHEET_ID);
}

function safeSendEmail_(to, subject, htmlBody, plainBody) {
  let safeRecipient = to;
  if (CONFIG.TEST_MODE) {
    const allowedEmails = [CONFIG.OWNER_EMAIL.toLowerCase(), CONFIG.ACCOUNTANT_EMAIL.toLowerCase()];
    const recipients = safeRecipient.split(',').map(function (e) { return e.trim().toLowerCase(); });
    const safeRecipients = recipients.filter(function (e) { return allowedEmails.indexOf(e) > -1; });
    if (safeRecipients.length === 0) {
      safeRecipient = CONFIG.OWNER_EMAIL;
      subject = '🧪 [TEST - Ville gått til: ' + to + '] ' + subject;
    } else {
      safeRecipient = safeRecipients.join(',');
    }
  }
  const emailParams = { to: safeRecipient, subject: subject };
  if (htmlBody) emailParams.htmlBody = htmlBody;
  if (plainBody) emailParams.body = plainBody;
  MailApp.sendEmail(emailParams);
  Logger.log('📧 E-post sendt til: ' + safeRecipient + ' | Emne: ' + subject);
}

function formatCurrency_(amount) {
  if (!amount || isNaN(amount)) return '0 kr';
  return Number(amount).toLocaleString('nb-NO') + ' kr';
}

function getWeekNumber_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
}

function parseDateString_(dateStr) {
  if (dateStr instanceof Date) return dateStr;
  if (typeof dateStr !== 'string' || !dateStr) return null;
  const parts = dateStr.split(' ')[0].split('.');
  if (parts.length >= 3) return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
  return null;
}

function setDropdown_(sheet, col, values) {
  sheet.getRange(2, col, 500).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build()
  );
}

function getOrCreateRootFolder_() {
  if (CONFIG.ROOT_FOLDER_ID) return DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  const folders = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME);
}

function createOppdragFolder_(folderName) {
  const root = getOrCreateRootFolder_();
  const folder = root.createFolder(folderName);
  folder.createFolder('Bilder');
  folder.createFolder('Dokumenter fra megler');
  folder.createFolder('Rapport');
  return folder;
}

function getOrCreateLabel_(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) label = GmailApp.createLabel(labelName);
  return label;
}

function sendChatAlert_(message) {
  if (CONFIG.CHAT_WEBHOOK_URL === 'SETT_INN_WEBHOOK_URL_HER') {
    safeSendEmail_(CONFIG.OWNER_EMAIL, '🔔 Naava Takst Alert', null, message.replace(/\*/g, '').replace(/_/g, ''));
    return;
  }
  try {
    UrlFetchApp.fetch(CONFIG.CHAT_WEBHOOK_URL, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ text: message })
    });
  } catch (e) {
    safeSendEmail_(CONFIG.OWNER_EMAIL, '🔔 Naava Takst Alert', null, message.replace(/\*/g, ''));
  }
}

function extractDriveFolderId_(value) {
  const s = String(value || '').trim();
  if (!s) return '';
  const m1 = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m1) return m1[1];
  const m2 = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2) return m2[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(s) && s.indexOf('http') !== 0) return s;
  return '';
}

function getOrCreateAvsluttedeOppdragFolder_() {
  const root = CONFIG.ROOT_FOLDER_ID ? DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID) : getOrCreateRootFolder_();
  const targetName = 'avsluttede oppdrag';
  const it = root.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    if (String(f.getName() || '').toLowerCase() === targetName) return f;
  }
  return root.createFolder('avsluttede oppdrag');
}

function moveOppdragFolderToAvsluttede_(folderId) {
  if (!folderId) throw new Error('Mangler folderId');
  const folder = DriveApp.getFolderById(folderId);
  const targetParent = getOrCreateAvsluttedeOppdragFolder_();
  targetParent.addFolder(folder);
  const parents = folder.getParents();
  while (parents.hasNext()) {
    const p = parents.next();
    if (p.getId() !== targetParent.getId()) p.removeFolder(folder);
  }
  return targetParent.getId();
}

// ============================================================
// 1. INITIAL SETUP
// ============================================================
function initialSetup() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let isStandalone = false;
  if (!ss) {
    isStandalone = true;
    if (CONFIG.SHEET_ID && CONFIG.SHEET_ID !== '') {
      try { ss = SpreadsheetApp.openById(CONFIG.SHEET_ID); } catch (e) { }
    }
    if (!ss) {
      ss = SpreadsheetApp.create('Naava Takst Oppdragslogg');
      Logger.log('⚠️ Oppdater CONFIG.SHEET_ID til: ' + ss.getId());
    }
  }

  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  else { sheet.clear(); sheet.clearConditionalFormatRules(); }

  ['Ark 1', 'Sheet1'].forEach(function (name) {
    const s = ss.getSheetByName(name);
    if (s && s.getLastRow() === 0) { try { ss.deleteSheet(s); } catch (e) { } }
  });

  const headers = [
    'Oppdragsnr', 'Dato mottatt', 'Kilde', 'Oppdragstype', 'Adresse',
    'Oppdragsgiver', 'Selger', 'Selger tlf', 'Selger e-post',
    'Megler/bestiller', 'Megler e-post', 'Faktura ref', 'Faktura sendes til',
    'Fakturamotaker', 'Boligtype', 'Areal (m²)', 'Antall tilleggsbygg',
    'Rapporttype', 'Med markedsverdi', 'Timer',
    'Pris inkl. mva', 'Pris eks. mva', 'MVA-beløp',
    'Avstand (km t/r)', 'Reisekostnad eks mva', 'Reisekostnad inkl mva',
    'Sum (ferge/bom)', 'Antall som deler reise',
    'Status', 'Befaring dato', 'Befaring klokkeslett',
    'Dato statusendring', 'Timestamp (intern)', 'Link til mappe', 'Notater',
    'Scan IVIT'  // AJ - NY
  ];

  sheet.getRange(1, 1, 1, NUM_COLS).setValues([headers]);
  sheet.getRange(1, 1, 1, NUM_COLS)
    .setFontWeight('bold').setBackground('#1a5c2a')
    .setFontColor('#ffffff').setFontSize(10).setWrap(true);
  sheet.setFrozenRows(1);

  const widths = {};
  widths[COL.OPPDRAGSNR] = 110; widths[COL.DATO_MOTTATT] = 130;
  widths[COL.KILDE] = 110; widths[COL.OPPDRAGSTYPE] = 180;
  widths[COL.ADRESSE] = 250; widths[COL.OPPDRAGSGIVER] = 180;
  widths[COL.SELGER] = 160; widths[COL.SELGER_TLF] = 120;
  widths[COL.SELGER_EPOST] = 180; widths[COL.MEGLER] = 180;
  widths[COL.MEGLER_EPOST] = 200; widths[COL.FAKTURA_REF] = 120;
  widths[COL.FAKTURA_SENDES_TIL] = 180; widths[COL.FAKTURAMOTAKER] = 180;
  widths[COL.BOLIGTYPE] = 260; widths[COL.AREAL] = 90;
  widths[COL.ANTALL_TILLEGGSBYGG] = 100; widths[COL.RAPPORTTYPE] = 300;
  widths[COL.MED_MARKEDSVERDI] = 120; widths[COL.TIMER] = 90;
  widths[COL.PRIS_INKL] = 120; widths[COL.PRIS_EKS] = 120;
  widths[COL.MVA_BELOP] = 100; widths[COL.AVSTAND_KM] = 110;
  widths[COL.REISE_EKS] = 140; widths[COL.REISE_INKL] = 140;
  widths[COL.SUM_FERGE_BOM] = 140; widths[COL.ANTALL_DELE_REISE] = 100;
  widths[COL.STATUS] = 160; widths[COL.BEFARING_DATO] = 130;
  widths[COL.BEFARING_KL] = 120; widths[COL.DATO_STATUSENDRING] = 150;
  widths[COL.TIMESTAMP] = 180; widths[COL.LINK_MAPPE] = 300;
  widths[COL.NOTATER] = 250; widths[COL.SCAN_IVIT] = 80;
  Object.entries(widths).forEach(function (e) { sheet.setColumnWidth(parseInt(e[0]), e[1]); });

  sheet.hideColumns(COL.TIMESTAMP);

  // Dropdowns
  setDropdown_(sheet, COL.KILDE, ['Megler-epost', 'IVIT', 'Manuell']);
  setDropdown_(sheet, COL.OPPDRAGSTYPE, [
    'Tilstandsrapport', 'Tilstandsrapport m/markedsverdi', 'Skadetakst',
    'Reklamasjon', 'Vurderingsoppdrag', 'Bistand overtagelse',
    'Fukt-/fuktskadevurdering', 'Byggelånskontroll', 'Forhåndstakst', 'Verditakst', 'Annet'
  ]);
  setDropdown_(sheet, COL.BOLIGTYPE, [
    'Leilighet', 'Rekkehus/leilighet 2-4-mannsbolig', 'Enebolig/fritidsbolig',
    'Frittstående bygg', 'Næringsbygg', 'Annet'
  ]);
  setDropdown_(sheet, COL.RAPPORTTYPE, [
    'Tilstandsrapport m/teknisk og markedsverdi', 'Tilstandsrapport',
    'Skadetakstrapport', 'Reklamasjonsrapport', 'Vurderingsrapport',
    'Overtagelsesrapport', 'Annen rapport'
  ]);
  setDropdown_(sheet, COL.STATUS, [
    'Mottatt', 'Avtalt befaring', 'Befart', 'Utkast',
    'Endelig rapport', 'Kan faktureres', 'Fakturert', 'Oppdrag kansellert', 'Oppdrag fullført'
  ]);

  // Tallformat
  [COL.PRIS_INKL, COL.PRIS_EKS, COL.MVA_BELOP, COL.REISE_EKS, COL.REISE_INKL, COL.SUM_FERGE_BOM].forEach(function (c) {
    sheet.getRange(2, c, 500).setNumberFormat('#,##0 "kr"');
  });
  sheet.getRange(2, COL.AVSTAND_KM, 500).setNumberFormat('#,##0 "km"');
  sheet.getRange(2, COL.BEFARING_DATO, 500).setNumberFormat('dd.MM.yyyy');

  // Checkboxer
  sheet.getRange(2, COL.MED_MARKEDSVERDI, 500).insertCheckboxes();
  sheet.getRange(2, COL.SCAN_IVIT, 500).insertCheckboxes();

  setupConditionalFormatting_(sheet, NUM_COLS);
  setupDashboard_(ss);
  getOrCreateRootFolder_();
  setupTriggers_(ss);

  const modeLabel = CONFIG.TEST_MODE ? 'TESTMODUS' : 'PRODUKSJON';
  const confirmMsg = 'Naava Takst v4.0 — ' + modeLabel + '\n' +
    NUM_COLS + ' kolonner | Sheet ID: ' + ss.getId() + '\n' + ss.getUrl();
  if (!isStandalone) {
    try { SpreadsheetApp.getUi().alert('✅ ' + confirmMsg); } catch (e) { Logger.log('✅ ' + confirmMsg); }
  } else {
    Logger.log('✅ ' + confirmMsg);
  }
}

// ============================================================
// 2. E-POST SCANNING
// ============================================================
function scanIncomingEmails() {
  const label = getOrCreateLabel_('Takst-Behandlet');
  let query;
  if (CONFIG.TEST_MODE) {
    const keywords = CONFIG.TRIGGER_KEYWORDS.map(kw => `"${kw}"`).join(' OR ');
    query = 'from:' + CONFIG.TEST_SENDER + ' (' + keywords + ') -label:Takst-Behandlet after:2026/02/05';
  } else {
    const keywords = CONFIG.TRIGGER_KEYWORDS.map(function (k) { return '"' + k + '"'; }).join(' OR ');
    const meglerQuery = '(' + keywords + ')';
    const ivitQuery = '(from:no-reply.takst@ivit.no)';
    query = '(' + meglerQuery + ' OR ' + ivitQuery + ') -label:Takst-Behandlet newer_than:2d';
  }
  let threads;
  try { threads = GmailApp.search(query, 0, 20); } catch (e) { Logger.log('Gmail-søk feilet: ' + e.message); return; }
  Logger.log('Fant ' + threads.length + ' tråder');

  for (var ti = 0; ti < threads.length; ti++) {
    var thread = threads[ti];
    try {
      var messages = thread.getMessages();
      var msg = messages[messages.length - 1];
      var subject = msg.getSubject();
      var body = msg.getPlainBody();
      var from = msg.getFrom();
      var date = msg.getDate();

      if (CONFIG.TEST_MODE && from.toLowerCase().indexOf(CONFIG.TEST_SENDER.toLowerCase()) === -1) continue;

      var isIVIT = (from || '').toLowerCase().indexOf('no-reply.takst@ivit.no') > -1;
      var parsed = null;
      if (isIVIT) {
        parsed = parseIVITEmail_(subject, body, from);
      } else {
        parsed = assessAndParseWithAI_(subject, body, from);
        if (!parsed) { thread.addLabel(label); thread.markRead(); continue; }
      }
      if (CONFIG.TEST_MODE) parsed = sanitizeParsedData_(parsed);

      createNewOppdrag_(parsed, date);
      thread.addLabel(label);
      thread.markRead();
    } catch (loopErr) {
      Logger.log('FEIL i tråd ' + ti + ': ' + loopErr.message);
    }
  }
}

function sanitizeParsedData_(parsed) {
  const allowedEmails = [CONFIG.OWNER_EMAIL.toLowerCase(), CONFIG.ACCOUNTANT_EMAIL.toLowerCase()];
  if (parsed.selgerEpost && allowedEmails.indexOf(parsed.selgerEpost.toLowerCase()) === -1) {
    parsed.selgerEpost = CONFIG.ACCOUNTANT_EMAIL;
  }
  if (parsed.meglerEpost && allowedEmails.indexOf(parsed.meglerEpost.toLowerCase()) === -1) {
    parsed.meglerEpost = CONFIG.ACCOUNTANT_EMAIL;
  }
  return parsed;
}

// ============================================================
// 3. PARSERE
// ============================================================
function parseIVITEmail_(subject, body, from) {
  const text = (body || '').replace(/\r/g, '');
  const result = {
    kilde: 'IVIT', oppdragstype: 'Tilstandsrapport',
    adresse: '', oppdragsgiver: '', selger: '', selgerTlf: '', selgerEpost: '',
    megler: '', meglerEpost: '', fakturaRef: '', fakturaSendesTil: '', notater: ''
  };
  const ordreMatch = text.match(/(?:^|\n)\s*[Oo]rdre\s*nummer\s*[:\t ]?\s*([A-Z0-9-]+)\s*(?:\n|$)/);
  if (ordreMatch) result.fakturaRef = ordreMatch[1].trim();

  let adresseMatch = text.match(/følgende eiendom[:\s]*\n\s*(.+?)(?:\n|,\s*gnr)/i);
  if (adresseMatch) result.adresse = adresseMatch[1].trim();
  if (!result.adresse) {
    adresseMatch = text.match(/[Aa]dress(?:e[n]?\s+er|e[:\s])\s*(.+?)(?:\n|$)/i);
    if (adresseMatch) result.adresse = adresseMatch[1].trim();
  }
  if (!result.adresse) {
    adresseMatch = text.match(/([A-ZÆØÅ][a-zæøåA-ZÆØÅ]*(?:ringen|veien|gata|gaten|vegen|stien|bakken|lia|haugen|åsen|berget|stranda|plassen|torget|brygga|bøen|øen|tunet|marka|jordet|løkka)\s+\d+[A-Za-z]?(?:\s*,?\s*\d{4}\s+[A-ZÆØÅa-zæøå]+)?)/);
    if (adresseMatch) result.adresse = adresseMatch[1].trim();
  }
  const hilsenBlock = text.match(/[Vv]ennlig hilsen\s*\n([\s\S]*?)(?:\n\s*\n|$)/);
  if (hilsenBlock) {
    const lines = hilsenBlock[1].split('\n').map(function (l) { return l.trim(); }).filter(Boolean);
    if (lines.length >= 1) result.megler = lines[0];
    if (lines.length >= 2) result.fakturaSendesTil = lines[1];
  }
  const emailMatch = from.match(/<(.+?)>/);
  result.meglerEpost = emailMatch ? emailMatch[1] : from.trim();
  return result;
}

function assessAndParseWithAI_(subject, body, from) {
  if (!CONFIG.OPENAI_API_KEY) {
    var kw = CONFIG.TRIGGER_KEYWORDS.some(function (k) {
      return (subject + ' ' + body).toLowerCase().indexOf(k) > -1;
    });
    return kw ? parseMeglerEmail_(subject, body, from) : null;
  }
  var emailFrom = from || '';
  var emailMatch = emailFrom.match(/<(.+?)>/);
  var senderEmail = emailMatch ? emailMatch[1] : emailFrom.trim();
  var senderName = (emailFrom.match(/^"?([^"<]+)"?\s*</) || [])[1] || senderEmail;

  var systemPrompt =
    'Du er et system som vurderer innkommende e-poster for et takstfirma i Norge (Naava Takst). ' +
    'Du skal avgjøre om e-posten er en NY bestilling eller forespørsel om et takstoppdrag. ' +
    '\n\nSett "relevant": false for ALLE disse tilfellene:' +
    '\n- Oppfølging, statusoppdatering eller melding om et eksisterende oppdrag' +
    '\n- Oversending, vedlegg eller videresending av en ferdig eller tidligere rapport' +
    '\n- Avlysning, utsettelse eller omplanlegging av befaring' +
    '\n- Spørsmål om et pågående oppdrag' +
    '\n- Videresending av dokumenter til eksisterende oppdrag' +
    '\n- Spam, nyhetsbrev, fakturaer, kvitteringer, automatiske varsler' +
    '\n- Intern kommunikasjon eller svar på e-post fra Naava Takst selv' +
    '\n\nSett "relevant": true KUN hvis e-posten er en tydelig ny forespørsel eller bestilling. ' +
    'Tvilstilfeller = IKKE relevante. Svar KUN med gyldig JSON.';

  var userPrompt =
    'E-post mottatt av Naava Takst:\n\nFra: ' + from + '\nEmne: ' + subject +
    '\n\nInnhold:\n' + (body || '').substring(0, 3000) +
    '\n\nSvar med JSON:\n{\n  "relevant": true/false,\n  "begrunnelse": "...",\n' +
    '  "oppdragstype": "...", "rapporttype": "...", "adresse": "...",\n' +
    '  "oppdragsgiver": "...", "selger": "...", "selgerTlf": "...", "selgerEpost": "...",\n' +
    '  "megler": "...", "meglerEpost": "...", "fakturaRef": "...",\n' +
    '  "fakturaSendesTil": "...", "notater": "..."\n}';

  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + CONFIG.OPENAI_API_KEY },
      payload: JSON.stringify({
        model: CONFIG.OPENAI_MODEL, temperature: 0,
        messages: [{ role: 'system', content: systemPrompt }, { role: 'user', content: userPrompt }]
      }),
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) return parseMeglerEmail_(subject, body, from);
    var data = JSON.parse(response.getContentText());
    var text = data.choices[0].message.content.trim();
    var jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) return parseMeglerEmail_(subject, body, from);
    var result = JSON.parse(jsonMatch[0]);
    if (!result.relevant) return null;
    return {
      kilde: 'Megler-epost', oppdragstype: result.oppdragstype || 'Annet',
      rapporttype: result.rapporttype || 'Annen rapport',
      adresse: String(result.adresse || '').trim(),
      oppdragsgiver: String(result.oppdragsgiver || '').trim(),
      selger: String(result.selger || '').trim(),
      selgerTlf: String(result.selgerTlf || '').replace(/[^\d+]/g, ''),
      selgerEpost: String(result.selgerEpost || '').trim(),
      megler: String(result.megler || senderName).trim(),
      meglerEpost: String(result.meglerEpost || senderEmail).trim(),
      fakturaRef: String(result.fakturaRef || '').trim(),
      fakturaSendesTil: String(result.fakturaSendesTil || '').trim(),
      notater: String(result.notater || '').trim()
    };
  } catch (e) {
    return parseMeglerEmail_(subject, body, from);
  }
}

function parseMeglerEmail_(subject, body, from) {
  const text = (body || '').replace(/\r/g, '');
  const result = {
    kilde: 'Megler-epost', oppdragstype: 'Tilstandsrapport', rapporttype: '',
    adresse: '', oppdragsgiver: '', selger: '', selgerTlf: '', selgerEpost: '',
    megler: '', meglerEpost: '', fakturaRef: '', fakturaSendesTil: '', notater: ''
  };
  let adresseMatch = text.match(/følgende eiendom[:\s]*\n\s*(.+?)(?:\n|,\s*gnr)/i);
  if (adresseMatch) result.adresse = adresseMatch[1].trim();
  if (!result.adresse) {
    adresseMatch = text.match(/[Aa]dress(?:e[n]?\s+er|e[:\s])\s*(.+?)(?:\n|$)/i);
    if (adresseMatch) result.adresse = adresseMatch[1].trim();
  }
  if (!result.adresse) {
    adresseMatch = text.match(/([A-ZÆØÅ][a-zæøåA-ZÆØÅ]*(?:ringen|veien|gata|gaten|vegen|stien|bakken|lia|haugen|åsen|berget|stranda|plassen|torget|brygga|bøen|øen|tunet|marka|jordet|løkka)\s+\d+[A-Za-z]?(?:\s*,?\s*\d{4}\s+[A-ZÆØÅa-zæøå]+)?)/);
    if (adresseMatch) result.adresse = adresseMatch[1].trim();
  }
  const oppdrMatch = text.match(/oppdragsgiver\s+(.+?)(?:,|\n)/i);
  if (oppdrMatch) { result.oppdragsgiver = oppdrMatch[1].trim(); result.selger = result.oppdragsgiver; }
  const selgerBlock = text.match(/[Ss]elger treffes på[:\s]*\n([\s\S]*?)(?:\n\s*\n|Vi ber)/);
  if (selgerBlock) {
    const block = selgerBlock[1];
    const tlfMatch = block.match(/(?:Tlf|Telefon|Mob)[.:\s]*(\d[\d\s]+\d)/i);
    if (tlfMatch) result.selgerTlf = tlfMatch[1].replace(/\s/g, '');
    const epostMatch = block.match(/[Ee]-?post[:\s]*([^\s\n]+@[^\s\n]+)/);
    if (epostMatch) result.selgerEpost = epostMatch[1].trim();
  }
  if (!result.selgerTlf) {
    const tlfFallback = text.match(/(?:tlf|telefon|mob(?:il)?)[.:\s]*(\d[\d\s]+\d)/i);
    if (tlfFallback) result.selgerTlf = tlfFallback[1].replace(/\s/g, '');
  }
  if (!result.selgerEpost) {
    const epostFallback = text.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    if (epostFallback) result.selgerEpost = epostFallback[1].trim();
  }
  if (!result.oppdragsgiver) {
    const navnMatch = text.match(/(?:mvh|med vennlig hilsen|hilsen|vennlig hilsen)\s*\n?\s*([A-ZÆØÅ][a-zæøå]+(?:\s+[A-ZÆØÅ][a-zæøå]+)*)/i);
    if (navnMatch) { result.oppdragsgiver = navnMatch[1].trim(); if (!result.selger) result.selger = result.oppdragsgiver; }
  }
  const refMatch = text.match(/(?:ref\.?\s*(?:nr\.?)?|referanse(?:nummer)?)[:\s]*(\d+)/i);
  if (refMatch) result.fakturaRef = refMatch[1].trim();
  const fakturaMatch = text.match(/faktura\s+sendes\s+(.+?)(?:\.|$)/im);
  if (fakturaMatch) result.fakturaSendesTil = fakturaMatch[1].trim();
  const emailMatch = from.match(/<(.+?)>/);
  result.meglerEpost = emailMatch ? emailMatch[1] : from.trim();
  const nameMatch = from.match(/^"?([^"<]+)"?\s*</);
  result.megler = nameMatch ? nameMatch[1].trim() : result.meglerEpost;
  return result;
}

// ============================================================
// 4. REISEBEREGNING
// ============================================================
function calculateDistance_(destinationAddress) {
  if (!CONFIG.OPENAI_API_KEY) return null;
  const prompt =
    'Estimert kjøreavstand (én vei) mellom:\nFra: ' + CONFIG.BASE_ADDRESS +
    '\nTil: ' + destinationAddress +
    '\n\nSvar KUN med JSON: {"km_en_vei": <tall>, "km_tur_retur": <tall>, "estimert_tid_min": <tall>}';
  try {
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + CONFIG.OPENAI_API_KEY },
      payload: JSON.stringify({
        model: CONFIG.OPENAI_MODEL, temperature: 0,
        messages: [{ role: 'system', content: 'Du svarer kun med gyldig JSON.' }, { role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) return null;
    const data = JSON.parse(response.getContentText());
    const text = data.choices[0].message.content.trim();
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) return null;
    const result = JSON.parse(jsonMatch[0]);
    return { kmEnVei: Number(result.km_en_vei) || 0, kmTurRetur: Number(result.km_tur_retur) || 0, estimertTidMin: Number(result.estimert_tid_min) || 0 };
  } catch (e) { return null; }
}

function calculateTravelCost_(kmTurRetur) {
  const fakturerbarKm = Math.max(0, kmTurRetur - CONFIG.REISE_INKLUDERT_KM);
  const kostnadEksMva = fakturerbarKm * CONFIG.REISE_SATS_EKS_MVA;
  return {
    totalKm: kmTurRetur, fakturerbarKm: fakturerbarKm,
    kostnadEksMva: Math.round(kostnadEksMva),
    kostnadInklMva: Math.round(kostnadEksMva * (1 + CONFIG.MVA_RATE))
  };
}

/**
 * Beregn total reisekostnad for en rad inkl. ferge/bom og deling.
 */
function recalculateReiseForRow_(sheet, row) {
  const km = parseFloat(String(sheet.getRange(row, COL.AVSTAND_KM).getValue()).replace(',', '.')) || 0;
  const fergeBomInkl = parseFloat(String(sheet.getRange(row, COL.SUM_FERGE_BOM).getValue()).replace(/[^\d.,]/g, '').replace(',', '.')) || 0;
  const antallDeler = Math.max(1, parseInt(sheet.getRange(row, COL.ANTALL_DELE_REISE).getValue()) || 1);

  const tc = calculateTravelCost_(km);

  // Ferge/bom legges til som inkl mva
  const totalInkl = tc.kostnadInklMva + fergeBomInkl;
  const totalEks = Math.round(totalInkl / (1 + CONFIG.MVA_RATE));

  // Del på antall personer
  const reiseInkl = Math.round(totalInkl / antallDeler);
  const reiseEks = Math.round(reiseInkl / (1 + CONFIG.MVA_RATE));

  sheet.getRange(row, COL.REISE_EKS).setValue(reiseEks);
  sheet.getRange(row, COL.REISE_INKL).setValue(reiseInkl);
}

// ============================================================
// 5. PRISBEREGNING (oppdatert med tilleggsbygg + markedsverdi)
// ============================================================
/**
 * Basispris fra PRISLISTE + tilleggsbygg (1250 inkl/stk) + markedsverdi (2000 inkl).
 * Frittstående bygg bruker kun stykk-pris × antall.
 */
function calculatePrice_(sheet, row) {
  const boligtype = sheet.getRange(row, COL.BOLIGTYPE).getValue();
  const arealStr = sheet.getRange(row, COL.AREAL).getValue();
  const rapporttype = sheet.getRange(row, COL.RAPPORTTYPE).getValue();
  if (!boligtype || !rapporttype) return;

  const areal = arealStr ? parseFloat(arealStr) : 0;
  const inkluderMarked = sheet.getRange(row, COL.MED_MARKEDSVERDI).getValue() === true;
  const antallTillegg = parseInt(sheet.getRange(row, COL.ANTALL_TILLEGGSBYGG).getValue()) || 0;

  if (!PRISLISTE[boligtype]) return;
  const priser = PRISLISTE[boligtype];

  // Basispris (inkl mva)
  let basisInkl = 0;
  if (boligtype === 'Frittstående bygg') {
    // Frittstående bygg: stykk-pris, brukes kun som tillegg
    basisInkl = 0;
  } else {
    for (let t = 0; t < priser.length; t++) {
      if (areal <= priser[t].maxAreal || priser[t].maxAreal === Infinity) {
        basisInkl = priser[t].pris;
        break;
      }
    }
  }

  // Tillegg: markedsverdi (+2000 inkl)
  const markedTillegg = inkluderMarked ? TILLEGG_MARKEDSVERDI_INKL : 0;

  // Tillegg: tilleggsbygg (+1250 inkl per stk)
  const tilleggsbyggTillegg = antallTillegg * TILLEGG_PER_TILLEGGSBYGG_INKL;

  const totalInkl = basisInkl + markedTillegg + tilleggsbyggTillegg;
  const totalEks = Math.round(totalInkl / (1 + CONFIG.MVA_RATE));
  const mva = totalInkl - totalEks;

  sheet.getRange(row, COL.PRIS_INKL).setValue(totalInkl);
  sheet.getRange(row, COL.PRIS_EKS).setValue(totalEks);
  sheet.getRange(row, COL.MVA_BELOP).setValue(mva);
}

// ============================================================
// 6. OPPRETT NYTT OPPDRAG
// ============================================================
function createNewOppdrag_(parsed, date) {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) { Logger.log('FEIL: Fant ikke ark "' + CONFIG.SHEET_NAME + '"'); return; }

    // Duplikatsjekk
    if (sheet.getLastRow() > 1) {
      const existingData = sheet.getDataRange().getValues();
      for (let i = 1; i < existingData.length; i++) {
        if (existingData[i][COL.ADRESSE - 1] === parsed.adresse &&
          existingData[i][COL.TIMESTAMP - 1] &&
          new Date(existingData[i][COL.TIMESTAMP - 1]).toDateString() === date.toDateString()) {
          Logger.log('Duplikat: ' + parsed.adresse); return;
        }
      }
    }

    const lastRow = sheet.getLastRow();
    const oppdragsnr = 'NT-' + Utilities.formatDate(date, 'Europe/Oslo', 'yyyyMM') + '-' + String(lastRow).padStart(3, '0');

    let folderUrl = '';
    try {
      const folderName = (parsed.adresse || 'Ukjent') + ' - ' + Utilities.formatDate(date, 'Europe/Oslo', 'dd.MM.yyyy');
      const folder = createOppdragFolder_(folderName);
      folderUrl = folder.getUrl();
    } catch (folderErr) { Logger.log('ADVARSEL mappe: ' + folderErr.message); }

    const datoMottatt = Utilities.formatDate(date, 'Europe/Oslo', 'dd.MM.yyyy HH:mm');
    const timestamp = date.toISOString();

    let avstandKm = '', reiseEks = 0, reiseInkl = 0, reiseNotat = '';
    if (parsed.adresse) {
      const distResult = calculateDistance_(parsed.adresse);
      if (distResult) {
        const travelCost = calculateTravelCost_(distResult.kmTurRetur);
        avstandKm = distResult.kmTurRetur;
        reiseEks = travelCost.kostnadEksMva;
        reiseInkl = travelCost.kostnadInklMva;
        reiseNotat = distResult.kmTurRetur + ' km t/r' +
          (travelCost.fakturerbarKm > 0 ? ', ' + travelCost.fakturerbarKm + ' km fakturerbart' : ' (inkludert)');
      }
    }

    let fullNotater = '';
    if (parsed.notater) fullNotater += parsed.notater;
    if (reiseNotat) fullNotater += (fullNotater ? ' | ' : '') + 'Reise: ' + reiseNotat;

    const isIvit = parsed.kilde === 'IVIT';

    const newRow = [];
    newRow[COL.OPPDRAGSNR - 1]          = oppdragsnr;
    newRow[COL.DATO_MOTTATT - 1]        = datoMottatt;
    newRow[COL.KILDE - 1]               = parsed.kilde;
    newRow[COL.OPPDRAGSTYPE - 1]        = parsed.oppdragstype;
    newRow[COL.ADRESSE - 1]             = folderUrl
      ? '=HYPERLINK("' + folderUrl + '","' + parsed.adresse.replace(/"/g, '""') + '")'
      : parsed.adresse;
    newRow[COL.OPPDRAGSGIVER - 1]       = parsed.oppdragsgiver;
    newRow[COL.SELGER - 1]              = parsed.selger;
    newRow[COL.SELGER_TLF - 1]          = parsed.selgerTlf;
    newRow[COL.SELGER_EPOST - 1]        = parsed.selgerEpost;
    newRow[COL.MEGLER - 1]              = parsed.megler;
    newRow[COL.MEGLER_EPOST - 1]        = parsed.meglerEpost;
    newRow[COL.FAKTURA_REF - 1]         = parsed.fakturaRef;
    newRow[COL.FAKTURA_SENDES_TIL - 1]  = parsed.fakturaSendesTil;
    newRow[COL.FAKTURAMOTAKER - 1]      = '';
    newRow[COL.BOLIGTYPE - 1]           = '';
    newRow[COL.AREAL - 1]              = '';
    newRow[COL.ANTALL_TILLEGGSBYGG - 1] = '';
    newRow[COL.RAPPORTTYPE - 1]         = parsed.rapporttype || '';
    newRow[COL.MED_MARKEDSVERDI - 1]    = false;
    newRow[COL.TIMER - 1]              = '';
    newRow[COL.PRIS_INKL - 1]          = '';
    newRow[COL.PRIS_EKS - 1]           = '';
    newRow[COL.MVA_BELOP - 1]          = '';
    newRow[COL.AVSTAND_KM - 1]         = avstandKm;
    newRow[COL.REISE_EKS - 1]          = reiseEks;
    newRow[COL.REISE_INKL - 1]         = reiseInkl;
    newRow[COL.SUM_FERGE_BOM - 1]      = '';
    newRow[COL.ANTALL_DELE_REISE - 1]  = '';
    newRow[COL.STATUS - 1]             = 'Mottatt';
    newRow[COL.BEFARING_DATO - 1]      = '';
    newRow[COL.BEFARING_KL - 1]        = '';
    newRow[COL.DATO_STATUSENDRING - 1] = datoMottatt;
    newRow[COL.TIMESTAMP - 1]          = timestamp;
    newRow[COL.LINK_MAPPE - 1]         = folderUrl;
    newRow[COL.NOTATER - 1]            = fullNotater;
    newRow[COL.SCAN_IVIT - 1]          = isIvit;  // Auto-check for IVIT-oppdrag

    sheet.appendRow(newRow);
    const newRowNum = sheet.getLastRow();
    sheet.getRange(newRowNum, COL.MED_MARKEDSVERDI).insertCheckboxes();
    sheet.getRange(newRowNum, COL.SCAN_IVIT).insertCheckboxes();
    sheet.getRange(newRowNum, COL.MED_MARKEDSVERDI).setValue(false);
    sheet.getRange(newRowNum, COL.SCAN_IVIT).setValue(isIvit);

    Logger.log('✅ Oppdrag opprettet: ' + oppdragsnr + ' | ' + parsed.adresse);
  } catch (err) {
    Logger.log('❌ FEIL i createNewOppdrag_: ' + err.message);
  }
}

// ============================================================
// 7. ONEDIT — oppdatert med Q, AA, AB triggers
// ============================================================
function onEditHandler(e) {
  if (!e || !e.range) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.SHEET_NAME) return;
    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < 2) return;

    // Status
    if (col === COL.STATUS) handleStatusChange_(sheet, row, e.value);

    // Pris: boligtype, areal, rapporttype, markedsverdi, tilleggsbygg
    if ([COL.BOLIGTYPE, COL.AREAL, COL.RAPPORTTYPE, COL.MED_MARKEDSVERDI, COL.ANTALL_TILLEGGSBYGG].indexOf(col) > -1) {
      calculatePrice_(sheet, row);
    }

    // Manuell endring av pris inkl → synk eks og mva
    if (col === COL.PRIS_INKL) {
      const prisInkl = e.value ? parseFloat(String(e.value).replace(/[^\d]/g, '')) : 0;
      if (prisInkl > 0) {
        const prisEks = Math.round(prisInkl / (1 + CONFIG.MVA_RATE));
        sheet.getRange(row, COL.PRIS_EKS).setValue(prisEks);
        sheet.getRange(row, COL.MVA_BELOP).setValue(prisInkl - prisEks);
      }
    }

    // Befaring
    if (col === COL.BEFARING_DATO || col === COL.BEFARING_KL) {
      const befaringDato = sheet.getRange(row, COL.BEFARING_DATO).getValue();
      if (befaringDato) handleBefaringBooked_(sheet, row, befaringDato, sheet.getRange(row, COL.BEFARING_KL).getValue());
    }

    // Adresse endret → rekalk reise
    if (col === COL.ADRESSE && e.value) recalculateTravel_(sheet, row, e.value);

    // Reise: avstand km, ferge/bom, eller antall deler endret → rekalk total reise
    if ([COL.AVSTAND_KM, COL.SUM_FERGE_BOM, COL.ANTALL_DELE_REISE].indexOf(col) > -1) {
      recalculateReiseForRow_(sheet, row);
    }

    // Reise: manuell endring eks → synk inkl
    if (col === COL.REISE_EKS) {
      const eks = parseFloat(String(sheet.getRange(row, COL.REISE_EKS).getValue()).replace(',', '.'));
      if (!isNaN(eks) && eks >= 0) {
        sheet.getRange(row, COL.REISE_INKL).setValue(Math.round(eks * (1 + CONFIG.MVA_RATE)));
      }
    }

    // Reise: manuell endring inkl → synk eks
    if (col === COL.REISE_INKL) {
      const inkl = parseFloat(String(sheet.getRange(row, COL.REISE_INKL).getValue()).replace(',', '.'));
      if (!isNaN(inkl) && inkl >= 0) {
        sheet.getRange(row, COL.REISE_EKS).setValue(Math.round(inkl / (1 + CONFIG.MVA_RATE)));
      }
    }

    // Synk markedsverdi-checkbox når rapporttype endres
    if (col === COL.RAPPORTTYPE) {
      const rt = e.value || '';
      sheet.getRange(row, COL.MED_MARKEDSVERDI).setValue(rt === 'Tilstandsrapport m/teknisk og markedsverdi');
    }

    // Dashboard refresh
    if ([COL.OPPDRAGSTYPE, COL.BOLIGTYPE, COL.AREAL, COL.RAPPORTTYPE,
         COL.PRIS_INKL, COL.AVSTAND_KM, COL.REISE_EKS, COL.REISE_INKL,
         COL.STATUS, COL.ANTALL_TILLEGGSBYGG, COL.MED_MARKEDSVERDI,
         COL.SUM_FERGE_BOM, COL.ANTALL_DELE_REISE].indexOf(col) > -1) {
      updateDashboard();
    }
  } finally {
    lock.releaseLock();
  }
}

function handleStatusChange_(sheet, row, newStatus) {
  const now = new Date();
  const datoStr = Utilities.formatDate(now, 'Europe/Oslo', 'dd.MM.yyyy HH:mm');
  sheet.getRange(row, COL.DATO_STATUSENDRING).setValue(datoStr);
  const rowData = sheet.getRange(row, 1, 1, NUM_COLS).getValues()[0];
  const oppdragsnr = rowData[COL.OPPDRAGSNR - 1];
  const adresse = rowData[COL.ADRESSE - 1];
  const prisInkl = rowData[COL.PRIS_INKL - 1] || 0;
  const reiseInkl = rowData[COL.REISE_INKL - 1] || 0;

  if (newStatus === 'Kan faktureres') {
    safeSendEmail_(CONFIG.ACCOUNTANT_EMAIL, '💰 Klar til fakturering: ' + adresse + ' (' + oppdragsnr + ')', buildFakturaEmail_(rowData, datoStr));
  }
  if (newStatus === 'Oppdrag kansellert') {
    sheet.getRange(row, 1, 1, NUM_COLS).setBackground('#eeeeee');
    return;
  }
  if (newStatus === 'Oppdrag fullført') {
    const folderId = extractDriveFolderId_(rowData[COL.LINK_MAPPE - 1]);
    if (folderId) { try { moveOppdragFolderToAvsluttede_(folderId); } catch (e) { Logger.log('Flytt feilet: ' + e.message); } }
    return;
  }
  if (newStatus === 'Fakturert') {
    sheet.getRange(row, 1, 1, NUM_COLS).setBackground('#c8e6c9');
    archiveRow_(sheet, row);
  }
}

function handleBefaringBooked_(sheet, row, befaringDato, befaringTid) {
  const rowData = sheet.getRange(row, 1, 1, NUM_COLS).getValues()[0];
  const currentStatus = rowData[COL.STATUS - 1];
  if (currentStatus === 'Mottatt') {
    sheet.getRange(row, COL.STATUS).setValue('Avtalt befaring');
    sheet.getRange(row, COL.DATO_STATUSENDRING).setValue(Utilities.formatDate(new Date(), 'Europe/Oslo', 'dd.MM.yyyy HH:mm'));
  }
}

function recalculateTravel_(sheet, row, address) {
  const distResult = calculateDistance_(address);
  if (distResult) {
    sheet.getRange(row, COL.AVSTAND_KM).setValue(distResult.kmTurRetur);
    recalculateReiseForRow_(sheet, row);
  }
}

// ============================================================
// 8. IVIT WEBHOOK SCRAPING (integrert)
// ============================================================
/**
 * Hent oppdragsdata fra iVit via webhook-scraper på Sliplane.
 */
function fetchIvitData_(address) {
  if (!address || address.trim().length === 0) throw new Error('Address is required');
  const payload = JSON.stringify({ address: address.trim() });
  const options = {
    method: 'post', contentType: 'application/json',
    payload: payload, muteHttpExceptions: true,
  };
  if (CONFIG.IVIT_WEBHOOK_SECRET) {
    options.headers = { 'Authorization': 'Bearer ' + CONFIG.IVIT_WEBHOOK_SECRET };
  }
  try {
    const response = UrlFetchApp.fetch(CONFIG.IVIT_WEBHOOK_URL, options);
    const body = JSON.parse(response.getContentText());
    if (response.getResponseCode() === 200 && body.success) {
      Logger.log('iVit data mottatt for: ' + address);
      return body;
    } else {
      Logger.log('iVit feil: ' + (body.error || 'Ukjent'));
      return body;
    }
  } catch (e) {
    Logger.log('iVit request feilet: ' + e.message);
    return { success: false, error: 'Request feilet: ' + e.message, address: address };
  }
}

/**
 * Prosesser IVIT-rader som har "Scan IVIT" (AJ) avhuket.
 * Filtrerer på: Kilde=IVIT, Dato mottatt > cutoff, Scan IVIT=true.
 * Skriver resultater og fjerner avhukingen etterpå.
 */
function processIvitRows() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) { Logger.log('Fant ikke ark "' + CONFIG.SHEET_NAME + '"'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  let processed = 0;

  for (let row = 2; row <= lastRow; row++) {
    // Sjekk "Scan IVIT" checkbox først (raskest filter)
    const scanIvit = sheet.getRange(row, COL.SCAN_IVIT).getValue();
    if (scanIvit !== true) continue;

    const kilde = sheet.getRange(row, COL.KILDE).getValue().toString().trim();
    if (kilde !== 'IVIT') continue;

    // Parse og sjekk dato mottatt
    const datoMottatt = sheet.getRange(row, COL.DATO_MOTTATT).getValue();
    if (!datoMottatt) continue;
    let mottattDate;
    if (datoMottatt instanceof Date) {
      mottattDate = datoMottatt;
    } else {
      const parts = datoMottatt.toString().match(/(\d{2})\.(\d{2})\.(\d{4})\s*(\d{2}):(\d{2})/);
      if (!parts) continue;
      mottattDate = new Date(parts[3], parts[2] - 1, parts[1], parts[4], parts[5]);
    }
    if (mottattDate <= CONFIG.IVIT_CUTOFF_DATE) continue;

    // Hent adresse (kan være HYPERLINK-formel)
    let address = sheet.getRange(row, COL.ADRESSE).getDisplayValue().toString().trim();
    if (!address) continue;

    Logger.log('Rad ' + row + ': scraper "' + address + '"');
    const result = fetchIvitData_(address);

    if (result.success) {
      const d = result.data;
      if (d.befaring_dato) sheet.getRange(row, COL.BEFARING_DATO).setValue(d.befaring_dato);
      if (d.befaring_klokkeslett) sheet.getRange(row, COL.BEFARING_KL).setValue(d.befaring_klokkeslett);
      if (d.fakturareferanse) sheet.getRange(row, COL.FAKTURA_REF).setValue(d.fakturareferanse);
      if (d.med_markedsverdi) {
        const isMed = String(d.med_markedsverdi).toLowerCase().indexOf('med') > -1;
        sheet.getRange(row, COL.MED_MARKEDSVERDI).setValue(isMed);
      }
      // Fjern avhuking og sett timestamp
      sheet.getRange(row, COL.SCAN_IVIT).setValue(false);
      Logger.log('Rad ' + row + ': OK');
      processed++;
    } else {
      sheet.getRange(row, COL.NOTATER).setValue('iVit feil: ' + (result.error || 'Ukjent feil'));
      // La avhukingen stå slik at den kan prøves igjen
      Logger.log('Rad ' + row + ': Feil — ' + result.error);
    }

    SpreadsheetApp.flush();
    Utilities.sleep(2000);
  }

  Logger.log('IVIT scan ferdig. Prosesserte ' + processed + ' rad(er).');
}

// ============================================================
// 9. PÅMINNELSER
// ============================================================
function checkReminders() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    const status = data[i][COL.STATUS - 1];
    const timestamp = data[i][COL.TIMESTAMP - 1];
    const adresse = data[i][COL.ADRESSE - 1];
    const oppdragsnr = data[i][COL.OPPDRAGSNR - 1];
    if (status !== 'Mottatt' || !timestamp) continue;
    const timerSiden = (now - new Date(timestamp)) / (1000 * 60 * 60);
    if (timerSiden >= CONFIG.URGENT_HOURS && timerSiden < CONFIG.URGENT_HOURS + 1.5) {
      sendChatAlert_('🚨 *HASTER:* ' + adresse + ' (' + oppdragsnr + ') — ' + Math.round(timerSiden) + ' timer!');
    }
  }
}

// ============================================================
// 10. UKENTLIG RAPPORT
// ============================================================
function sendWeeklyReport() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const monday = new Date(now);
  monday.setDate(now.getDate() - (now.getDay() === 0 ? 6 : now.getDay() - 1));
  monday.setHours(0, 0, 0, 0);

  const fakturerbare = [], fakturerte = [], ventende = [];
  let totalInklMva = 0, totalEksMva = 0, totalMva = 0, totalReiseEks = 0, totalReiseInkl = 0;

  for (let i = 1; i < data.length; i++) {
    const status = data[i][COL.STATUS - 1];
    const prisInkl = data[i][COL.PRIS_INKL - 1] || 0;
    const prisEks = data[i][COL.PRIS_EKS - 1] || 0;
    const mvaB = data[i][COL.MVA_BELOP - 1] || 0;
    const reiseEks = data[i][COL.REISE_EKS - 1] || 0;
    const reiseInkl = data[i][COL.REISE_INKL - 1] || 0;
    const statusDatoStr = data[i][COL.DATO_STATUSENDRING - 1];
    let sDato = parseDateString_(statusDatoStr);
    if (!sDato) continue;
    const item = {
      oppdragsnr: data[i][COL.OPPDRAGSNR - 1], adresse: data[i][COL.ADRESSE - 1],
      oppdragstype: data[i][COL.OPPDRAGSTYPE - 1], boligtype: data[i][COL.BOLIGTYPE - 1],
      prisInkl: prisInkl, prisEks: prisEks, mvaBeløp: mvaB,
      reiseEks: reiseEks, reiseInkl: reiseInkl,
      fakturaRef: data[i][COL.FAKTURA_REF - 1], fakturaTil: data[i][COL.FAKTURA_SENDES_TIL - 1],
      statusDato: statusDatoStr
    };
    if (status === 'Kan faktureres' && sDato >= monday) {
      fakturerbare.push(item);
      totalInklMva += prisInkl; totalEksMva += prisEks; totalMva += mvaB;
      totalReiseEks += reiseEks; totalReiseInkl += reiseInkl;
    }
    if (status === 'Fakturert' && sDato >= monday) fakturerte.push(item);
    if (status === 'Kan faktureres') ventende.push(item);
  }

  const html = buildWeeklyReportHtml_(fakturerbare, fakturerte, ventende, {
    totalInklMva: totalInklMva, totalEksMva: totalEksMva, totalMva: totalMva,
    totalReiseEks: totalReiseEks, totalReiseInkl: totalReiseInkl
  });
  safeSendEmail_(CONFIG.OWNER_EMAIL + ',' + CONFIG.ACCOUNTANT_EMAIL,
    '📊 Naava Takst — Ukerapport uke ' + getWeekNumber_(now), html);
}

// ============================================================
// 11. DASHBOARD
// ============================================================
function updateDashboard() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  let dash = ss.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);
  if (!sheet || !dash) return;
  if (sheet.getLastRow() < 2) { dash.clear(); dash.getRange(1, 1).setValue('Ingen oppdrag ennå.'); return; }

  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const curMonth = now.getMonth(), curYear = now.getFullYear();
  const statusList = ['Mottatt', 'Avtalt befaring', 'Befart', 'Utkast', 'Endelig rapport', 'Kan faktureres', 'Fakturert', 'Oppdrag kansellert', 'Oppdrag fullført'];

  const stats = {
    total: data.length - 1, byStatus: {}, byType: {}, byBolig: {}, byBoligBucket: {},
    omsMåned: 0, omsÅr: 0, reiseMåned: 0, reiseÅr: 0,
    uteståendeInkl: 0, uteståendeEks: 0, snittPrisInkl: 0, snittPrisEks: 0, snittAntall: 0,
    trendMonths: {}
  };
  statusList.forEach(function (s) { stats.byStatus[s] = 0; });

  const monthsToShow = 6;
  const monthKeys = [];
  for (let k = monthsToShow - 1; k >= 0; k--) {
    const d = new Date(now.getFullYear(), now.getMonth() - k, 1);
    const key = Utilities.formatDate(d, 'Europe/Oslo', 'yyyy-MM');
    monthKeys.push(key);
    stats.trendMonths[key] = { omsInkl: 0, reiseInkl: 0, totalInkl: 0 };
  }

  let sumPrisInkl = 0, sumPrisEks = 0, countPris = 0;
  for (let i = 1; i < data.length; i++) {
    const status = data[i][COL.STATUS - 1];
    const type = data[i][COL.OPPDRAGSTYPE - 1];
    const bolig = data[i][COL.BOLIGTYPE - 1];
    const areal = data[i][COL.AREAL - 1] ? Number(data[i][COL.AREAL - 1]) : 0;
    const prisInkl = Number(data[i][COL.PRIS_INKL - 1] || 0);
    const prisEks = Number(data[i][COL.PRIS_EKS - 1] || 0);
    const reiseInkl = Number(data[i][COL.REISE_INKL - 1] || 0);

    if (stats.byStatus[status] !== undefined) stats.byStatus[status]++;
    if (type) stats.byType[type] = (stats.byType[type] || 0) + 1;
    if (bolig) stats.byBolig[bolig] = (stats.byBolig[bolig] || 0) + 1;
    const bucket = classifyBoligBucket_(bolig, areal);
    if (bucket) stats.byBoligBucket[bucket] = (stats.byBoligBucket[bucket] || 0) + 1;

    if (status === 'Kan faktureres') {
      stats.uteståendeInkl += (prisInkl + reiseInkl);
      stats.uteståendeEks += Math.round((prisInkl + reiseInkl) / (1 + CONFIG.MVA_RATE));
    }
    if (status === 'Fakturert' || status === 'Kan faktureres') {
      stats.omsÅr += prisInkl; stats.reiseÅr += reiseInkl;
      const statusDato = parseDateString_(data[i][COL.DATO_STATUSENDRING - 1]);
      if (statusDato) {
        if (statusDato.getMonth() === curMonth && statusDato.getFullYear() === curYear) {
          stats.omsMåned += prisInkl; stats.reiseMåned += reiseInkl;
        }
        const mKey = Utilities.formatDate(new Date(statusDato.getFullYear(), statusDato.getMonth(), 1), 'Europe/Oslo', 'yyyy-MM');
        if (stats.trendMonths[mKey]) {
          stats.trendMonths[mKey].omsInkl += prisInkl;
          stats.trendMonths[mKey].reiseInkl += reiseInkl;
          stats.trendMonths[mKey].totalInkl += (prisInkl + reiseInkl);
        }
      }
      if (prisInkl > 0) { sumPrisInkl += prisInkl; sumPrisEks += prisEks; countPris += 1; }
    }
  }
  stats.snittAntall = countPris;
  stats.snittPrisInkl = countPris ? Math.round(sumPrisInkl / countPris) : 0;
  stats.snittPrisEks = countPris ? Math.round(sumPrisEks / countPris) : 0;

  dash.clear();
  const rows = [];
  rows.push(['🏢 NAAVA TAKST DASHBOARD' + (CONFIG.TEST_MODE ? ' 🧪 TEST' : ''), '', '', '', '']);
  rows.push([Utilities.formatDate(now, 'Europe/Oslo', 'dd.MM.yyyy HH:mm'), '', '', '', '']);
  rows.push(['', '', '', '', '']);
  rows.push(['📊 OVERSIKT', '', '', '', '']);
  rows.push(['Totalt antall oppdrag', stats.total, '', '', '']);
  rows.push(['Aktive', stats.total - (stats.byStatus['Fakturert'] || 0) - (stats.byStatus['Oppdrag kansellert'] || 0) - (stats.byStatus['Oppdrag fullført'] || 0), '', '', '']);
  rows.push(['Utestående (Eks mva)', formatCurrency_(stats.uteståendeEks), '', '', '']);
  rows.push(['Utestående (Inkl mva)', formatCurrency_(stats.uteståendeInkl), '', '', '']);
  rows.push(['Snitt pris (Eks mva)', formatCurrency_(stats.snittPrisEks), '', '', '']);
  rows.push(['Snitt pris (Inkl mva)', formatCurrency_(stats.snittPrisInkl), '', '', '']);
  rows.push(['', '', '', '', '']);
  rows.push(['💰 OMSETNING', 'Eks mva', 'Inkl mva', '', '']);
  rows.push(['Måned', formatCurrency_(Math.round(stats.omsMåned / (1 + CONFIG.MVA_RATE))), formatCurrency_(stats.omsMåned), '', '']);
  rows.push(['År', formatCurrency_(Math.round(stats.omsÅr / (1 + CONFIG.MVA_RATE))), formatCurrency_(stats.omsÅr), '', '']);
  rows.push(['', '', '', '', '']);
  rows.push(['🚗 REISE', 'Eks mva', 'Inkl mva', '', '']);
  rows.push(['Måned', formatCurrency_(Math.round(stats.reiseMåned / (1 + CONFIG.MVA_RATE))), formatCurrency_(stats.reiseMåned), '', '']);
  rows.push(['År', formatCurrency_(Math.round(stats.reiseÅr / (1 + CONFIG.MVA_RATE))), formatCurrency_(stats.reiseÅr), '', '']);
  rows.push(['', '', '', '', '']);
  rows.push(['📋 STATUS', 'Antall', '', '', '']);
  statusList.forEach(function (s) { rows.push([s, stats.byStatus[s] || 0, '', '', '']); });
  rows.push(['', '', '', '', '']);
  rows.push(['📌 OPPDRAGSTYPER', 'Antall', '', '', '']);
  Object.keys(stats.byType).sort().forEach(function (t) { rows.push([t, stats.byType[t], '', '', '']); });
  rows.push(['', '', '', '', '']);
  rows.push(['🏠 BOLIGSTØRRELSE', 'Antall', '', '', '']);
  Object.keys(stats.byBoligBucket).sort().forEach(function (k) { rows.push([k, stats.byBoligBucket[k], '', '', '']); });
  rows.push(['', '', '', '', '']);
  rows.push(['📈 INNTJENING SISTE ' + monthsToShow + ' MND (Inkl mva)', 'Oppdrag', 'Reise', 'Total', '']);
  monthKeys.forEach(function (k) { const v = stats.trendMonths[k]; rows.push([k, v.omsInkl, v.reiseInkl, v.totalInkl, '']); });

  dash.getRange(1, 1, rows.length, 5).setValues(rows);
  dash.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  dash.setColumnWidth(1, 360); dash.setColumnWidth(2, 180); dash.setColumnWidth(3, 180); dash.setColumnWidth(4, 180);
}

function classifyBoligBucket_(boligtype, areal) {
  if (!boligtype) return '';
  const a = Number(areal || 0);
  if (boligtype === 'Leilighet') return a && a <= 80 ? 'Leilighet (0–80 m²)' : 'Leilighet (80+ m²)';
  if (boligtype === 'Rekkehus/leilighet 2-4-mannsbolig') return a && a <= 80 ? 'Rekkehus/2-4 (0–80 m²)' : 'Rekkehus/2-4 (80+ m²)';
  if (boligtype === 'Enebolig/fritidsbolig') {
    if (a && a <= 150) return 'Enebolig/fritid (0–150 m²)';
    if (a && a <= 250) return 'Enebolig/fritid (150–250 m²)';
    return 'Enebolig/fritid (250+ m²)';
  }
  if (boligtype === 'Frittstående bygg') return 'Frittstående bygg';
  return boligtype;
}

// ============================================================
// HTML BUILDERS
// ============================================================
function buildFakturaEmail_(rowData, datoStr) {
  const fields = [
    ['Oppdragsnr', rowData[COL.OPPDRAGSNR - 1]], ['Type', rowData[COL.OPPDRAGSTYPE - 1]],
    ['Adresse', rowData[COL.ADRESSE - 1]], ['Megler', rowData[COL.MEGLER - 1]],
    ['Boligtype', rowData[COL.BOLIGTYPE - 1]], ['Rapporttype', rowData[COL.RAPPORTTYPE - 1]],
    ['Faktura ref', rowData[COL.FAKTURA_REF - 1]], ['Faktura til', rowData[COL.FAKTURA_SENDES_TIL - 1]],
  ];
  const prisInkl = rowData[COL.PRIS_INKL - 1] || 0;
  const prisEks = rowData[COL.PRIS_EKS - 1] || 0;
  const mva = rowData[COL.MVA_BELOP - 1] || 0;
  const reiseEks = rowData[COL.REISE_EKS - 1] || 0;
  const reiseInkl = rowData[COL.REISE_INKL - 1] || 0;
  let info = ''; fields.forEach(function (f, i) {
    if (f[1]) info += '<tr style="background:' + (i % 2 === 0 ? '#fff' : '#f9f9f9') + '"><td style="padding:8px;font-weight:bold;">' + f[0] + '</td><td style="padding:8px;">' + f[1] + '</td></tr>';
  });
  const prices =
    '<tr><td style="padding:8px;font-weight:bold;">Oppdrag eks mva</td><td style="padding:8px;text-align:right;">' + formatCurrency_(prisEks) + '</td></tr>' +
    '<tr style="background:#f9f9f9"><td style="padding:8px;font-weight:bold;">MVA 25%</td><td style="padding:8px;text-align:right;">' + formatCurrency_(mva) + '</td></tr>' +
    '<tr><td style="padding:8px;font-weight:bold;">Oppdrag inkl mva</td><td style="padding:8px;text-align:right;font-weight:bold;">' + formatCurrency_(prisInkl) + '</td></tr>' +
    '<tr style="background:#f9f9f9"><td style="padding:8px;font-weight:bold;">Reise eks mva</td><td style="padding:8px;text-align:right;">' + formatCurrency_(reiseEks) + '</td></tr>' +
    '<tr><td style="padding:8px;font-weight:bold;">Reise inkl mva</td><td style="padding:8px;text-align:right;">' + formatCurrency_(reiseInkl) + '</td></tr>' +
    '<tr style="background:#1a5c2a;color:white;font-size:15px;"><td style="padding:12px;font-weight:bold;">TOTAL INKL MVA</td><td style="padding:12px;text-align:right;font-weight:bold;">' + formatCurrency_(prisInkl + reiseInkl) + '</td></tr>';
  return '<div style="font-family:Arial,sans-serif;max-width:600px;">' +
    '<div style="background:#1a5c2a;color:white;padding:20px;border-radius:8px 8px 0 0;"><h2 style="margin:0;">💰 Klar til fakturering</h2></div>' +
    '<div style="padding:20px;border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;">' +
    '<table style="width:100%;border-collapse:collapse;">' + info + '</table>' +
    '<hr style="margin:16px 0;border:none;border-top:2px solid #1a5c2a;">' +
    '<table style="width:100%;border-collapse:collapse;">' + prices + '</table></div></div>';
}

function buildWeeklyReportHtml_(fakturerbare, fakturerte, ventende, totals) {
  let rows = '';
  fakturerbare.forEach(function (o) {
    rows += '<tr><td style="padding:5px;border-bottom:1px solid #eee;">' + o.oppdragsnr + '</td><td style="padding:5px;">' + o.adresse + '</td><td style="padding:5px;">' + o.oppdragstype + '</td><td style="padding:5px;">' + (o.fakturaRef || '-') + '</td><td style="padding:5px;text-align:right;">' + formatCurrency_(o.prisEks) + '</td><td style="padding:5px;text-align:right;">' + formatCurrency_(o.mvaBeløp) + '</td><td style="padding:5px;text-align:right;font-weight:bold;">' + formatCurrency_(o.prisInkl) + '</td><td style="padding:5px;text-align:right;">' + formatCurrency_(o.reiseInkl) + '</td></tr>';
  });
  return '<div style="font-family:Arial,sans-serif;max-width:900px;">' +
    '<div style="background:#1a5c2a;color:white;padding:20px;border-radius:8px 8px 0 0;"><h2 style="margin:0;">📊 Ukerapport uke ' + getWeekNumber_(new Date()) + '</h2></div>' +
    '<div style="padding:20px;border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;">' +
    '<h3>💰 Fakturerbart (' + fakturerbare.length + ')</h3>' +
    (fakturerbare.length > 0 ?
      '<table style="width:100%;border-collapse:collapse;font-size:11px;"><tr style="background:#f5f5f5;"><th style="padding:5px;text-align:left;">Nr</th><th>Adresse</th><th>Type</th><th>Ref</th><th style="text-align:right;">Eks</th><th style="text-align:right;">MVA</th><th style="text-align:right;">Inkl</th><th style="text-align:right;">Reise</th></tr>' + rows +
      '<tr style="background:#1a5c2a;color:white;font-weight:bold;"><td colspan="4">TOTAL</td><td colspan="4" style="text-align:right;font-size:14px;">' + formatCurrency_(totals.totalInklMva + totals.totalReiseInkl) + '</td></tr></table>' :
      '<p style="color:#666;">Ingen denne uken.</p>') +
    '</div></div>';
}

// ============================================================
// SETUP & TRIGGERS
// ============================================================
function setupConditionalFormatting_(sheet, colCount) {
  const range = sheet.getRange(2, 1, 500, colCount);
  const rules = [
    { status: 'Mottatt', color: '#fff3e0' }, { status: 'Avtalt befaring', color: '#e3f2fd' },
    { status: 'Befart', color: '#e8eaf6' }, { status: 'Utkast', color: '#fce4ec' },
    { status: 'Endelig rapport', color: '#f3e5f5' }, { status: 'Kan faktureres', color: '#e8f5e9' },
    { status: 'Fakturert', color: '#c8e6c9' }, { status: 'Oppdrag kansellert', color: '#eeeeee' },
    { status: 'Oppdrag fullført', color: '#d9ead3' },
  ];
  sheet.setConditionalFormatRules(rules.map(function (r) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + STATUS_COL_LETTER + '2="' + r.status + '"')
      .setBackground(r.color).setRanges([range]).build();
  }));
}

function setupDashboard_(ss) {
  let dash = ss.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);
  if (!dash) dash = ss.insertSheet(CONFIG.DASHBOARD_SHEET_NAME);
  dash.clear();
  dash.getRange(1, 1).setValue('📊 Dashboard');
}

function setupTriggers_(ss) {
  const fns = ['scanIncomingEmails', 'checkReminders', 'sendWeeklyReport', 'updateDashboard', 'onEditHandler', 'processIvitRows'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (fns.indexOf(t.getHandlerFunction()) > -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scanIncomingEmails').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('processIvitRows').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('checkReminders').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('sendWeeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(16).create();
  ScriptApp.newTrigger('updateDashboard').timeBased().everyHours(1).create();
  if (!ss) ss = getSpreadsheet_();
  if (ss) ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss).onEdit().create();
}

function archiveRow_(sheet, row) {
  const ss = getSpreadsheet_();
  const archive = ss.getSheetByName('arkiv');
  if (!archive) return;
  const rowData = sheet.getRange(row, 1, 1, NUM_COLS).getValues()[0];
  if (archive.getLastRow() === 0) archive.getRange(1, 1, 1, NUM_COLS).setValues(sheet.getRange(1, 1, 1, NUM_COLS).getValues());
  const oppdragsnr = rowData[COL.OPPDRAGSNR - 1];
  if (archive.getLastRow() >= 2) {
    const existing = archive.getRange(2, 1, archive.getLastRow() - 1, 1).getValues().flat();
    if (existing.indexOf(oppdragsnr) > -1) return;
  }
  archive.appendRow(rowData);
}

// ============================================================
// MENY
// ============================================================
function onOpen() {
  const mode = CONFIG.TEST_MODE ? ' 🧪' : '';
  SpreadsheetApp.getUi()
    .createMenu('🏠 Naava Takst' + mode)
    .addItem('📝 Manuell inntak', 'setupManualIntakeSheet_')
    .addItem('✅ Registrer fra inntak', 'registerManualFromSheet_')
    .addSeparator()
    .addItem('🔍 Scan e-post nå', 'scanIncomingEmails')
    .addItem('🔄 Scan IVIT nå', 'processIvitRows')
    .addItem('🧪 Test IVIT-tilkobling', 'testIvitConnection_')
    .addSeparator()
    .addItem('📊 Oppdater dashboard', 'updateDashboard')
    .addItem('📧 Send ukerapport', 'testWeeklyReport')
    .addItem('⏰ Sjekk påminnelser', 'checkReminders')
    .addSeparator()
    .addItem('⚙️ Kjør oppsett', 'initialSetup')
    .addToUi();
}

function testIvitConnection_() {
  const result = fetchIvitData_('Smithsgata 5, 6100 VOLDA');
  const ui = SpreadsheetApp.getUi();
  if (result.success) {
    ui.alert('✅ IVIT-tilkobling OK!\n\n' + JSON.stringify(result.data, null, 2));
  } else {
    ui.alert('❌ IVIT-feil: ' + (result.error || 'Ukjent'));
  }
}

function testWeeklyReport() { sendWeeklyReport(); SpreadsheetApp.getUi().alert('✅ Ukerapport sendt!'); }

// ============================================================
// MANUELL INNTAK
// ============================================================
function setupManualIntakeSheet_() {
  const ss = getSpreadsheet_();
  const name = 'Manuell inntak';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1).setValue('Manuell registrering').setFontWeight('bold');
  const rows = [
    ['Adresse *', ''], ['Selger/kontaktperson', ''], ['Telefon', ''],
    ['E-post (kunde)', ''], ['Boligtype *', ''], ['Rapporttype *', ''],
    ['Areal (ca. m²)', ''], ['Megler / bestiller', ''], ['Merknad', '']
  ];
  sh.getRange(3, 1, rows.length, 2).setValues(rows);
  sh.setColumnWidth(1, 360); sh.setColumnWidth(2, 420);
  sh.getRange(3, 1, rows.length, 1).setFontWeight('bold');
  sh.getRange(3, 2, rows.length, 1).setBackground('#fffde7');
  sh.getRange(14, 1).setValue('Status:'); sh.getRange(14, 2).setValue('Klar');
  const boligTyper = ['Leilighet', 'Rekkehus/leilighet 2-4-mannsbolig', 'Enebolig/fritidsbolig', 'Frittstående bygg', 'Næringsbygg', 'Annet'];
  const rapportTyper = ['Tilstandsrapport m/teknisk og markedsverdi', 'Tilstandsrapport', 'Skadetakstrapport', 'Reklamasjonsrapport', 'Vurderingsrapport', 'Overtagelsesrapport', 'Annen rapport'];
  sh.getRange(7, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(boligTyper, true).setAllowInvalid(false).build());
  sh.getRange(8, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(rapportTyper, true).setAllowInvalid(false).build());
}

function registerManualFromSheet_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName('Manuell inntak');
  if (!sh) throw new Error('Fant ikke "Manuell inntak"');
  const v = (r) => String(sh.getRange(r, 2).getValue() || '').trim();
  const data = {
    adresse: v(3), selger: v(4), telefon: v(5), epost: v(6),
    boligtype: v(7), rapporttype: v(8), areal: v(9), megler: v(10), merknad: v(11)
  };
  if (!data.adresse || !data.boligtype || !data.rapporttype) {
    sh.getRange(14, 2).setValue('Feil: fyll ut Adresse, Boligtype, Rapporttype'); return;
  }
  sh.getRange(14, 2).setValue('Registrerer...');

  const mapped = mapManualTypes_(data.boligtype, data.rapporttype);
  let parsed = {
    kilde: 'Manuell', oppdragstype: mapped.oppdragstype, adresse: data.adresse,
    oppdragsgiver: data.megler || data.selger || '', selger: data.selger || '',
    selgerTlf: String(data.telefon || '').replace(/[^\d+]/g, ''), selgerEpost: data.epost || '',
    megler: data.megler || '', meglerEpost: '', fakturaRef: '', fakturaSendesTil: '',
    notater: data.merknad || '', rapporttype: mapped.rapporttype
  };
  if (CONFIG.TEST_MODE) parsed = sanitizeParsedData_(parsed);
  createNewOppdrag_(parsed, new Date());

  const logg = ss.getSheetByName(CONFIG.SHEET_NAME);
  const row = logg.getLastRow();
  if (mapped.boligtype) logg.getRange(row, COL.BOLIGTYPE).setValue(mapped.boligtype);
  if (data.areal) logg.getRange(row, COL.AREAL).setValue(Number(data.areal));
  calculatePrice_(logg, row);
  sh.getRange(3, 2, 9, 1).clearContent();
  sh.getRange(14, 2).setValue('OK: registrert');
}

function mapManualTypes_(boligtypeUi, rapporttypeUi) {
  const rt = String(rapporttypeUi || '').toLowerCase();
  let oppdragstype = 'Annet', rapporttype = 'Annen rapport';
  if (rt === 'tilstandsrapport') { oppdragstype = 'Tilstandsrapport'; rapporttype = 'Tilstandsrapport'; }
  else if (rt.indexOf('markedsverdi') > -1) { oppdragstype = 'Tilstandsrapport m/markedsverdi'; rapporttype = 'Tilstandsrapport m/teknisk og markedsverdi'; }
  else if (rt.indexOf('skadetakst') > -1) { oppdragstype = 'Skadetakst'; rapporttype = 'Skadetakstrapport'; }
  else if (rt.indexOf('reklamasjon') > -1) { oppdragstype = 'Reklamasjon'; rapporttype = 'Reklamasjonsrapport'; }
  return { boligtype: String(boligtypeUi || ''), oppdragstype: oppdragstype, rapporttype: rapporttype };
}
