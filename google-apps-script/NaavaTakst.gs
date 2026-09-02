// ============================================================
// NAAVA TAKST — OPPDRAGSHÅNDTERING v5
// ============================================================
// ⚠️ TESTMODUS: Alle e-poster sendes KUN til adm@afki.no
//    og jacob.e.holen@gmail.com. Ingen eksterne mottakere.
// ⚠️ Gmail-scan krever avsender jacob.e.holen@gmail.com
//    for å unngå å fange opp eksisterende e-poster.
// ============================================================

// ---- KONFIGURASJON ----
const CONFIG = {
  // ======= TESTMODUS =======
  TEST_MODE: false,  // ← Sett til false for produksjon
  // ==========================

  // ⚠️ VIKTIG: Sett denne til ID fra din Google Sheet URL
  // URL: https://docs.google.com/spreadsheets/d/DENNE_ID/edit
  SHEET_ID: '1VEzaCNEkvbWYZf0UOrj6IG5PUQB3hL1enB24PvsQ1MI',

  OWNER_EMAIL: 'jacob@naava.no',
  ACCOUNTANT_EMAIL: 'regnskap@naava.no',
  ROOT_FOLDER_NAME: 'Naava',
  ROOT_FOLDER_ID: '1nDgJrHWnEdkkG1OR90vGmHby4CTYA7J_',
  SHEET_NAME: 'Oppdragslogg',
  DASHBOARD_SHEET_NAME: 'Dashboard',
  CHAT_WEBHOOK_URL: 'SETT_INN_WEBHOOK_URL_HER',

  // Takstmannens utgangspunkt for reiseberegning
  BASE_ADDRESS: 'Postveien 15, 6018 Ålesund',

  // Reisekostnad
  REISE_SATS_EKS_MVA: 10,
  REISE_INKLUDERT_KM: 50,

  // Claude API — nøkkelen leses fra Script Properties, ikke fra kildekoden.
  // Sett den under Prosjektinnstillinger > Script Properties i Apps Script-editoren.
  OPENAI_API_KEY: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
  OPENAI_MODEL: 'gpt-5.2',

  // ⚠️ Gmail-søk: I testmodus kreves avsender fra TEST_SENDER
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

  IVIT_WEBHOOK_URL: 'https://naavaivit.sliplane.app/webhook', // <-- Endre til din URL
  IVIT_WEBHOOK_SECRET: 'b9a863c2f947e4d54b6feda001cb15c0ad5ec49ce2450e6f591c1528d85b85b8', // <-- Legg inn secret hvis du har det

  // For sortering
  STATUS_PRIORITY: {
    'Mottatt': 1,
    'Avtalt befaring': 2,
    'Befart': 3,
    'Utkast': 4,
    'Endelig rapport': 5,
    'Kan faktureres': 6,
    'Fakturert': 7,
    'Oppdrag fullført': 8
  },
};


// ============================================================
// KOLONNEINDEKSER (1-basert, matcher rekkefølgen i regnearket)
// ============================================================
const COL = {
  KAN_FAKTURERES: 1,   // <-- NY KOLONNE (A)
  OPPDRAGSNR: 2,
  DATO_MOTTATT: 3,
  KILDE: 4,
  SCAN_IVIT: 5,
  OPPDRAGSTYPE: 6,
  ADRESSE: 7,
  OPPDRAGSGIVER: 8,
  SELGER: 9,
  SELGER_TLF: 10,
  SELGER_EPOST: 11,
  MEGLER: 12,
  MEGLER_EPOST: 13,
  FAKTURA_REF: 14,
  STATUS: 15,  // <-- Har flyttet seg fra N til O
  FAKTURA_SENDES_TIL: 16,
  FAKTURAMOTAKER: 17,
  BOLIGTYPE: 18,
  AREAL: 19,
  ANTALL_TILLEGGSBYGG: 20,
  RAPPORTTYPE: 21,
  MED_MARKEDSVERDI: 22,
  TIMER: 23,
  PRIS_INKL: 24,
  PRIS_EKS: 25,
  MVA_BELOP: 26,
  AVSTAND_KM: 27,
  REISE_EKS: 28,
  REISE_INKL: 29,
  SUM_FERGE_BOM: 30,
  ANTALL_DELE_REISE: 31,
  BEFARING_DATO: 32,
  BEFARING_KL: 33,
  DATO_STATUSENDRING: 34,
  TIMESTAMP: 35,
  LINK_MAPPE: 36,
  NOTATER: 37,
  PRODUKTNUMMER: 38,
  KOMMENTAR_REGNSKAP: 39,
  KANSELLERT: 40,  // <-- NY KOLONNE PÅ SLUTTEN (AN)
};

const NUM_COLS = 40;
const STATUS_COL_LETTER = 'O'; // Kolonne 15


// ============================================================
// SIKKERHETSSJEKK: Alle e-poster rutes gjennom denne
// ============================================================
// Hent spreadsheet — fungerer fra BÅDE UI og triggere
function getSpreadsheet_() {
  // Prøv aktiv først (fungerer fra UI)
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  // Fallback til ID (fungerer fra triggere)
  return SpreadsheetApp.openById(CONFIG.SHEET_ID);
}


function safeSendEmail_(to, subject, htmlBody, plainBody) {
  let safeRecipient = to;

  if (CONFIG.TEST_MODE) {
    // I testmodus: BARE send til eier og regnskapsfører
    const allowedEmails = [
      CONFIG.OWNER_EMAIL.toLowerCase(),
      CONFIG.ACCOUNTANT_EMAIL.toLowerCase()
    ];

    // Filtrer mottakere
    const recipients = safeRecipient.split(',').map(function (e) { return e.trim().toLowerCase(); });
    const safeRecipients = recipients.filter(function (e) {
      return allowedEmails.indexOf(e) > -1;
    });

    if (safeRecipients.length === 0) {
      // Ingen tillatte mottakere — rut til eier med advarsel
      safeRecipient = CONFIG.OWNER_EMAIL;
      subject = '🧪 [TEST - Ville gått til: ' + to + '] ' + subject;
      Logger.log('⚠️ TEST_MODE: E-post omdirigert fra ' + to + ' til ' + safeRecipient);
    } else {
      safeRecipient = safeRecipients.join(',');
    }
  }

  const emailParams = {
    to: safeRecipient,
    subject: subject
  };

  if (htmlBody) emailParams.htmlBody = htmlBody;
  if (plainBody) emailParams.body = plainBody;

  MailApp.sendEmail(emailParams);
  Logger.log('📧 E-post sendt til: ' + safeRecipient + ' | Emne: ' + subject);
}


// ---- PRISLISTER ----
const PRISLISTE_MED_MARKED = {
  'Leilighet': [
    { maxAreal: 80, pris: 12000 },
    { maxAreal: Infinity, pris: 14000 }
  ],
  'Rekkehus/leilighet 2-4-mannsbolig': [
    { maxAreal: 80, pris: 14000 },
    { maxAreal: Infinity, pris: 16000 }
  ],
  'Enebolig/fritidsbolig': [
    { maxAreal: 150, pris: 18000 },
    { maxAreal: 250, pris: 20000 },
    { maxAreal: Infinity, pris: 22000 }
  ],
  'Frittstående bygg': [
    { maxAreal: Infinity, pris: 1250 }
  ]
};

const PRISLISTE_UTEN_MARKED = {
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


// ============================================================
// 1. INITIAL SETUP
// ============================================================
function initialSetup() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let isStandalone = false;

  if (!ss) {
    isStandalone = true;
    // Prøv å åpne via SHEET_ID først
    if (CONFIG.SHEET_ID && CONFIG.SHEET_ID !== '') {
      try {
        ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
        Logger.log('📄 Åpnet eksisterende regneark via ID: ' + ss.getUrl());
      } catch (e) {
        Logger.log('Kunne ikke åpne SHEET_ID, oppretter nytt...');
      }
    }
    // Hvis fortsatt ingen, opprett nytt
    if (!ss) {
      ss = SpreadsheetApp.create('Naava Takst Oppdragslogg');
      Logger.log('📄 Nytt regneark opprettet: ' + ss.getUrl());
      Logger.log('⚠️ Oppdater CONFIG.SHEET_ID til: ' + ss.getId());
    }
  }

  // --- Oppdragslogg ---
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  } else {
    sheet.clear();
    sheet.clearConditionalFormatRules();
  }

  ['Ark 1', 'Sheet1'].forEach(function (name) {
    const s = ss.getSheetByName(name);
    if (s && s.getLastRow() === 0) {
      try { ss.deleteSheet(s); } catch (e) { }
    }
  });

  const headers = [
    'Kan faktureres', 'Oppdragsnr', 'Dato mottatt', 'Kilde', 'Scan IVIT', 'Oppdragstype', 'Adresse',
    'Oppdragsgiver', 'Selger', 'Selger tlf', 'Selger e-post',
    'Megler/bestiller', 'Megler e-post', 'Faktura ref', 'Status', 'Faktura sendes til',
    'Fakturamotaker', 'Boligtype', 'Areal (m²)', 'Antall tilleggsbygg',
    'Rapporttype', 'Med markedsverdi', 'Timer',
    'Pris inkl. mva', 'Pris eks. mva', 'MVA-beløp',
    'Avstand (km t/r)', 'Reisekostnad eks mva', 'Reisekostnad inkl mva',
    'Sum (ferge/bom)', 'Antall som deler reise',
    'Befaring dato', 'Befaring klokkeslett',
    'Dato statusendring', 'Timestamp (intern)', 'Link til mappe', 'Notater',
    'Produktnummer', 'Kommentar til regnskap', 'Kansellert'
  ];
  const colCount = headers.length; // = NUM_COLS = 35

  sheet.getRange(1, 1, 1, colCount).setValues([headers]);
  sheet.getRange(1, 1, 1, colCount)
    .setFontWeight('bold').setBackground('#1a5c2a')
    .setFontColor('#ffffff').setFontSize(10).setWrap(true);
  sheet.setFrozenRows(1);

  const widths = {};
  widths[COL.OPPDRAGSNR] = 110;
  widths[COL.DATO_MOTTATT] = 130;
  widths[COL.KILDE] = 110;
  widths[COL.OPPDRAGSTYPE] = 180;
  widths[COL.ADRESSE] = 250;
  widths[COL.OPPDRAGSGIVER] = 180;
  widths[COL.SELGER] = 160;
  widths[COL.SELGER_TLF] = 120;
  widths[COL.SELGER_EPOST] = 180;
  widths[COL.MEGLER] = 180;
  widths[COL.MEGLER_EPOST] = 200;
  widths[COL.FAKTURA_REF] = 120;
  widths[COL.FAKTURA_SENDES_TIL] = 180;
  widths[COL.FAKTURAMOTAKER] = 180;
  widths[COL.BOLIGTYPE] = 260;
  widths[COL.AREAL] = 90;
  widths[COL.ANTALL_TILLEGGSBYGG] = 100;
  widths[COL.RAPPORTTYPE] = 300;
  widths[COL.MED_MARKEDSVERDI] = 120;
  widths[COL.TIMER] = 90;
  widths[COL.PRIS_INKL] = 120;
  widths[COL.PRIS_EKS] = 120;
  widths[COL.MVA_BELOP] = 100;
  widths[COL.AVSTAND_KM] = 110;
  widths[COL.REISE_EKS] = 140;
  widths[COL.REISE_INKL] = 140;
  widths[COL.SUM_FERGE_BOM] = 140;
  widths[COL.ANTALL_DELE_REISE] = 100;
  widths[COL.STATUS] = 160;
  widths[COL.BEFARING_DATO] = 130;
  widths[COL.BEFARING_KL] = 120;
  widths[COL.DATO_STATUSENDRING] = 150;
  widths[COL.TIMESTAMP] = 180;
  widths[COL.LINK_MAPPE] = 300;
  widths[COL.NOTATER] = 250;
  Object.entries(widths).forEach(function (e) { sheet.setColumnWidth(parseInt(e[0]), e[1]); });
  sheet.hideColumns(COL.TIMESTAMP);

  // Dropdowns
  setDropdown_(sheet, COL.KILDE, ['Megler-epost', 'IVIT', 'Manuell', 'Privat']);
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
    'Endelig rapport', 'Kan faktureres', 'Fakturert', 'Oppdrag fullført'
  ]);

  // Tallformat: kroner
  [COL.PRIS_INKL, COL.PRIS_EKS, COL.MVA_BELOP, COL.REISE_EKS, COL.REISE_INKL, COL.SUM_FERGE_BOM].forEach(function (c) {
    sheet.getRange(2, c, 500).setNumberFormat('#,##0 "kr"');
  });
  sheet.getRange(2, COL.AVSTAND_KM, 500).setNumberFormat('#,##0 "km"');
  sheet.getRange(2, COL.BEFARING_DATO, 500).setNumberFormat('dd.MM.yyyy');

  // Checkbox for «Med markedsverdi»
  sheet.getRange(2, COL.MED_MARKEDSVERDI, 500).insertCheckboxes();

  setupConditionalFormatting_(sheet, colCount);
  setupDashboard_(ss);
  getOrCreateRootFolder_();
  setupTriggers_(ss);

  // Bekreftelse — bruk UI hvis tilgjengelig, ellers Logger
  const modeLabel = CONFIG.TEST_MODE ? 'TESTMODUS' : 'PRODUKSJON';
  const confirmMsg = 'Naava Takst v3.1 — ' + modeLabel + '\n' +
    colCount + ' kolonner | Sheet ID: ' + ss.getId() + '\n' + ss.getUrl();

  if (!isStandalone) {
    try {
      SpreadsheetApp.getUi().alert('✅ ' + confirmMsg);
    } catch (e) {
      Logger.log('✅ ' + confirmMsg);
    }
  } else {
    Logger.log('✅ ' + confirmMsg);
    // Send bekreftelse per e-post
    MailApp.sendEmail({
      to: CONFIG.OWNER_EMAIL,
      subject: '✅ Naava Takst oppsett fullført',
      body: confirmMsg + '\n\nÅpne regnearket og lim inn scriptet via Utvidelser > Apps Script for meny og onEdit-trigger.'
    });
  }
}



// ============================================================
// 2. E-POST SCANNING — Sikret for test
// ============================================================
function scanIncomingEmails() {
  const label = getOrCreateLabel_('Takst-Behandlet');

  let query;
  if (CONFIG.TEST_MODE) {
    // TESTMODUS: Kun fra test-avsender, etter bestemt dato, ikke allerede behandlet
    const keywords = CONFIG.TRIGGER_KEYWORDS.map(kw => `"${kw}"`).join(' OR ');
    query = 'from:' + CONFIG.TEST_SENDER + ' (' + keywords + ') -label:Takst-Behandlet after:2026/02/05';
    Logger.log('🧪 TESTMODUS: Scanner e-post fra ' + CONFIG.TEST_SENDER + ' etter 06.02.2026');
  } else {
    const keywords = CONFIG.TRIGGER_KEYWORDS.map(function (k) { return '"' + k + '"'; }).join(' OR ');
    const meglerQuery = '(' + keywords + ')';
    const ivitQuery = '(from:no-reply.takst@ivit.no)';
    query = '(' + meglerQuery + ' OR ' + ivitQuery + ') -label:Takst-Behandlet newer_than:2d';
  }
  Logger.log('Gmail query: ' + query);


  let threads;
  try {
    threads = GmailApp.search(query, 0, 20);
  } catch (e) {
    Logger.log('Gmail-søk feilet: ' + e.message);
    return;
  }

  Logger.log('Fant ' + threads.length + ' tråder');

  Logger.log('Starter prosessering av ' + threads.length + ' tråder...');

  for (var ti = 0; ti < threads.length; ti++) {
    var thread = threads[ti];
    try {
      var messages = thread.getMessages();
      var msg = messages[messages.length - 1];

      var subject = msg.getSubject();
      var body = msg.getPlainBody();
      var from = msg.getFrom();
      var date = msg.getDate();

      Logger.log('--- Tråd ' + ti + ': "' + subject + '" fra ' + from);

      // Dobbeltsjekk avsender i testmodus
      if (CONFIG.TEST_MODE) {
        var fromLower = from.toLowerCase();
        if (fromLower.indexOf(CONFIG.TEST_SENDER.toLowerCase()) === -1) {
          Logger.log('  Hopper over — feil avsender');
          continue;
        }
      }

      var isIVIT = (from || '').toLowerCase().indexOf('no-reply.takst@ivit.no') > -1;
      Logger.log('  IVIT: ' + isIVIT);

      // 👇 NY SJEKK: Ignorer IVIT-eposter som inneholder signaturen/navnet ditt 👇
      if (isIVIT && (body || '').indexOf('Jacob Engholm Holen') > -1) {
        Logger.log('  Hopper over IVIT-oppdrag: Inneholder signaturen til Jacob.');
        thread.addLabel(label); // Merker den som behandlet så den ikke sjekkes igjen
        continue; // Avbryter og går til neste e-post
      }

      var parsed = null;

      if (isIVIT) {
        parsed = parseIVITEmail_(subject, body, from);
      } else {
        parsed = assessAndParseWithAI_(subject, body, from);
        if (!parsed) {
          Logger.log('  AI vurderte: ikke relevant — hopper over');
          thread.addLabel(label);
          continue;
        }
        Logger.log('  AI vurderte: relevant — oppdragstype: ' + parsed.oppdragstype);
      }

      if (CONFIG.TEST_MODE) {
        parsed = sanitizeParsedData_(parsed);
      }

      Logger.log('  Adresse: "' + parsed.adresse + '"');
      Logger.log('  Selger: ' + parsed.selger);
      Logger.log('  Kaller createNewOppdrag_...');
      createNewOppdrag_(parsed, date);
      Logger.log('  Retur fra createNewOppdrag_');

      thread.addLabel(label);
      Logger.log('  Ferdig med tråd ' + ti);

    } catch (loopErr) {
      Logger.log('FEIL i tråd ' + ti + ': ' + loopErr.message + '\n' + loopErr.stack);
    }
  }

  Logger.log('Scan ferdig.');
}


// ============================================================
// SANITERING — Fjerner alle eksterne e-poster/data i testmodus
// ============================================================
function sanitizeParsedData_(parsed) {
  const allowedEmails = [
    CONFIG.OWNER_EMAIL.toLowerCase(),
    CONFIG.ACCOUNTANT_EMAIL.toLowerCase()
  ];

  // Erstatt selger e-post (kunde) → jacob (test-kunde)
  if (parsed.selgerEpost && allowedEmails.indexOf(parsed.selgerEpost.toLowerCase()) === -1) {
    parsed.selgerEpost = CONFIG.ACCOUNTANT_EMAIL; // Kunde = jacob i test
  }

  // Erstatt megler e-post (innsender) → jacob (test-innsender)
  if (parsed.meglerEpost && allowedEmails.indexOf(parsed.meglerEpost.toLowerCase()) === -1) {
    parsed.meglerEpost = CONFIG.ACCOUNTANT_EMAIL; // Innsender = jacob i test
  }

  return parsed;
}


// ============================================================
// 3. PARSERE
// ============================================================

function parseMeglerEmail_(subject, body, from) {
  const text = (body || '').replace(/\r/g, '');
  const result = {
    kilde: 'Megler-epost',
    oppdragstype: 'Tilstandsrapport',
    rapporttype: '',
    adresse: '',
    oppdragsgiver: '',
    selger: '',
    selgerTlf: '',
    selgerEpost: '',
    megler: '',
    meglerEpost: '',
    fakturaRef: '',
    fakturaSendesTil: '',
    notater: ''
  };

  // Adresse — prøv flere formater
  let adresseMatch = text.match(/følgende eiendom[:\s]*\n\s*(.+?)(?:\n|,\s*gnr)/i);
  if (adresseMatch) {
    result.adresse = adresseMatch[1].trim();
  }
  if (!result.adresse) {
    adresseMatch = text.match(/[Aa]dress(?:e[n]?\s+er|e[:\s])\s*(.+?)(?:\n|$)/i);
    if (adresseMatch) result.adresse = adresseMatch[1].trim();
  }
  if (!result.adresse) {
    adresseMatch = text.match(/([A-ZÆØÅ][a-zæøåA-ZÆØÅ]*(?:ringen|veien|gata|gaten|vegen|stien|bakken|lia|haugen|åsen|berget|stranda|plassen|torget|brygga|bøen|øen|tunet|marka|jordet|løkka)\s+\d+[A-Za-z]?(?:\s*,?\s*\d{4}\s+[A-ZÆØÅa-zæøå]+)?)/);
    if (adresseMatch) result.adresse = adresseMatch[1].trim();
  }

  const fullMatch = text.match(/((?:[A-ZÆØÅ][\wæøå]+\s+\d+[A-Za-z]?)(?:,\s*gnr[^)]+\))?(?:\s*i\s+[A-ZÆØÅa-zæøå]+)?)/);
  if (fullMatch) result.notater = fullMatch[1].trim();

  const oppdrMatch = text.match(/oppdragsgiver\s+(.+?)(?:,|\n)/i);
  if (oppdrMatch) {
    result.oppdragsgiver = oppdrMatch[1].trim();
    result.selger = result.oppdragsgiver;
  }

  const selgerBlock = text.match(/[Ss]elger treffes på[:\s]*\n([\s\S]*?)(?:\n\s*\n|Vi ber)/);
  if (selgerBlock) {
    const block = selgerBlock[1];
    const tlfMatch = block.match(/(?:Tlf|Telefon|Mob)[.:\s]*(\d[\d\s]+\d)/i);
    if (tlfMatch) result.selgerTlf = tlfMatch[1].replace(/\s/g, '');
    const epostMatch = block.match(/[Ee]-?post[:\s]*([^\s\n]+@[^\s\n]+)/);
    if (epostMatch) result.selgerEpost = epostMatch[1].trim();
  }

  // Fallback: uformelle e-poster — fang tlf, navn (mvh/hilsen), e-post fra body
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
    if (navnMatch) {
      result.oppdragsgiver = navnMatch[1].trim();
      if (!result.selger) result.selger = result.oppdragsgiver;
    }
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
// OPENAI-BASERT VURDERING OG PARSING AV MEGLERMAIL
// Returnerer null hvis e-posten ikke er et takstoppdrag,
// ellers returnerer samme struktur som parseMeglerEmail_.
// ============================================================
function assessAndParseWithAI_(subject, body, from) {
  if (!CONFIG.OPENAI_API_KEY) {
    Logger.log('OpenAI API key ikke satt — faller tilbake til keyword-matching');
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
    '\n- Oversending, vedlegg eller videresending av en ferdig eller tidligere rapport ' +
    '(signalord: "oppdatert", "tilbakemelding", "tidligere", "revidert")' +
    '\n- Avlysning, utsettelse eller omplanlegging av befaring' +
    '\n- Spørsmål om et pågående oppdrag (f.eks. "er taksten oppdatert", "har dere mottatt")' +
    '\n- Videresending av dokumenter til et oppdrag som allerede eksisterer' +
    '\n- Spam, nyhetsbrev, fakturaer, kvitteringer, automatiske varsler' +
    '\n- Intern kommunikasjon eller svar på e-post fra Naava Takst selv' +
    '\n\nSett "relevant": true KUN hvis e-posten er en tydelig ny forespørsel eller bestilling ' +
    'der noen ønsker at Naava Takst skal utføre en befaring eller rapport på en eiendom ' +
    'som det ikke allerede er opprettet oppdrag for. ' +
    'Tvilstilfeller skal vurderes som IKKE relevante (relevant: false). ' +
    'Svar KUN med gyldig JSON. Ingen forklaringstekst utenfor JSON.';

  var userPrompt =
    'E-post mottatt av Naava Takst:\n\n' +
    'Fra: ' + from + '\n' +
    'Emne: ' + subject + '\n\n' +
    'Innhold:\n' + (body || '').substring(0, 3000) + '\n\n' +
    'Spør deg selv: Er dette en NY bestilling av et takstoppdrag, eller handler det om noe som allerede er i gang?\n\n' +
    'Svar med dette JSON-skjemaet:\n' +
    '{\n' +
    '  "relevant": true,\n' +
    '  "begrunnelse": "<én setning om hvorfor dette er en ny bestilling>",\n' +
    '  "oppdragstype": "<én av: Tilstandsrapport | Tilstandsrapport m/markedsverdi | Skadetakst | Reklamasjon | Vurderingsoppdrag | Bistand overtagelse | Fukt-/fuktskadevurdering | Byggelånskontroll | Forhåndstakst | Verditakst | Annet>",\n' +
    '  "rapporttype": "<én av: Tilstandsrapport m/teknisk og markedsverdi | Tilstandsrapport | Skadetakstrapport | Reklamasjonsrapport | Vurderingsrapport | Overtagelsesrapport | Annen rapport>",\n' +
    '  "adresse": "<gateadresse, postnr, poststed>",\n' +
    '  "oppdragsgiver": "<firmanavn eller personnavn til bestiller>",\n' +
    '  "selger": "<selgers fulle navn>",\n' +
    '  "selgerTlf": "<selgers telefonnummer, kun sifre>",\n' +
    '  "selgerEpost": "<selgers e-postadresse>",\n' +
    '  "megler": "<meglerens navn>",\n' +
    '  "meglerEpost": "<meglerens e-postadresse>",\n' +
    '  "fakturaRef": "<referansenummer for faktura>",\n' +
    '  "fakturaSendesTil": "<hvem faktura skal sendes til>",\n' +
    '  "notater": "<andre relevante opplysninger, gnr/bnr, seksjoner, spesielle instrukser>"\n' +
    '}\n' +
    'Feltene kan være tomme strenger ("") hvis informasjonen ikke finnes. ' +
    'Ikke gjett — fyll kun ut det som faktisk står i e-posten.';

  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + CONFIG.OPENAI_API_KEY
      },
      payload: JSON.stringify({
        model: CONFIG.OPENAI_MODEL,
        temperature: 0,
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userPrompt }
        ]
      }),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      Logger.log('OpenAI vurdering feilet (HTTP ' + response.getResponseCode() + ') — faller tilbake til parseMeglerEmail_');
      return parseMeglerEmail_(subject, body, from);
    }

    var data = JSON.parse(response.getContentText());
    var text = data.choices[0].message.content.trim();

    // Strip eventuelle ```json ``` wrapper
    var jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      Logger.log('OpenAI svarte ikke med JSON — faller tilbake');
      return parseMeglerEmail_(subject, body, from);
    }

    var result = JSON.parse(jsonMatch[0]);

    if (!result.relevant) {
      return null;
    }

    return {
      kilde: 'Megler-epost',
      oppdragstype: result.oppdragstype || 'Annet',
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
    Logger.log('assessAndParseWithAI_ feil: ' + e.message + ' — faller tilbake til parseMeglerEmail_');
    return parseMeglerEmail_(subject, body, from);
  }
}

function parseIVITEmail_(subject, body, from) {
  const text = (body || '').replace(/\r/g, '');
  const result = {
    kilde: 'IVIT',
    oppdragstype: 'Tilstandsrapport',
    adresse: '',
    oppdragsgiver: '',
    selger: '',
    selgerTlf: '',
    selgerEpost: '',
    megler: '',
    meglerEpost: '',
    fakturaRef: '',
    fakturaSendesTil: '',
    notater: ''
  };

  const ordreMatch = text.match(/(?:^|\n)\s*[Oo]rdre\s*nummer\s*[:\t ]?\s*([A-Z0-9-]+)\s*(?:\n|$)/);
  if (ordreMatch) result.fakturaRef = ordreMatch[1].trim();

  // Adresse — prøv flere formater
  let adresseMatch = text.match(/følgende eiendom[:\s]*\n\s*(.+?)(?:\n|,\s*gnr)/i);
  if (adresseMatch) {
    result.adresse = adresseMatch[1].trim();
  }
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


function detectOppdragstype_(text) {
  const t = text.toLowerCase();
  if (t.indexOf('reklamasjon') > -1) return 'Reklamasjon';
  if (t.indexOf('skade') > -1 || t.indexOf('skadetakst') > -1) return 'Skadetakst';
  if (t.indexOf('vurderingsoppdrag') > -1 || t.indexOf('verditakst') > -1) return 'Vurderingsoppdrag';
  if (t.indexOf('overtagelse') > -1 || t.indexOf('overtakelse') > -1) return 'Bistand overtagelse';
  if (t.indexOf('fukt') > -1) return 'Fukt-/fuktskadevurdering';
  if (t.indexOf('byggelån') > -1) return 'Byggelånskontroll';
  if (t.indexOf('forhåndstakst') > -1) return 'Forhåndstakst';
  if (t.indexOf('markedsverdi') > -1) return 'Tilstandsrapport m/markedsverdi';
  if (t.indexOf('tilstandsrapport') > -1) return 'Tilstandsrapport';
  return 'Annet';
}


// ============================================================
// 4. REISEBEREGNING VIA OPENAI API
// ============================================================
function calculateDistance_(destinationAddress) {
  if (!CONFIG.OPENAI_API_KEY) {
    Logger.log('OpenAI API key ikke satt — hopper over reiseberegning');
    return null;
  }

  const prompt =
    'Jeg trenger en estimert kjøreavstand (én vei) mellom disse to adressene i Norge.\n\n' +
    'Fra: ' + CONFIG.BASE_ADDRESS + '\n' +
    'Til: ' + destinationAddress + '\n\n' +
    'Svar KUN med et JSON-objekt i dette formatet, ingen annen tekst:\n' +
    '{"km_en_vei": <tall>, "km_tur_retur": <tall>, "estimert_tid_min": <tall>}\n\n' +
    'Gi realistisk kjøreavstand (ikke luftlinje).';

  try {
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + CONFIG.OPENAI_API_KEY
      },
      payload: JSON.stringify({
        model: CONFIG.OPENAI_MODEL,
        temperature: 0,
        messages: [
          { role: 'system', content: 'Du svarer kun med gyldig JSON.' },
          { role: 'user', content: prompt }
        ]
      }),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      Logger.log('OpenAI API feil (HTTP ' + response.getResponseCode() + ')');
      return null;
    }

    const data = JSON.parse(response.getContentText());
    const text = data.choices[0].message.content.trim();
    const jsonMatch = text.match(/\{[\s\S]*\}/);

    if (!jsonMatch) return null;

    const result = JSON.parse(jsonMatch[0]);

    return {
      kmEnVei: Number(result.km_en_vei) || 0,
      kmTurRetur: Number(result.km_tur_retur) || 0,
      estimertTidMin: Number(result.estimert_tid_min) || 0
    };

  } catch (e) {
    Logger.log('OpenAI API feil: ' + e.message);
    return null;
  }
}



function calculateTravelCost_(kmTurRetur, sumFergeBomInklMva = 0) {
  const fakturerbarKm = Math.max(0, kmTurRetur - CONFIG.REISE_INKLUDERT_KM);

  // 1. Regn ut kostnad for selve kjøringen
  const kmKostnadEks = fakturerbarKm * CONFIG.REISE_SATS_EKS_MVA;
  const kmKostnadInkl = kmKostnadEks * (1 + CONFIG.MVA_RATE);

  // 2. Håndter ferge/bom (som allerede er inkl mva)
  const bomInkl = sumFergeBomInklMva;
  const bomEks = bomInkl / (1 + CONFIG.MVA_RATE);

  // 3. Legg dem sammen
  return {
    totalKm: kmTurRetur,
    fakturerbarKm: fakturerbarKm,
    kostnadEksMva: Math.round(kmKostnadEks + bomEks),
    kostnadInklMva: Math.round(kmKostnadInkl + bomInkl)
  };
}


// ============================================================
// 5. OPPRETT NYTT OPPDRAG
// ============================================================
function createNewOppdrag_(parsed, date, skipOrdreGate) {
  try {
    Logger.log('>>> createNewOppdrag_ start: ' + parsed.adresse);

    const ss = getSpreadsheet_();
    Logger.log('  Sheet ID: ' + ss.getId());

    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      Logger.log('  FEIL: Fant ikke ark "' + CONFIG.SHEET_NAME + '"');
      return;
    }
    Logger.log('  Ark funnet: ' + sheet.getName() + ', rader: ' + sheet.getLastRow());

    // Duplikatsjekk
    let monthDuplicateWarning = '';
    if (sheet.getLastRow() > 1) {
      const existingData = sheet.getDataRange().getValues();
      const incomingAddrBase = parsed.adresse.split(',')[0].trim().toLowerCase();

      for (let i = 1; i < existingData.length; i++) {
        const rowAddrRaw = String(existingData[i][COL.ADRESSE - 1] || '');
        const rowTimestamp = existingData[i][COL.TIMESTAMP - 1];
        if (!rowTimestamp) continue;

        const rowDate = new Date(rowTimestamp);

        // Eksakt samme dag: Avbryt
        if (rowAddrRaw === parsed.adresse && rowDate.toDateString() === date.toDateString()) {
          Logger.log('  Eksakt duplikat samme dag, avbryter: ' + parsed.adresse);
          return;
        }

        // Samme måned: Varsle
        let rowAddrClean = rowAddrRaw;
        const hyperMatch = rowAddrRaw.match(/=HYPERLINK\("[^"]+","([^"]+)"\)/);
        if (hyperMatch) rowAddrClean = hyperMatch[1].replace(/""/g, '"');
        const rowAddrBase = rowAddrClean.split(',')[0].trim().toLowerCase();

        if (rowAddrBase === incomingAddrBase && rowDate.getMonth() === date.getMonth() && rowDate.getFullYear() === date.getFullYear()) {
          monthDuplicateWarning = '⚠️ Samme adresse reg. tidligere denne mnd!';
        }
      }
    }

    const lastRow = sheet.getLastRow();
    const oppdragsnr = 'NT-' + Utilities.formatDate(date, 'Europe/Oslo', 'yyyyMM') + '-' + String(lastRow).padStart(3, '0');
    Logger.log('  Oppdragsnr: ' + oppdragsnr);

    let folderUrl = '';
    try {
      const folderName = (parsed.adresse || 'Ukjent') + ' - ' + Utilities.formatDate(date, 'Europe/Oslo', 'dd.MM.yyyy');
      const folder = createOppdragFolder_(folderName);
      folderUrl = folder.getUrl();
      Logger.log('  Mappe opprettet: ' + folderUrl);
    } catch (folderErr) {
      Logger.log('  ADVARSEL: Kunne ikke opprette mappe: ' + folderErr.message);
    }

    const datoMottatt = Utilities.formatDate(date, 'Europe/Oslo', 'dd.MM.yyyy HH:mm');
    const timestamp = date.toISOString();

    // Reiseberegning (hopper over hvis ingen API-nøkkel)
    let avstandKm = '';
    let reiseEks = 0;
    let reiseInkl = 0;
    let reiseNotat = '';

    if (parsed.adresse) {
      const distResult = calculateDistance_(parsed.adresse);
      if (distResult) {
        const travelCost = calculateTravelCost_(distResult.kmTurRetur);
        avstandKm = distResult.kmTurRetur;
        reiseEks = travelCost.kostnadEksMva;
        reiseInkl = travelCost.kostnadInklMva;
        reiseNotat = distResult.kmTurRetur + ' km t/r' +
          (travelCost.fakturerbarKm > 0 ?
            ', ' + travelCost.fakturerbarKm + ' km fakturerbart' :
            ' (inkludert)');
      }
    }

    let fullNotater = '';
    if (parsed.notater) fullNotater += parsed.notater;
    if (reiseNotat) fullNotater += (fullNotater ? ' | ' : '') + 'Reise: ' + reiseNotat;

    // Bygg rad i ny kolonnerekkefølge (NUM_COLS = 35 elementer)
    const newRow = [];
    newRow[COL.OPPDRAGSNR - 1] = oppdragsnr;
    newRow[COL.DATO_MOTTATT - 1] = datoMottatt;
    newRow[COL.KILDE - 1] = parsed.kilde;
    newRow[COL.OPPDRAGSTYPE - 1] = parsed.oppdragstype;
    newRow[COL.ADRESSE - 1] = folderUrl
      ? '=HYPERLINK("' + folderUrl + '","' + parsed.adresse.replace(/"/g, '""') + '")'
      : parsed.adresse;
    newRow[COL.OPPDRAGSGIVER - 1] = parsed.oppdragsgiver;
    newRow[COL.SELGER - 1] = parsed.selger;
    newRow[COL.SELGER_TLF - 1] = parsed.selgerTlf;
    newRow[COL.SELGER_EPOST - 1] = parsed.selgerEpost;
    newRow[COL.MEGLER - 1] = parsed.megler;
    newRow[COL.MEGLER_EPOST - 1] = parsed.meglerEpost;
    newRow[COL.FAKTURA_REF - 1] = parsed.fakturaRef;
    newRow[COL.FAKTURA_SENDES_TIL - 1] = parsed.fakturaSendesTil;
    newRow[COL.FAKTURAMOTAKER - 1] = '';
    newRow[COL.BOLIGTYPE - 1] = '';
    newRow[COL.AREAL - 1] = '';
    newRow[COL.ANTALL_TILLEGGSBYGG - 1] = '';
    newRow[COL.RAPPORTTYPE - 1] = parsed.rapporttype;
    newRow[COL.MED_MARKEDSVERDI - 1] = (parsed.rapporttype === 'Tilstandsrapport m/teknisk og markedsverdi');
    newRow[COL.TIMER - 1] = '';
    newRow[COL.PRIS_INKL - 1] = '';
    newRow[COL.PRIS_EKS - 1] = '';
    newRow[COL.MVA_BELOP - 1] = '';
    newRow[COL.AVSTAND_KM - 1] = avstandKm;
    newRow[COL.REISE_EKS - 1] = reiseEks;
    newRow[COL.REISE_INKL - 1] = reiseInkl;
    newRow[COL.SUM_FERGE_BOM - 1] = '';
    newRow[COL.ANTALL_DELE_REISE - 1] = '';
    // Nye oppdrag havner i 'Ordre' (venter på godkjenning) og ikke 'Mottatt' —
    // Jacob godkjenner/avviser dem i Ordre-fanen før de blir ekte oppdrag.
    // Manuell registrering hopper over gaten siden han allerede har lest over det han skriver inn.
    newRow[COL.STATUS - 1] = skipOrdreGate ? 'Mottatt' : 'Ordre';
    newRow[COL.BEFARING_DATO - 1] = '';
    newRow[COL.BEFARING_KL - 1] = '';
    newRow[COL.DATO_STATUSENDRING - 1] = datoMottatt;
    newRow[COL.TIMESTAMP - 1] = timestamp;
    newRow[COL.LINK_MAPPE - 1] = folderUrl;
    newRow[COL.NOTATER - 1] = fullNotater;
    //newRow[COL.SCAN_IVIT - 1] = (parsed.kilde === 'IVIT'); //Kommentert ut fordi vi vil ha kontroll over IVITscan

    Logger.log('  Skriver rad med ' + newRow.length + ' kolonner...');
    sheet.appendRow(newRow);
    const newRowNum = sheet.getLastRow();
    Logger.log('  ✅ Rad skrevet! Ny lastRow: ' + newRowNum);
    //sheet.getRange(newRowNum, COL.SCAN_IVIT).insertCheckboxes();    //Kommentert ut fordi vi vil ha kontroll over IVITscan
    sheet.getRange(newRowNum, COL.MED_MARKEDSVERDI).insertCheckboxes();
    sheet.getRange(newRowNum, COL.MED_MARKEDSVERDI).setValue(
      parsed.rapporttype === 'Tilstandsrapport m/teknisk og markedsverdi'
    );

    // Alerts
    let alertMsg = '📋 *Nytt oppdrag!* ' + parsed.oppdragstype + '\n' +
      '🏠 ' + parsed.adresse + '\n';
    if (parsed.megler) alertMsg += '🏢 ' + parsed.megler + '\n';
    if (parsed.fakturaRef) alertMsg += '🔖 Ref: ' + parsed.fakturaRef + '\n';
    if (avstandKm) alertMsg += '🚗 ' + avstandKm + ' km t/r';
    if (reiseEks > 0) alertMsg += ' → ' + formatCurrency_(reiseEks) + ' eks mva';
    //sendChatAlert_(alertMsg);

    //safeSendEmail_(
    //CONFIG.OWNER_EMAIL,
    //'📋 Nytt takstoppdrag: ' + parsed.adresse + ' (' + parsed.oppdragstype + ')',
    //buildNewOppdragEmail_(parsed, datoMottatt, folderUrl, oppdragsnr, avstandKm, reiseEks, reiseInkl)
    //);

    Logger.log('>>> createNewOppdrag_ ferdig OK');

  } catch (err) {
    Logger.log('❌ FEIL i createNewOppdrag_: ' + err.message);
    Logger.log('  Stack: ' + err.stack);
  }
}


// ============================================================
// 6. ONEDIT
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

    // Oppdaterer pris ved endring av relevante felt
    if (col === COL.BOLIGTYPE || col === COL.AREAL || col === COL.RAPPORTTYPE || col === COL.MED_MARKEDSVERDI || col === COL.ANTALL_TILLEGGSBYGG || col === COL.TIMER) {
      calculatePrice_(sheet, row);
    }

    // Hvis pris inkl mva endres manuelt, synk eks og mva
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

    // Adresse endret -> rekalk reise fra adresse og sjekk duplikat
    if (col === COL.ADRESSE && e.value) {
      recalculateTravel_(sheet, row, e.value);

      let rowDate = sheet.getRange(row, COL.TIMESTAMP).getValue();
      if (!rowDate || !(rowDate instanceof Date)) rowDate = new Date();
      flagDuplicateAddress_(sheet, row, e.value, rowDate);
    }

    // NYTT: Reise: avstand km t/r ELLER ferge/bom endret
    if (col === COL.AVSTAND_KM || col === COL.SUM_FERGE_BOM) {
      const kmStr = String(sheet.getRange(row, COL.AVSTAND_KM).getValue() || '0').replace(',', '.');
      const bomStr = String(sheet.getRange(row, COL.SUM_FERGE_BOM).getValue() || '0').replace(/[^\d.,-]/g, '').replace(',', '.');

      const km = parseFloat(kmStr);
      const bom = parseFloat(bomStr) || 0;

      if (!isNaN(km) && km >= 0) {
        const tc = calculateTravelCost_(km, bom);
        sheet.getRange(row, COL.REISE_EKS).setValue(tc.kostnadEksMva);
        sheet.getRange(row, COL.REISE_INKL).setValue(tc.kostnadInklMva);
      }
    }

    // Reise: eks mva endret -> synk inkl mva
    if (col === COL.REISE_EKS) {
      const eks = parseFloat(String(sheet.getRange(row, COL.REISE_EKS).getValue()).replace(',', '.'));
      if (!isNaN(eks) && eks >= 0) {
        sheet.getRange(row, COL.REISE_INKL).setValue(Math.round(eks * (1 + CONFIG.MVA_RATE)));
      }
    }

    // Reise: inkl mva endret -> synk eks mva
    if (col === COL.REISE_INKL) {
      const inkl = parseFloat(String(sheet.getRange(row, COL.REISE_INKL).getValue()).replace(',', '.'));
      if (!isNaN(inkl) && inkl >= 0) {
        sheet.getRange(row, COL.REISE_EKS).setValue(Math.round(inkl / (1 + CONFIG.MVA_RATE)));
      }
    }

    // Dashboard refresh for relevante felt (lagt til tilleggsbygg og bom her også)
    if ([COL.OPPDRAGSTYPE, COL.BOLIGTYPE, COL.AREAL, COL.RAPPORTTYPE,
    COL.PRIS_INKL, COL.AVSTAND_KM, COL.REISE_EKS, COL.REISE_INKL,
    COL.STATUS, COL.ANTALL_TILLEGGSBYGG, COL.SUM_FERGE_BOM].indexOf(col) > -1) {
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

  // HVIS VALGT FRA DROPDOWN:
  if (newStatus === 'Kan faktureres') {
    safeSendEmail_(
      CONFIG.ACCOUNTANT_EMAIL,
      '💰 Klar til fakturering: ' + adresse + ' (' + oppdragsnr + ')',
      buildFakturaEmail_(rowData, datoStr)
    );

    // Auto-hopp til Fakturert og fjern evt. checkbox-kryss
    sheet.getRange(row, COL.STATUS).setValue('Fakturert');
    sheet.getRange(row, COL.KAN_FAKTURERES).setValue(false);
    archiveRow_(sheet, row);
  }

  // Fullført oppdrag: Flytter Google Drive-mappen
  if (newStatus === 'Oppdrag fullført') {
    const folderUrlOrId = rowData[COL.LINK_MAPPE - 1];
    const folderId = extractDriveFolderId_(folderUrlOrId);
    if (folderId) {
      try {
        moveOppdragFolderToAvsluttede_(folderId);
      } catch (e) {
        Logger.log('Flytt mappe feilet: ' + e);
      }
    }
  }
}


// ============================================================
// 7. BEFARINGSBEKREFTELSE — Sendes via safeSendEmail_
// ============================================================
function handleBefaringBooked_(sheet, row, befaringDato, befaringTid) {
  const rowData = sheet.getRange(row, 1, 1, NUM_COLS).getValues()[0];
  const adresse = rowData[COL.ADRESSE - 1];
  const selger = rowData[COL.SELGER - 1];
  const selgerEpost = rowData[COL.SELGER_EPOST - 1];
  const meglerEpost = rowData[COL.MEGLER_EPOST - 1];
  const oppdragstype = rowData[COL.OPPDRAGSTYPE - 1];

  let datoFormatert;
  if (befaringDato instanceof Date) {
    datoFormatert = Utilities.formatDate(befaringDato, 'Europe/Oslo', 'dd.MM.yyyy');
  } else {
    datoFormatert = String(befaringDato);
  }
  const tidStr = befaringTid instanceof Date
    ? Utilities.formatDate(befaringTid, 'Europe/Oslo', 'HH:mm')
    : (befaringTid ? String(befaringTid) : 'Ikke angitt');

  // Auto-sett status til "Avtalt befaring"
  const currentStatus = rowData[COL.STATUS - 1];
  if (currentStatus === 'Mottatt') {
    sheet.getRange(row, COL.STATUS).setValue('Avtalt befaring');
    sheet.getRange(row, COL.DATO_STATUSENDRING).setValue(Utilities.formatDate(new Date(), 'Europe/Oslo', 'dd.MM.yyyy HH:mm'));
  }

  // Mottaker: selger eller megler — men ALLTID via safeSendEmail_ som sikrer testmodus
  let recipient = selgerEpost || meglerEpost || CONFIG.OWNER_EMAIL;

  // safeSendEmail_ håndterer testmodus-sjekken
  //safeSendEmail_(
  //  recipient,
  //  '✅ Befaring avtalt — ' + adresse + ' (' + datoFormatert + ')',
  //  buildBefaringEmail_(adresse, datoFormatert, tidStr, selger, oppdragstype)
  //);

  //sendChatAlert_(
  //  '📅 *Befaring avtalt*\n🏠 ' + adresse +
  //  '\n📅 ' + datoFormatert + ' kl ' + tidStr
  //);
}


function buildBefaringEmail_(adresse, dato, tid, selgerNavn, oppdragstype) {
  const kundeNavn = selgerNavn || 'kunde';

  return '<div style="font-family:Arial,sans-serif;max-width:600px;">' +
    '<div style="background:#1a5c2a;color:white;padding:20px;border-radius:8px 8px 0 0;text-align:center;">' +
    '<h2 style="margin:0;">Bekreftelse på befaring</h2></div>' +
    '<div style="padding:24px;border:1px solid #ddd;border-top:none;">' +
    '<p>Hei ' + kundeNavn + ',</p>' +
    '<p>Vi bekrefter herved avtalt befaring på følgende eiendom:</p>' +
    '<table style="width:100%;border-collapse:collapse;margin:16px 0;">' +
    '<tr style="background:#f5f5f5;"><td style="padding:10px;font-weight:bold;">Adresse:</td><td style="padding:10px;">' + adresse + '</td></tr>' +
    '<tr><td style="padding:10px;font-weight:bold;">Type oppdrag:</td><td style="padding:10px;">' + oppdragstype + '</td></tr>' +
    '<tr style="background:#f5f5f5;"><td style="padding:10px;font-weight:bold;">Dato:</td><td style="padding:10px;font-size:16px;font-weight:bold;">' + dato + '</td></tr>' +
    '<tr><td style="padding:10px;font-weight:bold;">Klokkeslett:</td><td style="padding:10px;font-size:16px;font-weight:bold;">' + tid + '</td></tr>' +
    '</table>' +

    '<div style="background:#f8f9fa;border:1px solid #e0e0e0;border-radius:6px;padding:16px;margin:16px 0;">' +
    '<h3 style="margin:0 0 12px;color:#1a5c2a;">📋 Sjekkliste før befaring</h3>' +
    '<p style="font-weight:bold;margin:12px 0 4px;">Rydd fri adkomst:</p>' +
    '<ul style="margin:0;padding-left:20px;color:#333;">' +
    '<li>Alle rom, boder, garasje, teknisk rom</li>' +
    '<li>Klargjør loft, krypkjeller, takluke, legg frem stige ved behov</li>' +
    '<li>Sørg for tilgang til sikringsskap, hovedkran, vannstoppeventil, stoppekran i våtrom, varmtvannsbereder</li>' +
    '<li>Sørg for tilgang til sluk, rørskap, synlige rør, under servanter</li></ul>' +
    '<p style="font-weight:bold;margin:12px 0 4px;">Noter kjente forhold:</p>' +
    '<ul style="margin:0;padding-left:20px;color:#333;"><li>Fukt, lekkasjer, sprekker, setninger, lukt, støy, tidligere skader</li></ul>' +
    '<p style="font-weight:bold;margin:12px 0 4px;">Finn frem dokumentasjon:</p>' +
    '<ul style="margin:0;padding-left:20px;color:#333;"><li>Dokumentasjon som er lagret digitalt kan sendes på epost</li></ul>' +
    '<p style="font-weight:bold;margin:12px 0 4px;">Oppgraderinger / oppussing:</p>' +
    '<ul style="margin:0;padding-left:20px;color:#333;"><li>Skriv en punktliste med årstall over hva som er gjort og når</li></ul>' +
    '<p style="font-weight:bold;margin:12px 0 4px;">Hulltaking:</p>' +
    '<ul style="margin:0;padding-left:20px;color:#333;">' +
    '<li>Iht. Avhendingslova (NS 3600) skal det utføres hulltaking i alle våtrom og i rom under terreng</li>' +
    '<li>Etter hulltaking monteres plastlokk</li></ul>' +
    '<div style="background:#fff3e0;border:1px solid #ffcc80;border-radius:4px;padding:12px;margin-top:12px;">' +
    '<p style="margin:0;font-weight:bold;">📌 Boligmappa</p>' +
    '<p style="margin:4px 0 0;">Logg inn på boligmappa.no → Velg riktig bolig → "Gi tilgang" → Legg inn vår e-post → Lesetilgang 1 måned</p></div>' +
    '</div>' +

    '<p style="margin-top:16px;color:#666;">Har du spørsmål? Ta gjerne kontakt.</p></div>' +

    '<div style="background:#f5f5f5;padding:16px;border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;font-size:13px;color:#666;">' +
    '<strong>Jacob Engholm Holen</strong><br>Takstingeniør<br>+47 469 49 615<br>jacob@naava.no<br>www.naava.no<br><br>' +
    '<em>Medlem av Norsk Takst og NITO</em></div></div>';
}


// ============================================================
// 8. REISE-REBEREGNING
// ============================================================
function recalculateTravel_(sheet, row, address) {
  const distResult = calculateDistance_(address);
  if (distResult) {
    // NYTT: Hent eksisterende ferge/bom slik at det ikke slettes
    const bomStr = String(sheet.getRange(row, COL.SUM_FERGE_BOM).getValue() || '0').replace(/[^\d.,-]/g, '').replace(',', '.');
    const bom = parseFloat(bomStr) || 0;

    const travelCost = calculateTravelCost_(distResult.kmTurRetur, bom);
    sheet.getRange(row, COL.AVSTAND_KM).setValue(distResult.kmTurRetur);
    sheet.getRange(row, COL.REISE_EKS).setValue(travelCost.kostnadEksMva);
    sheet.getRange(row, COL.REISE_INKL).setValue(travelCost.kostnadInklMva);
  }
}



// ============================================================
// 9. PRISREBEREGNING
// ============================================================
function calculatePrice_(sheet, row) {
  const boligtype = sheet.getRange(row, COL.BOLIGTYPE).getValue();
  const arealStr = sheet.getRange(row, COL.AREAL).getValue();
  const tilleggStr = sheet.getRange(row, COL.ANTALL_TILLEGGSBYGG).getValue();
  const inkluderMarked = sheet.getRange(row, COL.MED_MARKEDSVERDI).getValue() === true;
  const timerStr = sheet.getRange(row, COL.TIMER).getValue();

  if (!boligtype && !timerStr) return;

  const areal = arealStr ? parseFloat(arealStr) : 0;
  const tilleggsbygg = tilleggStr ? parseInt(tilleggStr, 10) : 0;
  const timer = timerStr ? parseFloat(String(timerStr).replace(',', '.')) : 0;

  const prisliste = getPricesFromSheet_();
  if (!prisliste) return;

  let pris = 0;
  let valgtProdNr = '';

  if (timer > 0 && !boligtype) {
    pris = timer * prisliste.timesats.pris;
    valgtProdNr = prisliste.timesats.prodNr;
  }
  else if (boligtype && prisliste[boligtype]) {
    const alternativer = prisliste[boligtype];
    if (boligtype === 'Frittstående bygg') {
      pris = inkluderMarked ? alternativer[0].prisMarked : alternativer[0].prisStandard;
      valgtProdNr = alternativer[0].prodNr;
    } else {
      for (let i = 0; i < alternativer.length; i++) {
        if (areal <= alternativer[i].maxAreal || alternativer[i].maxAreal === Infinity) {
          pris = inkluderMarked ? alternativer[i].prisMarked : alternativer[i].prisStandard;
          valgtProdNr = alternativer[i].prodNr;
          break;
        }
      }
    }
  }

  // Legg til Markedsverdi-produktnummeret i strengen hvis aktivert
  if (inkluderMarked && boligtype !== 'Frittstående bygg') {
    const markedsNr = prisliste.markedsverdi ? prisliste.markedsverdi.prodNr : '9';
    valgtProdNr = valgtProdNr ? valgtProdNr + ', ' + markedsNr : markedsNr;
  }

  if (tilleggsbygg > 0) pris += (tilleggsbygg * 1250);

  if (pris > 0) {
    const prisEks = Math.round(pris / (1 + CONFIG.MVA_RATE));
    sheet.getRange(row, COL.PRIS_INKL).setValue(pris);
    sheet.getRange(row, COL.PRIS_EKS).setValue(prisEks);
    sheet.getRange(row, COL.MVA_BELOP).setValue(pris - prisEks);
    sheet.getRange(row, COL.PRODUKTNUMMER).setValue(valgtProdNr);
  } else {
    sheet.getRange(row, COL.PRIS_INKL).setValue('');
    sheet.getRange(row, COL.PRIS_EKS).setValue('');
    sheet.getRange(row, COL.MVA_BELOP).setValue('');
    sheet.getRange(row, COL.PRODUKTNUMMER).setValue('');
  }
}


// ============================================================
// 10. PÅMINNELSER
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

    const mottattDato = new Date(timestamp);
    const timerSiden = (now - mottattDato) / (1000 * 60 * 60);

    //if (timerSiden >= CONFIG.REMINDER_HOURS && timerSiden < CONFIG.REMINDER_HOURS + 1.5) {
    //sendChatAlert_('⏰ *Påminnelse:* ' + adresse + ' (' + oppdragsnr + ') — ' + Math.round(timerSiden) + ' timer');
    //}
    //if (timerSiden >= CONFIG.URGENT_HOURS && timerSiden < CONFIG.URGENT_HOURS + 1.5) {
    //sendChatAlert_('🚨 *HASTER:* ' + adresse + ' (' + oppdragsnr + ') — ' + Math.round(timerSiden) + ' timer!');
    //}
  }
}


// ============================================================
// 11. UKENTLIG RAPPORT
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
    const mvaBeløp = data[i][COL.MVA_BELOP - 1] || 0;
    const reiseEks = data[i][COL.REISE_EKS - 1] || 0;
    const reiseInkl = data[i][COL.REISE_INKL - 1] || 0;
    const statusDatoStr = data[i][COL.DATO_STATUSENDRING - 1];

    let sDato = parseDateString_(statusDatoStr);
    if (!sDato) continue;

    const item = {
      oppdragsnr: data[i][COL.OPPDRAGSNR - 1],
      adresse: data[i][COL.ADRESSE - 1],
      oppdragstype: data[i][COL.OPPDRAGSTYPE - 1],
      boligtype: data[i][COL.BOLIGTYPE - 1],
      prisInkl: prisInkl, prisEks: prisEks,
      mvaBeløp: mvaBeløp, reiseEks: reiseEks, reiseInkl: reiseInkl,
      fakturaRef: data[i][COL.FAKTURA_REF - 1],
      fakturaTil: data[i][COL.FAKTURA_SENDES_TIL - 1],
      statusDato: statusDatoStr
    };

    if (status === 'Kan faktureres' && sDato >= monday) {
      fakturerbare.push(item);
      totalInklMva += prisInkl; totalEksMva += prisEks; totalMva += mvaBeløp;
      totalReiseEks += reiseEks; totalReiseInkl += reiseInkl;
    }
    if (status === 'Fakturert' && sDato >= monday) fakturerte.push(item);
    if (status === 'Kan faktureres') ventende.push(item);
  }

  const html = buildWeeklyReportHtml_(fakturerbare, fakturerte, ventende, {
    totalInklMva: totalInklMva, totalEksMva: totalEksMva, totalMva: totalMva,
    totalReiseEks: totalReiseEks, totalReiseInkl: totalReiseInkl
  });

  safeSendEmail_(
    CONFIG.OWNER_EMAIL + ',' + CONFIG.ACCOUNTANT_EMAIL,
    '📊 Naava Takst — Ukerapport uke ' + getWeekNumber_(now),
    html
  );

  sendChatAlert_(
    '📊 *Ukerapport*\n💰 ' + formatCurrency_(totalInklMva + totalReiseInkl) + ' totalt inkl mva'
  );
}


// ============================================================
// 12. DASHBOARD
// ============================================================
function updateDashboard() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  let dash = ss.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);
  if (!sheet || !dash) return;

  if (sheet.getLastRow() < 2) {
    dash.clear();
    dash.getRange(1, 1).setValue('Ingen oppdrag ennå.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const curMonth = now.getMonth(), curYear = now.getFullYear();

  const statusList = ['Mottatt', 'Avtalt befaring', 'Befart', 'Utkast', 'Endelig rapport', 'Kan faktureres', 'Fakturert', 'Oppdrag kansellert', 'Oppdrag fullført'];
  const stats = {
    total: data.length - 1,
    byStatus: {},
    byType: {},
    byBolig: {},
    byBoligBucket: {},

    omsMåned: 0,
    omsÅr: 0,
    reiseMåned: 0,
    reiseÅr: 0,

    uteståendeInkl: 0,
    uteståendeEks: 0,

    snittPrisInkl: 0,
    snittPrisEks: 0,
    snittAntall: 0,

    trendMonths: {} // key: YYYY-MM -> { omsInkl, reiseInkl, totalInkl }
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

  let sumPrisInkl = 0;
  let sumPrisEks = 0;
  let countPris = 0;

  // Ordre/Ordre avvist = venter på godkjenning, ikke ekte oppdrag ennå.
  const ORDRE_STATUSES_ = ['Ordre', 'Ordre avvist'];
  let ordreStageCount = 0;

  for (let i = 1; i < data.length; i++) {
    const status = data[i][COL.STATUS - 1];
    if (ORDRE_STATUSES_.indexOf(status) !== -1) { ordreStageCount++; continue; }
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
      stats.omsÅr += prisInkl;
      stats.reiseÅr += reiseInkl;

      const statusDatoStr = data[i][COL.DATO_STATUSENDRING - 1];
      const statusDato = parseDateString_(statusDatoStr);
      if (statusDato) {
        if (statusDato.getMonth() === curMonth && statusDato.getFullYear() === curYear) {
          stats.omsMåned += prisInkl;
          stats.reiseMåned += reiseInkl;
        }

        const mKey = Utilities.formatDate(new Date(statusDato.getFullYear(), statusDato.getMonth(), 1), 'Europe/Oslo', 'yyyy-MM');
        if (stats.trendMonths[mKey]) {
          stats.trendMonths[mKey].omsInkl += prisInkl;
          stats.trendMonths[mKey].reiseInkl += reiseInkl;
          stats.trendMonths[mKey].totalInkl += (prisInkl + reiseInkl);
        }
      }

      if (prisInkl > 0) {
        sumPrisInkl += prisInkl;
        sumPrisEks += prisEks;
        countPris += 1;
      }
    }
  }

  stats.total -= ordreStageCount;
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
  rows.push([
    'Aktive (ikke Fakturert)',
    stats.total - (stats.byStatus['Fakturert'] || 0) - (stats.byStatus['Oppdrag kansellert'] || 0) - (stats.byStatus['Oppdrag fullført'] || 0),
    '', '', ''
  ]);

  rows.push(['Utestående fordringer (Eks mva)', formatCurrency_(stats.uteståendeEks), '', '', '']);
  rows.push(['Utestående fordringer (Inkl mva)', formatCurrency_(stats.uteståendeInkl), '', '', '']);
  rows.push(['Gjennomsnitt pris per oppdrag (Eks mva)', formatCurrency_(stats.snittPrisEks), '', '', '']);
  rows.push(['Gjennomsnitt pris per oppdrag (Inkl mva)', formatCurrency_(stats.snittPrisInkl), '', '', '']);

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
  rows.push(['Totalt', stats.total, '', '', '']);
  Object.keys(stats.byType).sort().forEach(function (t) { rows.push([t, stats.byType[t], '', '', '']); });

  rows.push(['', '', '', '', '']);

  rows.push(['🏠 BOLIGSTØRRELSE', 'Antall', '', '', '']);
  Object.keys(stats.byBoligBucket).sort().forEach(function (k) { rows.push([k, stats.byBoligBucket[k], '', '', '']); });

  rows.push(['', '', '', '', '']);

  rows.push(['📈 INNTJENING SISTE ' + monthsToShow + ' MND (Inkl mva)', 'Oppdrag', 'Reise', 'Total', '']);
  monthKeys.forEach(function (k) {
    const v = stats.trendMonths[k];
    rows.push([k, v.omsInkl, v.reiseInkl, v.totalInkl, '']);
  });


  dash.getRange(1, 1, rows.length, 5).setValues(rows);

  dash.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  dash.setColumnWidth(1, 360);
  dash.setColumnWidth(2, 180);
  dash.setColumnWidth(3, 180);
  dash.setColumnWidth(4, 180);
  dash.setColumnWidth(5, 60);

  const chartHeaderRow = rows.findIndex(r => String(r[0]).indexOf('📈 INNTJENING SISTE') === 0) + 1; // 1-basert radnr i arket
  if (chartHeaderRow > 0) {
    const firstDataRow = chartHeaderRow + 1;
    dash.getRange(firstDataRow, 2, monthsToShow, 3).setNumberFormat('#,##0');
  }

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

function parseDateString_(dateStr) {
  if (dateStr instanceof Date) return dateStr;
  if (typeof dateStr !== 'string' || !dateStr) return null;
  const parts = dateStr.split(' ')[0].split('.');
  if (parts.length >= 3) return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
  return null;
}



// ============================================================
// HTML BUILDERS
// ============================================================

function buildNewOppdragEmail_(parsed, dato, folderUrl, oppdragsnr, avstandKm, reiseEks, reiseInkl) {
  const fields = [
    ['Oppdragsnr', oppdragsnr], ['Type', parsed.oppdragstype], ['Kilde', parsed.kilde],
    ['Adresse', parsed.adresse], ['Oppdragsgiver', parsed.oppdragsgiver],
    ['Selger', parsed.selger], ['Selger tlf', parsed.selgerTlf],
    ['Megler', parsed.megler], ['Faktura ref', parsed.fakturaRef],
    ['Faktura til', parsed.fakturaSendesTil], ['Mottatt', dato],
  ];
  if (avstandKm) {
    fields.push(['Avstand t/r', avstandKm + ' km']);
    fields.push(['Reise eks mva', formatCurrency_(reiseEks)]);
    fields.push(['Reise inkl mva', formatCurrency_(reiseInkl)]);
  }

  let detailRows = '';
  fields.forEach(function (f, i) {
    if (f[1]) detailRows += '<tr style="background:' + (i % 2 === 0 ? '#fff' : '#f9f9f9') + '"><td style="padding:8px;font-weight:bold;">' + f[0] + '</td><td style="padding:8px;">' + f[1] + '</td></tr>';
  });

  return '<div style="font-family:Arial,sans-serif;max-width:600px;">' +
    '<div style="background:#1a5c2a;color:white;padding:20px;border-radius:8px 8px 0 0;">' +
    '<h2 style="margin:0;">📋 ' + parsed.oppdragstype + '</h2></div>' +
    '<div style="padding:20px;border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;">' +
    (CONFIG.TEST_MODE ? '<div style="background:#fff3e0;padding:8px;border-radius:4px;margin-bottom:12px;">🧪 TESTMODUS — ingen eksterne e-poster sendt</div>' : '') +
    '<table style="width:100%;border-collapse:collapse;">' + detailRows + '</table>' +
    '<p style="margin-top:16px;"><a href="' + folderUrl + '" style="background:#1a5c2a;color:white;padding:10px 20px;text-decoration:none;border-radius:4px;">📁 Åpne mappe</a></p></div></div>';
}


function buildFakturaEmail_(rowData, datoStr) {
  const oppdragsnr = rowData[COL.OPPDRAGSNR - 1] || '';
  let adresseFull = String(rowData[COL.ADRESSE - 1] || '');

  const hyperMatch = adresseFull.match(/=HYPERLINK\("[^"]+","([^"]+)"\)/);
  if (hyperMatch) adresseFull = hyperMatch[1].replace(/""/g, '"');

  const selger = rowData[COL.SELGER - 1] || '';
  const selgerTlf = rowData[COL.SELGER_TLF - 1] || '';
  const selgerEpost = rowData[COL.SELGER_EPOST - 1] || '';
  const megler = rowData[COL.MEGLER - 1] || '';
  const meglerEpost = rowData[COL.MEGLER_EPOST - 1] || '';
  const fakturaRef = rowData[COL.FAKTURA_REF - 1] || '';
  const fakturaTil = rowData[COL.FAKTURA_SENDES_TIL - 1] || '';
  const rapporttype = rowData[COL.RAPPORTTYPE - 1] || rowData[COL.OPPDRAGSTYPE - 1] || 'Oppdrag';

  const prisEks = parseFloat(rowData[COL.PRIS_EKS - 1]) || 0;
  const reiseEks = parseFloat(rowData[COL.REISE_EKS - 1]) || 0;
  const kommentarRegnskap = String(rowData[COL.KOMMENTAR_REGNSKAP - 1] || '').trim();

  let gate = adresseFull;
  let postnr = '';
  let poststed = '';
  const match = adresseFull.match(/(.+?)(?:,\s*|\s+)(\d{4})\s+(.+)/);
  if (match) {
    gate = match[1].trim();
    postnr = match[2].trim();
    poststed = match[3].trim();
  }

  const isMegler = fakturaTil && selger && fakturaTil.toLowerCase().indexOf(selger.toLowerCase()) === -1;
  const kundeNavn = fakturaTil || selger || 'Ikke oppgitt';
  const kundeEpost = (isMegler && meglerEpost) ? meglerEpost : (selgerEpost || meglerEpost || '');
  const kundeRef = megler || kundeNavn;

  let html = '<div style="font-family:Arial,sans-serif; max-width:700px; color:#333;">';
  html += '<div style="background:#1a5c2a;color:white;padding:15px;border-radius:6px 6px 0 0;">';
  html += '<h2 style="margin:0;">Klar til fakturering: ' + oppdragsnr + '</h2></div>';
  html += '<div style="padding:25px; border:1px solid #ddd; border-top:none; border-radius:0 0 6px 6px;">';

  if (kommentarRegnskap) {
    html += '<div style="background-color: #fff3cd; color: #856404; padding: 15px; border-left: 5px solid #ffeeba; margin-bottom: 25px; border-radius: 4px;">';
    html += '<strong style="display:block; margin-bottom:5px; font-size:16px;">💬 Melding til regnskap:</strong>';
    html += kommentarRegnskap.replace(/\n/g, '<br>');
    html += '</div>';
  }

  // --- A. KUNDE ---
  html += '<h3 style="color:#1a5c2a; border-bottom:1px solid #ddd; padding-bottom:5px; margin-top:0;">Kunde</h3>';
  html += '<table style="width:100%; border-collapse:collapse; line-height:1.8;">';
  html += '<tr><td style="width:30px; color:#666;">1.</td><td style="width:180px; font-weight:bold;">Navn/Firma:</td><td>' + kundeNavn + '</td></tr>';
  html += '<tr><td style="color:#666;">2.</td><td style="font-weight:bold;">Telefon:</td><td>' + selgerTlf + '</td></tr>';
  html += '<tr><td style="color:#666;">3.</td><td style="font-weight:bold;">E-post:</td><td>' + kundeEpost + '</td></tr>';
  html += '<tr><td style="color:#666;">4.</td><td style="font-weight:bold;">Adresse:</td><td>' + gate + '</td></tr>';
  html += '<tr><td style="color:#666;">5.</td><td style="font-weight:bold;">Postnr.:</td><td>' + postnr + '</td></tr>';
  html += '<tr><td style="color:#666;">6.</td><td style="font-weight:bold;">Poststed:</td><td>' + poststed + '</td></tr>';
  html += '</table>';

  // --- B. FAKTURA ---
  html += '<h3 style="color:#1a5c2a; border-bottom:1px solid #ddd; padding-bottom:5px; margin-top:30px;">Faktura</h3>';
  html += '<table style="width:100%; border-collapse:collapse; line-height:1.8;">';
  html += '<tr><td style="width:30px; color:#666;">1.</td><td style="width:200px; font-weight:bold;">Kundens referanse:</td><td>' + kundeRef + '</td></tr>';
  html += '<tr><td style="color:#666;">2.</td><td style="font-weight:bold;">PO-nr./Ordrenr.:</td><td>' + (fakturaRef || '-') + '</td></tr>';
  html += '<tr><td style="color:#666;">3.</td><td style="font-weight:bold;">Eiendom eier:</td><td>' + selger + '</td></tr>';
  html += '<tr><td style="color:#666;">4.</td><td style="font-weight:bold;">Eiendom adresse:</td><td>' + gate + '</td></tr>';
  html += '<tr><td style="color:#666;">5.</td><td style="font-weight:bold;">Postnr. + Poststed:</td><td>' + (postnr ? postnr + ' ' + poststed : poststed) + '</td></tr>';
  html += '</table>';

  // --- C. PRODUKTER ---
  html += '<h4 style="margin-top:25px; margin-bottom:10px; color:#666;">6. Produkter (Pris eks. mva)</h4>';
  html += '<table style="width:100%; border-collapse:collapse; border:1px solid #ddd; font-size:14px;">';
  html += '<tr style="background:#f1f1f1; color:#333; font-weight:bold; text-align:left;">';
  html += '<th style="padding:10px; border-bottom:1px solid #ddd;">Produktnr.</th>';
  html += '<th style="padding:10px; border-bottom:1px solid #ddd;">Produktnavn</th>';
  html += '<th style="padding:10px; border-bottom:1px solid #ddd; text-align:right;">Pris eks. mva</th></tr>';

  const totalEks = prisEks;
  const inklMarked = rowData[COL.MED_MARKEDSVERDI - 1] === true;
  const antallTillegg = parseInt(rowData[COL.ANTALL_TILLEGGSBYGG - 1]) || 0;
  const boligtype = rowData[COL.BOLIGTYPE - 1] || '';
  const prodNrRaw = String(rowData[COL.PRODUKTNUMMER - 1] || '');

  let baseProd = prodNrRaw.split(',')[0].trim();
  let markedProd = prodNrRaw.includes(',') ? prodNrRaw.split(',')[1].trim() : '9';

  // --- Hent Areal-tekst fra Prisliste ---
  let storrelseTekst = '';
  const prislisteObj = getPricesFromSheet_();
  if (prislisteObj && boligtype && prislisteObj[boligtype]) {
    const arealNum = parseFloat(rowData[COL.AREAL - 1]) || 0;
    const alternativer = prislisteObj[boligtype];
    for (let i = 0; i < alternativer.length; i++) {
      if (arealNum <= alternativer[i].maxAreal || alternativer[i].maxAreal === Infinity) {
        storrelseTekst = alternativer[i].label; // F.eks "Under 80 m²"
        break;
      }
    }
  }

  // Regn baklengs fra totalpris for å finne linjeprisene eks mva
  let markedEks = (inklMarked && boligtype !== 'Frittstående bygg') ? ((prislisteObj && prislisteObj.markedsverdi ? prislisteObj.markedsverdi.pris : 2000) / 1.25) : 0;
  let tilleggEks = (antallTillegg * 1250) / 1.25;
  let baseEks = totalEks - markedEks - tilleggEks;
  if (baseEks < 0) { baseEks = totalEks; markedEks = 0; tilleggEks = 0; }

  const addRow = (pNr, pNavn, pPris) => {
    html += '<tr style="background:#fff;">';
    html += '<td style="padding:10px; border-bottom:1px solid #ddd;">' + pNr + '</td>';
    html += '<td style="padding:10px; border-bottom:1px solid #ddd;">' + pNavn + '</td>';
    html += '<td style="padding:10px; border-bottom:1px solid #ddd; text-align:right;">' + formatCurrency_(pPris) + '</td></tr>';
  };

  // Bygg produktnavn: f.eks "Tilstandsrapport - Leilighet (Under 80 m²)"
  let baseNavn = rapporttype;
  if (boligtype) baseNavn += ' - ' + boligtype;
  if (storrelseTekst && boligtype !== 'Frittstående bygg') baseNavn += ' (' + storrelseTekst + ')';

  addRow(baseProd || '-', baseNavn, baseEks);
  if (markedEks > 0) addRow(markedProd, 'Tilstandsrapport markedsverdi', markedEks);
  if (tilleggEks > 0) addRow('', 'Tilleggsbygg (' + antallTillegg + ' stk)', tilleggEks);
  if (reiseEks > 0) addRow('', 'Reisekostnad (inkl. evt bom/ferge)', reiseEks);

  html += '<tr style="font-weight:bold; background:#e8f5e9; color:#1a5c2a;">';
  html += '<td colspan="2" style="padding:12px;">TOTAL EKS. MVA:</td>';
  html += '<td style="padding:12px; text-align:right;">' + formatCurrency_(totalEks + reiseEks) + '</td></tr>';
  html += '</table>';

  if (CONFIG.TEST_MODE) {
    html += '<div style="background:#fff3e0;padding:8px;border-radius:4px;margin-top:20px;">🧪 TESTMODUS</div>';
  }

  html += '</div></div>';
  return html;
}


function buildWeeklyReportHtml_(fakturerbare, fakturerte, ventende, totals) {
  const ukeNr = getWeekNumber_(new Date());
  let rows = '';
  fakturerbare.forEach(function (o) {
    rows += '<tr><td style="padding:5px;border-bottom:1px solid #eee;">' + o.oppdragsnr + '</td><td style="padding:5px;">' + o.adresse + '</td><td style="padding:5px;">' + o.oppdragstype + '</td><td style="padding:5px;">' + (o.fakturaRef || '-') + '</td><td style="padding:5px;text-align:right;">' + formatCurrency_(o.prisEks) + '</td><td style="padding:5px;text-align:right;">' + formatCurrency_(o.mvaBeløp) + '</td><td style="padding:5px;text-align:right;font-weight:bold;">' + formatCurrency_(o.prisInkl) + '</td><td style="padding:5px;text-align:right;">' + formatCurrency_(o.reiseInkl) + '</td></tr>';
  });

  return '<div style="font-family:Arial,sans-serif;max-width:900px;">' +
    '<div style="background:#1a5c2a;color:white;padding:20px;border-radius:8px 8px 0 0;">' +
    '<h2 style="margin:0;">📊 Ukerapport uke ' + ukeNr + '</h2></div>' +
    '<div style="padding:20px;border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;">' +
    (CONFIG.TEST_MODE ? '<div style="background:#fff3e0;padding:8px;border-radius:4px;margin-bottom:12px;">🧪 TESTMODUS</div>' : '') +
    '<h3>💰 Fakturerbart (' + fakturerbare.length + ')</h3>' +
    (fakturerbare.length > 0 ?
      '<table style="width:100%;border-collapse:collapse;font-size:11px;"><tr style="background:#f5f5f5;"><th style="padding:5px;text-align:left;">Nr</th><th>Adresse</th><th>Type</th><th>Ref</th><th style="text-align:right;">Eks mva</th><th style="text-align:right;">MVA</th><th style="text-align:right;">Inkl mva</th><th style="text-align:right;">Reise inkl</th></tr>' + rows +
      '<tr style="background:#1a5c2a;color:white;font-weight:bold;"><td colspan="4">TOTAL</td><td colspan="4" style="text-align:right;font-size:14px;">' + formatCurrency_(totals.totalInklMva + totals.totalReiseInkl) + '</td></tr></table>' :
      '<p style="color:#666;">Ingen denne uken.</p>') +
    '</div></div>';
}


// ============================================================
// HJELPEFUNKSJONER
// ============================================================

function setDropdown_(sheet, col, values) {
  sheet.getRange(2, col, 500).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build()
  );
}

function getOrCreateRootFolder_() {
  if (CONFIG.ROOT_FOLDER_ID) {
    return DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  }
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

function setupConditionalFormatting_(sheet, colCount) {
  const range = sheet.getRange(2, 1, 500, colCount);
  const rules = [
    { status: 'Mottatt', color: '#fff3e0' },
    { status: 'Avtalt befaring', color: '#e3f2fd' },
    { status: 'Befart', color: '#e8eaf6' },
    { status: 'Utkast', color: '#fce4ec' },
    { status: 'Endelig rapport', color: '#f3e5f5' },
    { status: 'Kan faktureres', color: '#e8f5e9' },
    { status: 'Fakturert', color: '#c8e6c9' },
    { status: 'Oppdrag kansellert', color: '#eeeeee' },
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
  const fns = ['scanIncomingEmails', 'checkReminders', 'sendWeeklyReport', 'updateDashboard', 'onEditHandler', 'processIVITScraping'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (fns.indexOf(t.getHandlerFunction()) > -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scanIncomingEmails').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('checkReminders').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('sendWeeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(16).create();
  ScriptApp.newTrigger('updateDashboard').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('processIVITScraping').timeBased().everyMinutes(15).create();


  if (!ss) ss = getSpreadsheet_();
  if (ss) {
    ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss).onEdit().create();
  }
}


// ============================================================
// MENY + TEST (alle test-data bruker KUN test-epostadresser)
// ============================================================
function onOpen() {
  const mode = CONFIG.TEST_MODE ? ' 🧪' : '';
  SpreadsheetApp.getUi()
    .createMenu('🏠 Naava Takst' + mode)
    .addItem('📝 Åpne manuell inntak', 'setupManualIntakeSheet_')
    .addItem('✅ Registrer oppdrag fra inntak', 'registerManualFromSheet_')
    .addSeparator()
    //.addItem('📋 Test: Megler (Krogsveen)', 'testMeglerOppdrag')
    //.addItem('📋 Test: IVIT-oppdrag', 'testIVITOppdrag')
    //.addItem('📋 Test: Skadetakst', 'testSkadetakst')
    //.addItem('📋 Test: Reklamasjon', 'testReklamasjon')
    //.addItem('🧪 Test OpenAI API-vurdering', 'testAIAssessment')
    .addSeparator()
    .addItem('📊 Oppdater dashboard', 'updateDashboard')
    .addItem('📧 Send ukerapport nå', 'testWeeklyReport')
    .addItem('🔍 Scan e-post nå', 'scanIncomingEmails')
    .addItem('🔄 Rescan (fjern label + scan)', 'rescanEmails')
    .addItem('🐛 Debug scan', 'debugScan')
    .addItem('⏰ Sjekk påminnelser', 'checkReminders')
    .addSeparator()
    //.addItem('🧪 Test reiseberegning', 'testReiseberegning')
    //.addItem('🧪 Test iVit-tilkobling', 'testIvitConnection')
    //.addItem('⚙️ Kjør oppsett', 'initialSetup')
    .addItem('🌐 Hent manglende iVit-data nå', 'processIVITScraping')
    .addItem('💸 Send faktura til regnskap', 'sendFakturaTilRegnskap')
    .addItem('🙈 Skjul fullførte oppdrag', 'skjulFullforteOppdrag')
    .addItem('👁️ Vis alle oppdrag', 'visAlleOppdrag')
    .addItem('📊 Sorter etter status', 'sorterEtterStatus')

    .addToUi();
}


// ⚠️ ALL testdata bruker KUN adm@afki.no og jacob.e.holen@gmail.com
function testMeglerOppdrag() {
  createNewOppdrag_({
    kilde: 'Megler-epost',
    oppdragstype: 'Tilstandsrapport m/markedsverdi',
    adresse: 'Brattebergringen 2B, Ålesund',
    oppdragsgiver: 'Test Oppdragsgiver',
    selger: 'Test Selger',
    selgerTlf: '99900000',
    selgerEpost: CONFIG.ACCOUNTANT_EMAIL,  // ← Kunde = jacob
    megler: 'Test Megler (Krogsveen)',
    meglerEpost: CONFIG.ACCOUNTANT_EMAIL,  // ← Innsender = jacob
    fakturaRef: '76260009',
    fakturaSendesTil: 'Krogsveen (test)',
    notater: 'Brattebergringen 2B, gnr. 13, bnr. 521, snr. 2'
  }, new Date());
  SpreadsheetApp.getUi().alert('✅ Megler-oppdrag: Brattebergringen 2B');
}

function testIVITOppdrag() {
  createNewOppdrag_({
    kilde: 'IVIT',
    oppdragstype: 'Tilstandsrapport',
    adresse: 'Åsebøen 9, 6017 Ålesund',
    oppdragsgiver: '',
    selger: '',
    selgerTlf: '',
    selgerEpost: '',
    megler: 'Test Megler (EM1)',
    meglerEpost: CONFIG.ACCOUNTANT_EMAIL,  // ← Kun test-epost
    fakturaRef: 'E20260120-10154',
    fakturaSendesTil: 'EiendomsMegler 1 (test)',
    notater: ''
  }, new Date());
  SpreadsheetApp.getUi().alert('✅ IVIT-oppdrag: Åsebøen 9');
}

function testSkadetakst() {
  createNewOppdrag_({
    kilde: 'Manuell',
    oppdragstype: 'Skadetakst',
    adresse: 'Keiser Wilhelms gate 29, Ålesund',
    oppdragsgiver: 'Test Forsikring AS',
    selger: 'Test Person',
    selgerTlf: '99900000',
    selgerEpost: CONFIG.ACCOUNTANT_EMAIL,  // ← Kunde = jacob
    megler: '',
    meglerEpost: '',
    fakturaRef: 'SK-2026-0042',
    fakturaSendesTil: 'Test Forsikring AS',
    notater: 'Vannskade 2. etasje (testdata)'
  }, new Date());
  SpreadsheetApp.getUi().alert('✅ Skadetakst: Keiser Wilhelms gate 29');
}

function testReklamasjon() {
  createNewOppdrag_({
    kilde: 'Manuell',
    oppdragstype: 'Reklamasjon',
    adresse: 'Borgundveien 100, Ålesund',
    oppdragsgiver: 'Test Advokat',
    selger: '',
    selgerTlf: '',
    selgerEpost: '',
    megler: '',
    meglerEpost: CONFIG.ACCOUNTANT_EMAIL,  // ← Kun test-epost
    fakturaRef: 'REK-2026-015',
    fakturaSendesTil: 'Test Advokat',
    notater: 'Reklamasjon fukt kjeller (testdata)'
  }, new Date());
  SpreadsheetApp.getUi().alert('✅ Reklamasjon: Borgundveien 100');
}

function testReiseberegning() {
  const result = calculateDistance_('Brattebergringen 2B, Ålesund');
  if (result) {
    const cost = calculateTravelCost_(result.kmTurRetur);
    SpreadsheetApp.getUi().alert(
      '🚗 ' + CONFIG.BASE_ADDRESS + ' → Brattebergringen 2B\n\n' +
      'Én vei: ' + result.kmEnVei + ' km\n' +
      'Tur-retur: ' + result.kmTurRetur + ' km\n' +
      'Tid: ~' + result.estimertTidMin + ' min\n\n' +
      'Inkludert: ' + CONFIG.REISE_INKLUDERT_KM + ' km\n' +
      'Fakturerbart: ' + cost.fakturerbarKm + ' km\n' +
      'Sats: ' + CONFIG.REISE_SATS_EKS_MVA + ' kr/km eks mva\n\n' +
      'Reise eks mva: ' + cost.kostnadEksMva + ' kr\n' +
      'Reise inkl mva: ' + cost.kostnadInklMva + ' kr'
    );
  } else {
    SpreadsheetApp.getUi().alert('⚠️ Sjekk at Claude API-nøkkel er satt i CONFIG');
  }
}

// DEBUG: Vis nøyaktig hva Gmail finner
function debugScan() {
  const ui = SpreadsheetApp.getUi();

  // Test 1: Helt åpent søk fra Jacob
  const q1 = 'from:jacob.e.holen@gmail.com newer_than:1d';
  const t1 = GmailApp.search(q1, 0, 10);
  let msg1 = 'TEST 1: from:jacob newer_than:1d\nFant: ' + t1.length + ' tråder\n';
  t1.forEach(function (t) {
    const m = t.getMessages()[0];
    const labels = t.getLabels().map(function (l) { return l.getName(); }).join(', ');
    msg1 += '  → "' + m.getSubject() + '" | Labels: [' + labels + ']\n';
  });

  // Test 2: Med keywords
  const q2 = 'from:jacob.e.holen@gmail.com (takst OR tilstandsrapport) newer_than:1d';
  const t2 = GmailApp.search(q2, 0, 10);
  let msg2 = '\nTEST 2: + keywords (takst OR tilstandsrapport)\nFant: ' + t2.length + ' tråder\n';
  t2.forEach(function (t) {
    const m = t.getMessages()[0];
    msg2 += '  → "' + m.getSubject() + '"\n';
  });

  // Test 3: Med label-filter
  const q3 = 'from:jacob.e.holen@gmail.com (takst OR tilstandsrapport) -label:Takst-Behandlet newer_than:1d';
  const t3 = GmailApp.search(q3, 0, 10);
  let msg3 = '\nTEST 3: + minus label\nFant: ' + t3.length + ' tråder\n';

  // Test 4: Sjekk om labelen eksisterer
  const label = GmailApp.getUserLabelByName('Takst-Behandlet');
  let msg4 = '\nLabel "Takst-Behandlet" eksisterer: ' + (label ? 'JA' : 'NEI') + '\n';
  if (label) {
    const lt = label.getThreads();
    msg4 += 'Tråder med label: ' + lt.length + '\n';
    lt.forEach(function (t) {
      const m = t.getMessages()[0];
      msg4 += '  → "' + m.getSubject() + '"\n';
    });
  }

  ui.alert(msg1 + msg2 + msg3 + msg4);
}


// Fjern label fra alle test-eposter og scan på nytt
function rescanEmails() {
  const label = GmailApp.getUserLabelByName('Takst-Behandlet');
  if (label) {
    const threads = label.getThreads();
    threads.forEach(function (thread) {
      thread.removeLabel(label);
    });
    Logger.log('🔄 Fjernet label fra ' + threads.length + ' tråder');
  }

  // Slett eksisterende data (behold headers)
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }

  // Kjør scan
  scanIncomingEmails();
  SpreadsheetApp.getUi().alert('🔄 Rescan fullført! Sjekk Oppdragslogg.');
}


function testWeeklyReport() {
  sendWeeklyReport();
  SpreadsheetApp.getUi().alert('✅ Ukerapport sendt!');
}


function archiveRow_(sheet, row) {
  const ss = getSpreadsheet_();
  const archive = ss.getSheetByName('arkiv');
  if (!archive) {
    Logger.log('FEIL: Fant ikke arkiv-arket "arkiv"');
    return;
  }

  const rowData = sheet.getRange(row, 1, 1, NUM_COLS).getValues()[0];

  // Sørg for headers i arkiv hvis arket er tomt
  if (archive.getLastRow() === 0) {
    const headers = sheet.getRange(1, 1, 1, NUM_COLS).getValues();
    archive.getRange(1, 1, 1, NUM_COLS).setValues(headers);
  }

  // Unngå duplikat i arkiv basert på Oppdragsnr
  const oppdragsnr = rowData[COL.OPPDRAGSNR - 1];
  if (archive.getLastRow() >= 2) {
    const existing = archive.getRange(2, 1, archive.getLastRow() - 1, 1).getValues().flat();
    if (existing.indexOf(oppdragsnr) > -1) {
      Logger.log('Arkiv: Duplikat oppdragsnr, hopper over: ' + oppdragsnr);
      return;
    }
  }

  archive.appendRow(rowData);
  Logger.log('Arkiv: La til rad for ' + oppdragsnr);
}

function detectRapporttype_(fullText) {
  const t = (fullText || '').toLowerCase();

  if (t.indexOf('markedsverdi') > -1) return 'Tilstandsrapport m/teknisk og markedsverdi';
  if (t.indexOf('tilstandsrapport') > -1) return 'Tilstandsrapport';
  if (t.indexOf('skadetakst') > -1 || t.indexOf('skaderapport') > -1) return 'Skadetakstrapport';
  if (t.indexOf('reklamasjon') > -1) return 'Reklamasjonsrapport';
  if (t.indexOf('vurdering') > -1 || t.indexOf('verditakst') > -1) return 'Vurderingsrapport';
  if (t.indexOf('overtag') > -1 || t.indexOf('overtak') > -1) return 'Overtagelsesrapport';

  return 'Annen rapport';
}


function mapDetectedToOppdragstype_(detected) {
  const t = (detected || '').toLowerCase();

  if (t.indexOf('markedsverdi') > -1) return 'Tilstandsrapport m/markedsverdi';
  if (t.indexOf('tilstandsrapport') > -1) return 'Tilstandsrapport';

  if (t.indexOf('skade') > -1) return 'Skadetakst';
  if (t.indexOf('reklamasjon') > -1) return 'Reklamasjon';
  if (t.indexOf('vurdering') > -1) return 'Vurderingsoppdrag';
  if (t.indexOf('overtag') > -1 || t.indexOf('overtak') > -1) return 'Bistand overtagelse';
  if (t.indexOf('fukt') > -1) return 'Fukt-/fuktskadevurdering';
  if (t.indexOf('byggelån') > -1) return 'Byggelånskontroll';
  if (t.indexOf('forhånd') > -1) return 'Forhåndstakst';
  if (t.indexOf('verditakst') > -1) return 'Verditakst';

  return 'Annet';
}

// Mapper HTML-verdier til prosjektets dropdown-verdier (Oppdragstype + Rapporttype + Boligtype)
function mapManualTypes_(boligtypeUi, rapporttypeUi) {
  const bt = String(boligtypeUi || '').toLowerCase();
  let boligtype = 'Annet';
  if (bt.indexOf('leilighet') > -1) boligtype = 'Leilighet';
  else if (bt.indexOf('rekkehus') > -1) boligtype = 'Rekkehus/leilighet 2-4-mannsbolig';
  else if (bt.indexOf('enebolig') > -1) boligtype = 'Enebolig/fritidsbolig';

  const rt = String(rapporttypeUi || '').toLowerCase();
  let oppdragstype = 'Annet';
  let rapporttype = 'Annen rapport';

  if (rt === 'tilstandsrapport') {
    oppdragstype = 'Tilstandsrapport';
    rapporttype = 'Tilstandsrapport';
  } else if (rt.indexOf('markedsverdi') > -1) {
    oppdragstype = 'Tilstandsrapport m/markedsverdi';
    rapporttype = 'Tilstandsrapport m/teknisk og markedsverdi';
  } else if (rt.indexOf('skadetakst') > -1) {
    oppdragstype = 'Skadetakst';
    rapporttype = 'Skadetakstrapport';
  } else if (rt.indexOf('reklamasjon') > -1) {
    oppdragstype = 'Reklamasjon';
    rapporttype = 'Reklamasjonsrapport';
  } else if (rt.indexOf('forhåndstakst') > -1) {
    oppdragstype = 'Forhåndstakst';
    rapporttype = 'Annen rapport';
  }

  return { boligtype: boligtype, oppdragstype: oppdragstype, rapporttype: rapporttype };
}


function diagnoseRootFolderWrite() {
  const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  Logger.log('Root folder: ' + root.getName() + ' | ID: ' + root.getId());

  const test = root.createFolder('__root_write_test__ ' + new Date().toISOString());
  Logger.log('Created under root OK: ' + test.getId());

  test.createFolder('Bilder');
  test.createFolder('Dokumenter fra megler');
  test.createFolder('Rapport');
  Logger.log('Created child folders OK');

  test.setTrashed(true);
  Logger.log('Trashed test folder OK');
  return true;
}

function setupManualIntakeSheet_() {
  const ss = getSpreadsheet_();
  const name = 'Manuell inntak';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();

  sh.getRange(1, 1).setValue('Manuell registrering (fyll ut feltene og trykk "Registrer oppdrag")');
  sh.getRange(1, 1).setFontWeight('bold');

  const rows = [
    ['Adresse *', ''],
    ['Selger/kontaktperson', ''],
    ['Telefon', ''],
    ['E-post (kunde)', ''],
    ['Boligtype * (må matche dropdown i Oppdragslogg)', ''],
    ['Rapporttype * (må matche dropdown i Oppdragslogg)', ''],
    ['Areal (ca. m²)', ''],
    ['Megler / bestiller', ''],
    ['Merknad', '']
  ];

  sh.getRange(3, 1, rows.length, 2).setValues(rows);
  sh.setColumnWidth(1, 360);
  sh.setColumnWidth(2, 420);

  sh.getRange(3, 1, rows.length, 1).setFontWeight('bold');
  sh.getRange(3, 2, rows.length, 1).setBackground('#fffde7');

  sh.getRange(14, 1).setValue('Status:');
  sh.getRange(14, 2).setValue('Klar');

  // Legg enkel datavalidering for boligtype/rapporttype basert på dine eksisterende lister
  const boligTyper = ['Leilighet', 'Rekkehus/leilighet 2-4-mannsbolig', 'Enebolig/fritidsbolig', 'Frittstående bygg', 'Næringsbygg', 'Annet'];
  const rapportTyper = [
    'Tilstandsrapport m/teknisk og markedsverdi',
    'Tilstandsrapport',
    'Skadetakstrapport',
    'Reklamasjonsrapport',
    'Vurderingsrapport',
    'Overtagelsesrapport',
    'Annen rapport'
  ];

  sh.getRange(7, 2).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(boligTyper, true).setAllowInvalid(false).build()
  );
  sh.getRange(8, 2).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(rapportTyper, true).setAllowInvalid(false).build()
  );
}

function registerManualFromSheet_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName('Manuell inntak');
  if (!sh) throw new Error('Fant ikke arket "Manuell inntak". Kjør setupManualIntakeSheet_() først.');

  const v = (r) => String(sh.getRange(r, 2).getValue() || '').trim();

  const data = {
    adresse: v(3),
    selger: v(4),
    telefon: v(5),
    epost: v(6),
    boligtype: v(7),
    rapporttype: v(8),
    areal: v(9),
    megler: v(10),
    merknad: v(11)
  };

  if (!data.adresse || !data.boligtype || !data.rapporttype) {
    sh.getRange(14, 2).setValue('Feil: fyll ut Adresse, Boligtype, Rapporttype');
    return;
  }

  sh.getRange(14, 2).setValue('Registrerer...');

  // Bygg parsed og bruk eksisterende flyt
  const mapped = mapManualTypes_(data.boligtype, data.rapporttype);

  let parsed = {
    kilde: 'Manuell',
    oppdragstype: mapped.oppdragstype,
    adresse: data.adresse,
    oppdragsgiver: data.megler || data.selger || '',
    selger: data.selger || '',
    selgerTlf: String(data.telefon || '').replace(/[^\d+]/g, ''),
    selgerEpost: data.epost || '',
    megler: data.megler || '',
    meglerEpost: '',
    fakturaRef: '',
    fakturaSendesTil: '',
    notater: data.merknad || '',
    rapporttype: mapped.rapporttype
  };

  if (CONFIG.TEST_MODE) parsed = sanitizeParsedData_(parsed);

  const now = new Date();
  createNewOppdrag_(parsed, now, true);

  // Fyll boligtype/areal og kjør prisberegning på siste rad
  const logg = ss.getSheetByName(CONFIG.SHEET_NAME);
  const row = logg.getLastRow();
  if (mapped.boligtype) logg.getRange(row, COL.BOLIGTYPE).setValue(mapped.boligtype);
  if (data.areal) logg.getRange(row, COL.AREAL).setValue(Number(data.areal));
  calculatePrice_(logg, row);

  // Nullstill input
  sh.getRange(3, 2, 9, 1).clearContent();
  sh.getRange(14, 2).setValue('OK: registrert (sjekk Oppdragslogg)');
}

function getOrCreateAvsluttedeOppdragFolder_() {
  const root = CONFIG.ROOT_FOLDER_ID
    ? DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID)
    : getOrCreateRootFolder_();

  const targetName = 'avsluttede oppdrag';
  const it = root.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    if (String(f.getName() || '').toLowerCase() === targetName) return f;
  }
  return root.createFolder('avsluttede oppdrag');
}

function extractDriveFolderId_(value) {
  const s = String(value || '').trim();
  if (!s) return '';

  // Hvis cellen er en HYPERLINK-formel eller inneholder URL, plukk ut ID
  const m1 = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m1) return m1[1];

  const m2 = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2) return m2[1];

  // Hvis den allerede er en ID
  if (/^[a-zA-Z0-9_-]{10,}$/.test(s) && s.indexOf('http') !== 0) return s;

  return '';
}

function moveOppdragFolderToAvsluttede_(folderId) {
  if (!folderId) throw new Error('Mangler folderId');

  const folder = DriveApp.getFolderById(folderId);
  const targetParent = getOrCreateAvsluttedeOppdragFolder_();

  // Legg til i målmappen
  targetParent.addFolder(folder);

  // Fjern fra alle andre foreldre for å gjøre det til en faktisk "flytt"
  const parents = folder.getParents();
  while (parents.hasNext()) {
    const p = parents.next();
    if (p.getId() !== targetParent.getId()) {
      p.removeFolder(folder);
    }
  }

  return targetParent.getId();
}

function debugFindOrder() {
  const q = 'E20260223-10136';
  const threads = GmailApp.search(q, 0, 5);
  Logger.log('Found threads: ' + threads.length);
  threads.forEach(function (t) {
    const m = t.getMessages()[t.getMessages().length - 1];
    const labels = t.getLabels().map(function (l) { return l.getName(); }).join(', ');
    Logger.log('FROM: ' + m.getFrom());
    Logger.log('SUBJECT: ' + m.getSubject());
    Logger.log('LABELS: ' + labels);
  });
}

function debugWhichMailbox() {
  Logger.log('Effective user: ' + Session.getEffectiveUser().getEmail());
  Logger.log('Active user: ' + Session.getActiveUser().getEmail());

  const t = GmailApp.getInboxThreads(0, 1);
  if (!t.length) {
    Logger.log('Inbox is empty (or not accessible)');
    return;
  }
  const m = t[0].getMessages()[0];
  Logger.log('Sample inbox mail TO: ' + m.getTo());
  Logger.log('Sample inbox mail SUBJECT: ' + m.getSubject());
  Logger.log('Sample inbox mail FROM: ' + m.getFrom());
}

function debugFindOrderAnywhere() {
  const q = 'in:anywhere E20260223-10136';
  const threads = GmailApp.search(q, 0, 5);
  Logger.log('Query: ' + q);
  Logger.log('Found threads: ' + threads.length);
  threads.forEach(function (t) {
    const m = t.getMessages()[t.getMessages().length - 1];
    Logger.log('FROM: ' + m.getFrom());
    Logger.log('TO: ' + m.getTo());
    Logger.log('SUBJECT: ' + m.getSubject());
  });
}

function debugFindIvitByParts() {
  const queries = [
    'in:anywhere E20260223',
    'in:anywhere 10136',
    'in:anywhere subject:E20260223',
    'in:anywhere "mottatt en ordre i IVIT"',
    'in:anywhere "ordre i IVIT"',
    'in:anywhere subject:"er sent til ditt firma"'
  ];

  queries.forEach(function (q) {
    const n = GmailApp.search(q, 0, 5).length;
    Logger.log(q + ' -> ' + n);
  });
}

function debugScanQueryIvit() {
  const keywords = CONFIG.TRIGGER_KEYWORDS.map(function (k) { return '"' + k + '"'; }).join(' OR ');
  const ivitQuery = '(subject:IVIT OR "ordre i ivit" OR "mottatt en ordre i ivit" OR subject:"er sent til ditt firma")';
  const q = '((' + keywords + ') OR ' + ivitQuery + ') newer_than:30d';

  const threads = GmailApp.search(q, 0, 20);
  Logger.log('Query: ' + q);
  Logger.log('Threads: ' + threads.length);

  threads.forEach(function (t) {
    const m = t.getMessages()[t.getMessages().length - 1];
    Logger.log('FROM: ' + m.getFrom());
    Logger.log('SUBJECT: ' + m.getSubject());
  });
}

function debugOrderInThisMailbox() {
  const q = 'in:anywhere E20260223';
  const threads = GmailApp.search(q, 0, 5);
  Logger.log('User: ' + Session.getEffectiveUser().getEmail());
  Logger.log('Found: ' + threads.length);
  threads.forEach(function (t) {
    const m = t.getMessages()[t.getMessages().length - 1];
    Logger.log(m.getSubject());
  });
}

// ============================================================
// DIAGNOSE: Test at assessAndParseWithAI_ faktisk kaller API
// Kjør via Apps Script editor → Run → testAIAssessment
// ============================================================
function testAIAssessment() {
  const ui = SpreadsheetApp.getUi();

  // --- Test 1: Ekte takstoppdrag ---
  const relevant = assessAndParseWithAI_(
    'Forespørsel om tilstandsrapport – Storgata 12, Ålesund',
    'Hei,\n\nVi ønsker tilstandsrapport med markedsverdi på følgende eiendom:\n\nStorgata 12, 6004 Ålesund\n\nSelger: Kari Nordmann\nTlf: 90000000\nE-post: kari@example.com\n\nFaktura sendes til Krogsveen AS, ref. 76543210.\n\nMed vennlig hilsen\nOle Megler\nKrogsveen Ålesund',
    '"Ole Megler" <ole.megler@krogsveen.no>'
  );

  // --- Test 2: Irrelevant e-post ---
  const irrelevant = assessAndParseWithAI_(
    'Nyhetsbrev fra Eiendomsmegler 1 – April 2026',
    'Les vårt nyhetsbrev om boligmarkedet. Klikk her for å melde deg av.',
    '"EM1 Nyheter" <noreply@em1.no>'
  );

  // --- Test 3: Tom e-post (kanttilfelle) ---
  const empty = assessAndParseWithAI_('', '', 'ukjent@example.com');

  // Bygg rapport
  const rel = relevant
    ? '✅ Korrekt: vurdert som RELEVANT\n' +
    '  oppdragstype: ' + relevant.oppdragstype + '\n' +
    '  adresse: ' + relevant.adresse + '\n' +
    '  selger: ' + relevant.selger + '\n' +
    '  meglerEpost: ' + relevant.meglerEpost
    : '❌ Feil: vurdert som IKKE relevant (eller API-feil — sjekk logg)';

  const irrel = irrelevant === null
    ? '✅ Korrekt: vurdert som IKKE RELEVANT (null returnert)'
    : '❌ Feil: skulle vært null, men fikk:\n  oppdragstype: ' + irrelevant.oppdragstype;

  const emp = empty === null
    ? '✅ Korrekt: tom e-post gir null'
    : '⚠️  Tom e-post ga et resultat (sjekk om det er riktig)';

  // Vis HTTP-status ved å kalle råversjon for diagnostikk
  let httpStatus = '';
  try {
    const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + CONFIG.OPENAI_API_KEY
      },
      payload: JSON.stringify({
        model: CONFIG.OPENAI_MODEL,
        temperature: 0,
        max_completion_tokens: 5,
        messages: [{ role: 'user', content: 'ping' }]
      }),
      muteHttpExceptions: true
    });
    httpStatus = '🌐 API HTTP-status: ' + resp.getResponseCode();
    if (resp.getResponseCode() !== 200) {
      const body = resp.getContentText().substring(0, 300);
      httpStatus += '\n  Feilmelding fra API: ' + body;
    }
  } catch (e) {
    httpStatus = '❌ Nettverksfeil mot OpenAI: ' + e.message;
  }

  ui.alert(
    '🧪 OpenAI API-diagnose\n' +
    '────────────────────────────\n' +
    httpStatus + '\n\n' +
    'API-nøkkel starter med: ' + (CONFIG.OPENAI_API_KEY || '').substring(0, 10) + '...\n' +
    'Modell: ' + CONFIG.OPENAI_MODEL + '\n\n' +
    'TEST 1 – Relevant e-post:\n' + rel + '\n\n' +
    'TEST 2 – Irrelevant e-post:\n' + irrel + '\n\n' +
    'TEST 3 – Tom e-post:\n' + emp
  );
}

// ============================================================
// 13. IVIT WEBHOOK SCRAPER
// ============================================================

function fetchIvitData_(address) {
  if (!address || address.trim().length === 0) return { success: false, error: 'Tom adresse' };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ address: address.trim() }),
    muteHttpExceptions: true
  };

  if (CONFIG.IVIT_WEBHOOK_SECRET) {
    options.headers = { 'Authorization': 'Bearer ' + CONFIG.IVIT_WEBHOOK_SECRET };
  }

  try {
    const response = UrlFetchApp.fetch(CONFIG.IVIT_WEBHOOK_URL, options);
    const body = JSON.parse(response.getContentText());
    if (response.getResponseCode() === 200 && body.success) {
      return body;
    } else {
      return { success: false, error: body.error || 'Ukjent feil fra Webhook' };
    }
  } catch (e) {
    return { success: false, error: 'Forespørsel feilet: ' + e.message };
  }
}

function processIVITScraping() {
  Logger.log('START: processIVITScraping kjøres @ ' + new Date().toISOString());

  const ss = getSpreadsheet_();
  if (!ss) {
    Logger.log('FEIL: Kunne ikke hente spreadsheet (getSpreadsheet_() returnerte null/undefined)');
    return;
  }
  Logger.log('Spreadsheet ID: ' + ss.getId());

  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    Logger.log('FEIL: Ark "' + CONFIG.SHEET_NAME + '" finnes ikke');
    return;
  }
  Logger.log('Ark funnet: ' + sheet.getName() + ' | Ark ID: ' + sheet.getSheetId());

  const lastRow = sheet.getLastRow();
  Logger.log('lastRow = ' + lastRow);

  if (lastRow < 2) {
    Logger.log('Ingen data: lastRow < 2 → avslutter');
    return;
  }

  const NUM_COLS = 40; // Sørg for at dette stemmer (AJ = kolonne 36)
  Logger.log('Henter data: rad 2 → ' + lastRow + ', kolonner 1 → ' + NUM_COLS);

  const data = sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();
  Logger.log('Data hentet: ' + data.length + ' rader (forventet: ' + (lastRow - 1) + ')');

  if (data.length === 0) {
    Logger.log('Data-array er tom selv om lastRow >= 2 – noe er galt');
    return;
  }

  let processed = 0;

  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 2;
    const row = data[i];

    Logger.log('──────────────────────────────');
    Logger.log('Behandler rad ' + rowNum);

    // Les rå verdi
    const scanIvitRaw = row[COL.SCAN_IVIT - 1];
    Logger.log('  D (scan IVIT) rå verdi: ' + scanIvitRaw + ' (type: ' + typeof scanIvitRaw + ')');

    // Normaliser til boolean for enklere sjekk
    const scanIvit = scanIvitRaw === true || scanIvitRaw === 'true' || scanIvitRaw === 1;
    Logger.log('  → Tolket som: ' + scanIvit + ' (true/false)');

    if (!scanIvit) {
      Logger.log('  → Hopper over: checkbox ikke avkrysset');
      continue;
    }

    let notater = String(row[COL.NOTATER - 1] || '');
    let adresse = String(row[COL.ADRESSE - 1] || '');

    Logger.log('  Notater (rå): ' + notater.substring(0, 100) + (notater.length > 100 ? '...' : ''));
    Logger.log('  Adresse (rå): ' + adresse);

    // HYPERLINK-parsing
    const hyperMatch = adresse.match(/=HYPERLINK\("[^"]+","([^"]+)"\)/);
    if (hyperMatch) {
      adresse = hyperMatch[1].replace(/""/g, '"');
      Logger.log('  → Adresse ekstrahert fra HYPERLINK: ' + adresse);
    } else {
      Logger.log('  → Ingen HYPERLINK funnet i adresse');
    }

    if (!adresse || adresse.trim() === '') {
      Logger.log('  → FEIL: Ingen gyldig adresse etter parsing → hopper over');
      const newNotat = notater + (notater ? ' | ' : '') + '[iVit Feil: Tom/manglende adresse]';
      sheet.getRange(rowNum, COL.NOTATER).setValue(newNotat);
      sheet.getRange(rowNum, COL.SCAN_IVIT).setValue(false);
      continue;
    }

    Logger.log('Kaller fetchIvitData_ med adresse: ' + adresse);

    const result = fetchIvitData_(adresse);

    Logger.log('fetchIvitData_ returnerte: success = ' + result.success);
    if (!result.success) {
      Logger.log('  → Feilmelding: ' + (result.error || 'Ingen feilmelding gitt'));
    } else {
      Logger.log('  → Data mottatt: ' + JSON.stringify(result.data, null, 2).substring(0, 400) + '...');
    }

    if (result.success) {
      const d = result.data;

      // Grunnleggende oppdragsinfo
      if (d.fakturareferanse) sheet.getRange(rowNum, COL.FAKTURA_REF).setValue(d.fakturareferanse);
      if (d.befaring_dato) sheet.getRange(rowNum, COL.BEFARING_DATO).setValue(d.befaring_dato);
      if (d.befaring_klokkeslett) sheet.getRange(rowNum, COL.BEFARING_KL).setValue(d.befaring_klokkeslett);

      if (ivitDato && ivitDato !== "" && ivitDato !== "Ikke satt") {

        const statusRange = sheet.getRange(row, COL.STATUS);
        const currentStatus = statusRange.getValue();

        // Vi endrer bare hvis den står som 'Mottatt' (for å ikke overskrive manuelle valg)
        if (currentStatus === "Mottatt" || currentStatus === "") {
          statusRange.setValue("Avtalt befaring");
          sheet.getRange(row, COL.DATO_STATUSENDRING).setValue(new Date());
          Logger.log("Befaringsdato funnet i IVIT. Status satt til 'Avtalt befaring'.");
        }
      } else {
        Logger.log("Ingen befaringsdato i IVIT-data. Status forblir 'Mottatt'.");
      }

      // --- NYTT FRA OPUS: Selgerinfo ---
      if (d.selger) sheet.getRange(rowNum, COL.SELGER).setValue(d.selger);
      if (d.selger_tlf) sheet.getRange(rowNum, COL.SELGER_TLF).setValue(d.selger_tlf);
      if (d.selger_epost) sheet.getRange(rowNum, COL.SELGER_EPOST).setValue(d.selger_epost);

      // --- Boligtype, areal og tilleggsbygg ---
      if (d.boligtype) {
        let boligtypeStr = 'Annet'; // Standard hvis den ikke skjønner hva IVIT sender
        const btLow = String(d.boligtype).toLowerCase();

        if (btLow.indexOf('enebolig') > -1 || btLow.indexOf('fritid') > -1) {
          boligtypeStr = 'Enebolig/fritidsbolig';
        } else if (btLow.indexOf('rekkehus') > -1 || btLow.indexOf('mannsbolig') > -1) {
          boligtypeStr = 'Rekkehus/leilighet 2-4-mannsbolig';
        } else if (btLow.indexOf('leilighet') > -1) {
          boligtypeStr = 'Leilighet';
        } else if (btLow.indexOf('næring') > -1) {
          boligtypeStr = 'Næringsbygg';
        } else if (btLow.indexOf('frittstående') > -1 || btLow.indexOf('garasje') > -1) {
          boligtypeStr = 'Frittstående bygg';
        }

        sheet.getRange(rowNum, COL.BOLIGTYPE).setValue(boligtypeStr);
      }
      if (d.areal_bra != null) {
        sheet.getRange(rowNum, COL.AREAL).setValue(d.areal_bra);
      }
      if (d.antall_bygninger != null) {
        sheet.getRange(rowNum, COL.ANTALL_TILLEGGSBYGG).setValue(Math.max(0, d.antall_bygninger - 1));
      }

      // Sett kryss for markedsverdi hvis IVIT sier det
      if (d.med_markedsverdi === true || d.med_markedsverdi === 'true') {
        sheet.getRange(rowNum, COL.MED_MARKEDSVERDI).setValue(true);
      }

      // ⚠️ Tving prisberegning NÅ som vi har fått areal og boligtype fra IVIT!
      calculatePrice_(sheet, rowNum);

      Logger.log('Suksess for rad ' + rowNum);

      // Skru av checkboxen (Scan IVIT) når den er ferdig
      sheet.getRange(rowNum, COL.SCAN_IVIT).setValue(false);
      processed++;

    } else {
      const errorMsg = result.error || 'Ukjent feil fra fetchIvitData_';
      const newNotat = notater + (notater ? ' | ' : '') + '[iVit Feil: ' + errorMsg + ']';
      sheet.getRange(rowNum, COL.NOTATER).setValue(newNotat);
      sheet.getRange(rowNum, COL.SCAN_IVIT).setValue(false);
      Logger.log('  → Feil håndtert, checkbox skrudd AV, notat oppdatert');
    }
  }

  Logger.log('──────────────────────────────');
  Logger.log('FERDIG: Prosesserte ' + processed + ' rader');
  if (processed > 0) {
    Logger.log('Totalt oppdaterte IVIT-oppdrag: ' + processed);
  } else {
    Logger.log('Ingen rader ble behandlet (sjekk om noen hadde AJ=true)');
  }
}

function testIvitConnection() {
  const result = fetchIvitData_('Smithsgata 5, 6100 VOLDA');
  SpreadsheetApp.getUi().alert(
    'Resultat fra iVit:\n\n' + JSON.stringify(result, null, 2)
  );
}



// ============================================================
// 14. DUPLIKAT-SJEKK (Samme adresse, samme måned)
// ============================================================
function flagDuplicateAddress_(sheet, row, adresseText, dateObj) {
  if (!adresseText) return;
  const data = sheet.getDataRange().getValues();
  const targetMonth = dateObj.getMonth();
  const targetYear = dateObj.getFullYear();
  let found = false;

  // Hjelpefunksjon for å vaske adressen (fjerne hyperlink og poststed)
  const cleanAddr = (addr) => {
    let s = String(addr || '').toLowerCase();
    const m = s.match(/=hyperlink\("[^"]+","([^"]+)"\)/);
    if (m) s = m[1].replace(/""/g, '"');
    // Ta kun med gatenavn + husnummer (alt før første komma) for presis match
    return s.split(',')[0].trim();
  };

  const newAddrClean = cleanAddr(adresseText);
  if (!newAddrClean || newAddrClean.length < 4) return; // Ignorer for korte/tomme strenger

  // Gå gjennom alle rader og se etter match i samme måned
  for (let i = 1; i < data.length; i++) {
    if ((i + 1) === row) continue; // Hopp over raden vi sjekker

    let rowDate = data[i][COL.TIMESTAMP - 1];
    if (!rowDate) {
      // Fallback hvis Timestamp mangler
      const dm = data[i][COL.DATO_MOTTATT - 1];
      if (dm) {
        const parts = String(dm).match(/(\d{2})\.(\d{2})\.(\d{4})/);
        if (parts) rowDate = new Date(parts[3], parts[2] - 1, parts[1]);
      }
    }

    if (rowDate instanceof Date && !isNaN(rowDate.getTime())) {
      // Hvis måned og år er likt
      if (rowDate.getMonth() === targetMonth && rowDate.getFullYear() === targetYear) {
        const existingAddrClean = cleanAddr(data[i][COL.ADRESSE - 1]);
        if (existingAddrClean === newAddrClean) {
          found = true;
          break;
        }
      }
    }
  }

  const notatRange = sheet.getRange(row, COL.NOTATER);
  const adresseRange = sheet.getRange(row, COL.ADRESSE);
  let notat = String(notatRange.getValue() || '');
  const warning = '⚠️ Samme adresse reg. tidligere denne mnd!';

  if (found) {
    // Legg til varsel i notater hvis det ikke står der fra før
    if (notat.indexOf(warning) === -1) {
      notatRange.setValue(notat ? notat + ' | ' + warning : warning);
    }
    // Gjør adresse-teksten rød og fet
    adresseRange.setFontColor('red').setFontWeight('bold');
  } else {
    // Fjern advarsel og farge hvis adressen rettes/ikke er duplikat
    if (notat.indexOf(warning) > -1) {
      notat = notat.replace(' | ' + warning, '').replace(warning, '').trim();
      // Rens opp overflødige bindestreker
      if (notat.endsWith('|')) notat = notat.slice(0, -1).trim();
      notatRange.setValue(notat);
    }
    // Nullstill tekstfargen
    adresseRange.setFontColor(null).setFontWeight('normal');
  }
}


// ============================================================
// 15. DYNAMISK PRISLISTE FRA REGNEARK
// ============================================================
function getPricesFromSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Prisliste');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const priser = {
    timesats: { pris: 1500, prodNr: '' },
    markedsverdi: { pris: 2000, prodNr: '9' } // Fallback
  };

  for (let i = 1; i < data.length; i++) {
    const kategoriRaw = String(data[i][0]).trim();
    const strRaw = String(data[i][1]).trim(); // Teksten "Under 80 m²" osv.
    const prisStandard = parseFloat(String(data[i][2]).replace(/\s/g, '')) || 0;
    const prodNr = String(data[i][3] || '').trim(); // Henter nå fra Kolonne D

    if (!kategoriRaw) continue;

    if (kategoriRaw.toLowerCase().includes('timesats')) {
      priser.timesats = { pris: prisStandard, prodNr: prodNr };
      continue;
    }
    if (kategoriRaw.toLowerCase().includes('markedsverdi')) {
      priser.markedsverdi = { pris: prisStandard, prodNr: prodNr };
      continue;
    }

    let kategori = kategoriRaw;
    const katLow = kategoriRaw.toLowerCase();
    if (katLow.includes('rekkehus')) kategori = 'Rekkehus/leilighet 2-4-mannsbolig';
    else if (katLow.includes('enebolig')) kategori = 'Enebolig/fritidsbolig';
    else if (katLow.includes('leilighet')) kategori = 'Leilighet';
    else if (katLow.includes('frittstående')) kategori = 'Frittstående bygg';

    let maxAreal = Infinity;
    const nums = strRaw.match(/\d+/g);
    if (nums) {
      if (strRaw.toLowerCase().includes('under')) maxAreal = parseInt(nums[0], 10);
      else if (strRaw.includes('-')) maxAreal = parseInt(nums[1] || nums[0], 10);
      else if (strRaw.toLowerCase().includes('over')) maxAreal = Infinity;
    }

    if (!priser[kategori]) priser[kategori] = [];
    priser[kategori].push({
      maxAreal: maxAreal,
      prisStandard: prisStandard,
      prodNr: prodNr,
      label: strRaw // Lagrer teksten slik at e-posten kan bruke den
    });
  }

  for (let key in priser) {
    if (Array.isArray(priser[key])) {
      priser[key].sort(function (a, b) { return a.maxAreal - b.maxAreal; });
    }
  }

  return priser;
}

function calculatePrice_(sheet, row) {
  const boligtype = sheet.getRange(row, COL.BOLIGTYPE).getValue();
  const arealStr = sheet.getRange(row, COL.AREAL).getValue();
  const tilleggStr = sheet.getRange(row, COL.ANTALL_TILLEGGSBYGG).getValue();
  const inkluderMarked = sheet.getRange(row, COL.MED_MARKEDSVERDI).getValue() === true;
  const timerStr = sheet.getRange(row, COL.TIMER).getValue();

  if (!boligtype && !timerStr) return;

  const areal = arealStr ? parseFloat(arealStr) : 0;
  const tilleggsbygg = tilleggStr ? parseInt(tilleggStr, 10) : 0;
  const timer = timerStr ? parseFloat(String(timerStr).replace(',', '.')) : 0;

  const prisliste = getPricesFromSheet_();
  if (!prisliste) return;

  let pris = 0;
  let valgtProdNr = '';

  if (timer > 0 && !boligtype) {
    pris = timer * prisliste.timesats.pris;
    valgtProdNr = prisliste.timesats.prodNr;
  }
  else if (boligtype && prisliste[boligtype]) {
    const alternativer = prisliste[boligtype];
    if (boligtype === 'Frittstående bygg') {
      pris = alternativer[0].prisStandard;
      valgtProdNr = alternativer[0].prodNr;
    } else {
      for (let i = 0; i < alternativer.length; i++) {
        if (areal <= alternativer[i].maxAreal || alternativer[i].maxAreal === Infinity) {
          pris = alternativer[i].prisStandard;
          valgtProdNr = alternativer[i].prodNr;
          break;
        }
      }
    }
  }

  // Legger til prisen for markedsverdi og produktnummer hvis aktivert
  if (inkluderMarked && boligtype !== 'Frittstående bygg') {
    pris += (prisliste.markedsverdi ? prisliste.markedsverdi.pris : 2000);
    const markedsNr = prisliste.markedsverdi ? prisliste.markedsverdi.prodNr : '9';
    valgtProdNr = valgtProdNr ? valgtProdNr + ', ' + markedsNr : markedsNr;
  }

  if (tilleggsbygg > 0) pris += (tilleggsbygg * 1250);

  if (pris > 0) {
    const prisEks = Math.round(pris / (1 + CONFIG.MVA_RATE));
    sheet.getRange(row, COL.PRIS_INKL).setValue(pris);
    sheet.getRange(row, COL.PRIS_EKS).setValue(prisEks);
    sheet.getRange(row, COL.MVA_BELOP).setValue(pris - prisEks);
    sheet.getRange(row, COL.PRODUKTNUMMER).setValue(valgtProdNr);
  } else {
    sheet.getRange(row, COL.PRIS_INKL).setValue('');
    sheet.getRange(row, COL.PRIS_EKS).setValue('');
    sheet.getRange(row, COL.MVA_BELOP).setValue('');
    sheet.getRange(row, COL.PRODUKTNUMMER).setValue('');
  }
}


// ============================================================
// 16. SEND FAKTURA (VIA KNAPP)
// ============================================================
function sendFakturaTilRegnskap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  let count = 0;

  for (let i = data.length - 1; i >= 1; i--) {
    const rowNum = i + 1;
    const rowData = data[i];
    const isChecked = rowData[COL.KAN_FAKTURERES - 1] === true;
    const status = rowData[COL.STATUS - 1];

    // Utløses hvis boksen er krysset av ELLER hvis status er 'Kan faktureres'
    if (isChecked || status === 'Kan faktureres') {
      const oppdragsnr = rowData[COL.OPPDRAGSNR - 1];
      const adresse = rowData[COL.ADRESSE - 1];
      const datoStr = Utilities.formatDate(new Date(), 'Europe/Oslo', 'dd.MM.yyyy HH:mm');

      // 1. Send e-post
      const html = buildFakturaEmail_(rowData, datoStr);
      safeSendEmail_(
        CONFIG.ACCOUNTANT_EMAIL,
        '💰 Klar til fakturering: ' + adresse + ' (' + oppdragsnr + ')',
        html
      );

      // 2. Fjern kryss og sett status til Fakturert
      sheet.getRange(rowNum, COL.KAN_FAKTURERES).setValue(false);
      sheet.getRange(rowNum, COL.STATUS).setValue('Fakturert');

      // Valgfri backup til arkivet (beholder raden i hovedarket)
      archiveRow_(sheet, rowNum);

      count++;
    }
  }

  if (count > 0) {
    SpreadsheetApp.getUi().alert('✅ Suksess! ' + count + ' oppdrag ble sendt til regnskap.\nStatus ble automatisk endret til "Fakturert".');
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ Ingen rader var klare for fakturering.');
  }
}

// Funksjon for å skjule rader med status "Oppdrag fullført"
function skjulFullforteOppdrag() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const statusIndex = COL.STATUS - 1; // 0-basert index for array

  sheet.showRows(1, sheet.getMaxRows()); // Viser alt først for å unngå rot

  for (let i = 1; i < data.length; i++) {
    const status = data[i][statusIndex];
    if (status === "Oppdrag fullført") {
      sheet.hideRows(i + 1);
    }
  }
}

// Funksjon for å vise alle rader (hvis du trenger å se de fullførte igjen)
function visAlleOppdrag() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  sheet.showRows(1, sheet.getMaxRows());
}

function sorterEtterStatus() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2) return; // Ingen data å sortere

  // Hent alle data (unntatt header)
  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const data = range.getValues();

  // Sorteringslogikk
  data.sort(function(a, b) {
    const statusA = a[COL.STATUS - 1];
    const statusB = b[COL.STATUS - 1];
    
    const vektA = CONFIG.STATUS_PRIORITY[statusA] || 99; // 99 for ukjente statuser
    const vektB = CONFIG.STATUS_PRIORITY[statusB] || 99;
    
    // Sorter primært på status-vekt
    if (vektA !== vektB) {
      return vektA - vektB;
    }
    
    // Sekundært: Sorter på dato mottatt (nyeste øverst innenfor samme status)
    const datoA = new Date(a[COL.DATO_MOTTATT - 1]);
    const datoB = new Date(b[COL.DATO_MOTTATT - 1]);
    return datoB - datoA;
  });

  // Skriv sortert data tilbake
  range.setValues(data);
  SpreadsheetApp.getActiveSpreadsheet().toast("Oppdragsloggen er sortert etter arbeidsflyt.", "Sortering ferdig");
}