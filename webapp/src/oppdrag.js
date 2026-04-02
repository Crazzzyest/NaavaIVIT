const config = require('./config');
const { COL, NUM_COLS } = require('./columns');
const google = require('./google');
const { calculateDistance } = require('./ai');
const { formatDate, formatCurrency, extractAddressFromHyperlink } = require('./utils');

// ============================================================
// TRAVEL COST
// ============================================================

function calculateTravelCost(kmTurRetur, sumFergeBomInklMva = 0) {
  const fakturerbarKm = Math.max(0, kmTurRetur - config.travel.inkludertKm);
  const kmKostnadEks = fakturerbarKm * config.travel.satsEksMva;
  const kmKostnadInkl = kmKostnadEks * (1 + config.mvaRate);
  const bomInkl = sumFergeBomInklMva;
  const bomEks = bomInkl / (1 + config.mvaRate);

  return {
    totalKm: kmTurRetur,
    fakturerbarKm,
    kostnadEksMva: Math.round(kmKostnadEks + bomEks),
    kostnadInklMva: Math.round(kmKostnadInkl + bomInkl),
  };
}

// ============================================================
// PRICE CALCULATION
// ============================================================

const PRISLISTE_MED_MARKED = {
  'Leilighet': [
    { maxAreal: 80, pris: 12000 },
    { maxAreal: Infinity, pris: 14000 },
  ],
  'Rekkehus/leilighet 2-4-mannsbolig': [
    { maxAreal: 80, pris: 14000 },
    { maxAreal: Infinity, pris: 16000 },
  ],
  'Enebolig/fritidsbolig': [
    { maxAreal: 150, pris: 18000 },
    { maxAreal: 250, pris: 20000 },
    { maxAreal: Infinity, pris: 22000 },
  ],
  'Frittstående bygg': [
    { maxAreal: Infinity, pris: 1250 },
  ],
};

const PRISLISTE_UTEN_MARKED = {
  'Leilighet': [
    { maxAreal: 80, pris: 10000 },
    { maxAreal: Infinity, pris: 12000 },
  ],
  'Rekkehus/leilighet 2-4-mannsbolig': [
    { maxAreal: 80, pris: 12000 },
    { maxAreal: Infinity, pris: 14000 },
  ],
  'Enebolig/fritidsbolig': [
    { maxAreal: 150, pris: 16000 },
    { maxAreal: 250, pris: 18000 },
    { maxAreal: Infinity, pris: 20000 },
  ],
  'Frittstående bygg': [
    { maxAreal: Infinity, pris: 1250 },
  ],
};

async function getPricesFromSheet() {
  try {
    const data = await google.getSheetData('Prisliste');
    if (!data || data.length < 2) return null;

    const priser = {
      timesats: { pris: 1500, prodNr: '' },
      markedsverdi: { pris: 2000, prodNr: '9' },
    };

    for (let i = 1; i < data.length; i++) {
      const kategoriRaw = String(data[i][0] || '').trim();
      const strRaw = String(data[i][1] || '').trim();
      const prisStandard = parseFloat(String(data[i][2] || '').replace(/\s/g, '')) || 0;
      const prisMarked = parseFloat(String(data[i][3] || '').replace(/\s/g, '')) || 0;
      const prodNr = String(data[i][4] || '').trim();

      if (!kategoriRaw) continue;

      if (kategoriRaw.toLowerCase().includes('timesats')) {
        priser.timesats = { pris: prisStandard, prodNr };
        continue;
      }
      if (kategoriRaw.toLowerCase().includes('markedsverdi')) {
        priser.markedsverdi = { pris: prisStandard, prodNr };
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
      priser[kategori].push({ maxAreal, prisStandard, prisMarked, prodNr });
    }

    for (const key in priser) {
      if (Array.isArray(priser[key])) {
        priser[key].sort((a, b) => a.maxAreal - b.maxAreal);
      }
    }

    return priser;
  } catch {
    return null;
  }
}

function calculatePriceFromData(boligtype, areal, tilleggsbygg, inkluderMarked, timer, prisliste) {
  if (!boligtype && !timer) return null;

  const arealNum = areal ? parseFloat(areal) : 0;
  const tillegg = tilleggsbygg ? parseInt(tilleggsbygg, 10) : 0;
  const timerNum = timer ? parseFloat(String(timer).replace(',', '.')) : 0;

  let pris = 0;
  let valgtProdNr = '';

  if (prisliste) {
    // Use sheet-based pricing
    if (timerNum > 0 && !boligtype) {
      pris = timerNum * prisliste.timesats.pris;
      valgtProdNr = prisliste.timesats.prodNr;
    } else if (boligtype && prisliste[boligtype]) {
      const alt = prisliste[boligtype];
      if (boligtype === 'Frittstående bygg') {
        pris = inkluderMarked ? alt[0].prisMarked : alt[0].prisStandard;
        valgtProdNr = alt[0].prodNr;
      } else {
        for (const a of alt) {
          if (arealNum <= a.maxAreal || a.maxAreal === Infinity) {
            pris = inkluderMarked ? a.prisMarked : a.prisStandard;
            valgtProdNr = a.prodNr;
            break;
          }
        }
      }
    }

    if (inkluderMarked && boligtype !== 'Frittstående bygg') {
      const markedNr = prisliste.markedsverdi ? prisliste.markedsverdi.prodNr : '9';
      valgtProdNr = valgtProdNr ? `${valgtProdNr}, ${markedNr}` : markedNr;
    }
  } else {
    // Fallback to hardcoded pricing
    const liste = inkluderMarked ? PRISLISTE_MED_MARKED : PRISLISTE_UTEN_MARKED;
    if (boligtype && liste[boligtype]) {
      for (const entry of liste[boligtype]) {
        if (arealNum <= entry.maxAreal) {
          pris = entry.pris;
          break;
        }
      }
    }
  }

  if (tillegg > 0) pris += tillegg * 1250;
  if (pris <= 0) return null;

  const prisEks = Math.round(pris / (1 + config.mvaRate));
  return {
    prisInkl: pris,
    prisEks,
    mvaBeløp: pris - prisEks,
    produktnummer: valgtProdNr,
  };
}

async function calculatePriceForRow(row) {
  const data = await google.getSheetData(config.sheet.name);
  if (!data || row < 2 || row > data.length) return;

  const rowData = data[row - 1];
  const boligtype = rowData[COL.BOLIGTYPE - 1] || '';
  const areal = rowData[COL.AREAL - 1] || '';
  const tillegg = rowData[COL.ANTALL_TILLEGGSBYGG - 1] || '';
  const inkluderMarked = rowData[COL.MED_MARKEDSVERDI - 1] === true || rowData[COL.MED_MARKEDSVERDI - 1] === 'TRUE';
  const timer = rowData[COL.TIMER - 1] || '';

  const prisliste = await getPricesFromSheet();
  const result = calculatePriceFromData(boligtype, areal, tillegg, inkluderMarked, timer, prisliste);

  if (result) {
    await google.updateCells(config.sheet.name, row, [
      { col: COL.PRIS_INKL, value: result.prisInkl },
      { col: COL.PRIS_EKS, value: result.prisEks },
      { col: COL.MVA_BELOP, value: result.mvaBeløp },
      { col: COL.PRODUKTNUMMER, value: result.produktnummer },
    ]);
  } else {
    await google.updateCells(config.sheet.name, row, [
      { col: COL.PRIS_INKL, value: '' },
      { col: COL.PRIS_EKS, value: '' },
      { col: COL.MVA_BELOP, value: '' },
      { col: COL.PRODUKTNUMMER, value: '' },
    ]);
  }
}

// ============================================================
// CREATE NEW OPPDRAG
// ============================================================

async function createNewOppdrag(parsed, date) {
  try {
    console.log(`>>> createNewOppdrag: ${parsed.adresse}`);

    const data = await google.getSheetData(config.sheet.name);
    const lastRow = data.length;

    // Duplicate check
    if (lastRow > 1) {
      const incomingAddrBase = (parsed.adresse || '').split(',')[0].trim().toLowerCase();

      for (let i = 1; i < data.length; i++) {
        const rowAddrRaw = String(data[i][COL.ADRESSE - 1] || '');
        const rowTimestamp = data[i][COL.TIMESTAMP - 1];
        if (!rowTimestamp) continue;

        const rowDate = new Date(rowTimestamp);

        if (rowAddrRaw === parsed.adresse && rowDate.toDateString() === date.toDateString()) {
          console.log('  Exact duplicate same day, aborting');
          return null;
        }
      }
    }

    const oppdragsnr = 'NT-' + formatDate(date, 'yyyyMM') + '-' + String(lastRow).padStart(3, '0');
    console.log(`  Oppdragsnr: ${oppdragsnr}`);

    let folderUrl = '';
    try {
      const folderName = `${parsed.adresse || 'Ukjent'} - ${formatDate(date, 'dd.MM.yyyy')}`;
      const folder = await google.createOppdragFolder(folderName);
      folderUrl = folder.url;
      console.log(`  Folder created: ${folderUrl}`);
    } catch (e) {
      console.error('  WARNING: Could not create folder:', e.message);
    }

    const datoMottatt = formatDate(date, 'dd.MM.yyyy HH:mm');
    const timestamp = date.toISOString();

    // Travel calculation
    let avstandKm = '';
    let reiseEks = 0;
    let reiseInkl = 0;
    let reiseNotat = '';

    if (parsed.adresse) {
      const distResult = await calculateDistance(parsed.adresse);
      if (distResult) {
        const travelCost = calculateTravelCost(distResult.kmTurRetur);
        avstandKm = distResult.kmTurRetur;
        reiseEks = travelCost.kostnadEksMva;
        reiseInkl = travelCost.kostnadInklMva;
        reiseNotat = `${distResult.kmTurRetur} km t/r` +
          (travelCost.fakturerbarKm > 0
            ? `, ${travelCost.fakturerbarKm} km fakturerbart`
            : ' (inkludert)');
      }
    }

    let fullNotater = '';
    if (parsed.notater) fullNotater += parsed.notater;
    if (reiseNotat) fullNotater += (fullNotater ? ' | ' : '') + 'Reise: ' + reiseNotat;

    // Build row (NUM_COLS elements)
    const newRow = new Array(NUM_COLS).fill('');
    newRow[COL.OPPDRAGSNR - 1] = oppdragsnr;
    newRow[COL.DATO_MOTTATT - 1] = datoMottatt;
    newRow[COL.KILDE - 1] = parsed.kilde;
    newRow[COL.OPPDRAGSTYPE - 1] = parsed.oppdragstype;
    newRow[COL.ADRESSE - 1] = folderUrl
      ? `=HYPERLINK("${folderUrl}","${(parsed.adresse || '').replace(/"/g, '""')}")`
      : parsed.adresse;
    newRow[COL.OPPDRAGSGIVER - 1] = parsed.oppdragsgiver;
    newRow[COL.SELGER - 1] = parsed.selger;
    newRow[COL.SELGER_TLF - 1] = parsed.selgerTlf;
    newRow[COL.SELGER_EPOST - 1] = parsed.selgerEpost;
    newRow[COL.MEGLER - 1] = parsed.megler;
    newRow[COL.MEGLER_EPOST - 1] = parsed.meglerEpost;
    newRow[COL.FAKTURA_REF - 1] = parsed.fakturaRef;
    newRow[COL.FAKTURA_SENDES_TIL - 1] = parsed.fakturaSendesTil;
    newRow[COL.RAPPORTTYPE - 1] = parsed.rapporttype || '';
    newRow[COL.MED_MARKEDSVERDI - 1] = parsed.rapporttype === 'Tilstandsrapport m/teknisk og markedsverdi';
    newRow[COL.AVSTAND_KM - 1] = avstandKm;
    newRow[COL.REISE_EKS - 1] = reiseEks;
    newRow[COL.REISE_INKL - 1] = reiseInkl;
    newRow[COL.STATUS - 1] = 'Mottatt';
    newRow[COL.DATO_STATUSENDRING - 1] = datoMottatt;
    newRow[COL.TIMESTAMP - 1] = timestamp;
    newRow[COL.LINK_MAPPE - 1] = folderUrl;
    newRow[COL.NOTATER - 1] = fullNotater;

    await google.appendRow(config.sheet.name, newRow);
    console.log('  Row written');

    return oppdragsnr;

  } catch (err) {
    console.error('ERROR in createNewOppdrag:', err.message);
    return null;
  }
}

// ============================================================
// STATUS CHANGE HANDLING
// ============================================================

async function handleStatusChange(row, newStatus) {
  const now = new Date();
  const datoStr = formatDate(now, 'dd.MM.yyyy HH:mm');

  await google.updateCell(config.sheet.name, row, COL.DATO_STATUSENDRING, datoStr);

  const data = await google.getSheetData(config.sheet.name);
  const rowData = data[row - 1];

  if (newStatus === 'Kan faktureres') {
    const { buildFakturaEmail } = require('./emails');
    const adresse = rowData[COL.ADRESSE - 1];
    const oppdragsnr = rowData[COL.OPPDRAGSNR - 1];

    await google.sendEmail(
      config.email.accountantEmail,
      `Klar til fakturering: ${adresse} (${oppdragsnr})`,
      buildFakturaEmail(rowData, datoStr),
    );

    await google.updateCells(config.sheet.name, row, [
      { col: COL.STATUS, value: 'Fakturert' },
      { col: COL.KAN_FAKTURERES, value: false },
    ]);

    await archiveRow(rowData);
  }

  if (newStatus === 'Oppdrag fullført') {
    const { extractDriveFolderId } = require('./utils');
    const folderUrl = rowData[COL.LINK_MAPPE - 1];
    const folderId = extractDriveFolderId(folderUrl);
    if (folderId) {
      try {
        const targetId = await google.getOrCreateAvsluttedeFolder();
        await google.moveFolder(folderId, targetId);
      } catch (e) {
        console.error('Move folder failed:', e.message);
      }
    }
  }
}

async function archiveRow(rowData) {
  try {
    const archiveData = await google.getSheetData('arkiv');
    if (!archiveData || archiveData.length === 0) {
      const headerData = await google.getSheetData(config.sheet.name);
      if (headerData.length > 0) {
        await google.appendRow('arkiv', headerData[0]);
      }
    }

    const oppdragsnr = rowData[COL.OPPDRAGSNR - 1];
    if (archiveData && archiveData.length >= 2) {
      const existing = archiveData.slice(1).map(r => r[0]);
      if (existing.includes(oppdragsnr)) {
        console.log(`Archive: Duplicate ${oppdragsnr}, skipping`);
        return;
      }
    }

    await google.appendRow('arkiv', rowData);
    console.log(`Archived: ${oppdragsnr}`);
  } catch (e) {
    console.error('Archive error:', e.message);
  }
}

// ============================================================
// BEFARING HANDLING
// ============================================================

async function handleBefaringBooked(row) {
  const data = await google.getSheetData(config.sheet.name);
  const rowData = data[row - 1];
  const status = rowData[COL.STATUS - 1];

  if (status === 'Mottatt') {
    const now = new Date();
    await google.updateCells(config.sheet.name, row, [
      { col: COL.STATUS, value: 'Avtalt befaring' },
      { col: COL.DATO_STATUSENDRING, value: formatDate(now, 'dd.MM.yyyy HH:mm') },
    ]);
  }
}

// ============================================================
// RECALCULATE TRAVEL
// ============================================================

async function recalculateTravel(row, address) {
  const distResult = await calculateDistance(address);
  if (!distResult) return;

  const data = await google.getSheetData(config.sheet.name);
  const rowData = data[row - 1];
  const bomStr = String(rowData[COL.SUM_FERGE_BOM - 1] || '0').replace(/[^\d.,-]/g, '').replace(',', '.');
  const bom = parseFloat(bomStr) || 0;

  const travelCost = calculateTravelCost(distResult.kmTurRetur, bom);
  await google.updateCells(config.sheet.name, row, [
    { col: COL.AVSTAND_KM, value: distResult.kmTurRetur },
    { col: COL.REISE_EKS, value: travelCost.kostnadEksMva },
    { col: COL.REISE_INKL, value: travelCost.kostnadInklMva },
  ]);
}

// ============================================================
// SEND FAKTURA (batch)
// ============================================================

async function sendFakturaTilRegnskap() {
  const { buildFakturaEmail } = require('./emails');
  const data = await google.getSheetData(config.sheet.name);
  let count = 0;

  for (let i = data.length - 1; i >= 1; i--) {
    const rowData = data[i];
    const isChecked = rowData[COL.KAN_FAKTURERES - 1] === true || rowData[COL.KAN_FAKTURERES - 1] === 'TRUE';
    const status = rowData[COL.STATUS - 1];
    const rowNum = i + 1;

    if (isChecked || status === 'Kan faktureres') {
      const oppdragsnr = rowData[COL.OPPDRAGSNR - 1];
      const adresse = rowData[COL.ADRESSE - 1];
      const datoStr = formatDate(new Date(), 'dd.MM.yyyy HH:mm');

      const html = buildFakturaEmail(rowData, datoStr);
      await google.sendEmail(
        config.email.accountantEmail,
        `Klar til fakturering: ${adresse} (${oppdragsnr})`,
        html,
      );

      await google.updateCells(config.sheet.name, rowNum, [
        { col: COL.KAN_FAKTURERES, value: false },
        { col: COL.STATUS, value: 'Fakturert' },
      ]);

      await archiveRow(rowData);
      count++;
    }
  }

  return count;
}

// ============================================================
// MANUAL REGISTRATION
// ============================================================

async function registerManualOppdrag(data) {
  const mapTypes = (rapporttype) => {
    const rt = String(rapporttype || '').toLowerCase();
    if (rt === 'tilstandsrapport') return { oppdragstype: 'Tilstandsrapport', rapporttype: 'Tilstandsrapport' };
    if (rt.includes('markedsverdi')) return { oppdragstype: 'Tilstandsrapport m/markedsverdi', rapporttype: 'Tilstandsrapport m/teknisk og markedsverdi' };
    if (rt.includes('skadetakst')) return { oppdragstype: 'Skadetakst', rapporttype: 'Skadetakstrapport' };
    if (rt.includes('reklamasjon')) return { oppdragstype: 'Reklamasjon', rapporttype: 'Reklamasjonsrapport' };
    if (rt.includes('forhåndstakst')) return { oppdragstype: 'Forhåndstakst', rapporttype: 'Annen rapport' };
    return { oppdragstype: 'Annet', rapporttype: 'Annen rapport' };
  };

  const mapped = mapTypes(data.rapporttype);

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
    rapporttype: mapped.rapporttype,
  };

  if (config.testMode) {
    const { sanitizeParsedData } = require('./utils');
    parsed = sanitizeParsedData(parsed);
  }

  const oppdragsnr = await createNewOppdrag(parsed, new Date());

  // Set boligtype/areal and calculate price
  if (oppdragsnr) {
    const allData = await google.getSheetData(config.sheet.name);
    const row = allData.length; // last row (just appended)

    const updates = [];
    if (data.boligtype) updates.push({ col: COL.BOLIGTYPE, value: data.boligtype });
    if (data.areal) updates.push({ col: COL.AREAL, value: Number(data.areal) });
    if (updates.length > 0) {
      await google.updateCells(config.sheet.name, row, updates);
    }

    await calculatePriceForRow(row);
  }

  return oppdragsnr;
}

module.exports = {
  createNewOppdrag,
  calculateTravelCost,
  calculatePriceForRow,
  handleStatusChange,
  handleBefaringBooked,
  recalculateTravel,
  sendFakturaTilRegnskap,
  registerManualOppdrag,
  getPricesFromSheet,
};
