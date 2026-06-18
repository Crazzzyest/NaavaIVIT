const config = require('./config');
const google = require('./google');
const { COL } = require('./columns');
const { parseDateString } = require('./utils');

// 5-second in-process cache to protect Sheets API quota during burst loads
let cache = { ts: 0, rows: null };
const CACHE_TTL_MS = 5000;

// Sheet cells come back as locale-formatted strings ("16 000 kr", "12,800 kr",
// "0,25", "0.25", or plain numbers). This extracts the numeric value robustly,
// distinguishing comma-as-thousand-separator from comma-as-decimal-mark.
function parseSheetNumber(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return v;
  let s = String(v).replace(/[^\d.,-]/g, '').replace(/\s/g, '');
  if (!s) return null;

  const hasComma = s.includes(',');
  const hasDot = s.includes('.');

  if (hasComma && hasDot) {
    // Both present — assume comma is thousand separator (US format: "12,800.50")
    s = s.replace(/,/g, '');
  } else if (hasComma) {
    // Comma only — distinguish "12,800" (thousand sep) from "0,25" (decimal)
    const parts = s.split(',');
    const isThousandPattern =
      parts.length > 1 &&
      parts[parts.length - 1].length === 3 &&
      parts.every((p, i) => i === 0 || p.length === 3);
    s = isThousandPattern ? s.replace(/,/g, '') : s.replace(',', '.');
  } else if (hasDot) {
    // Dot only — distinguish "16.000" (thousand sep, NO common in Norway but Sheets API may emit it) from "0.25"
    const parts = s.split('.');
    const isThousandPattern =
      parts.length > 1 &&
      parts[parts.length - 1].length === 3 &&
      parts.every((p, i) => i === 0 || p.length === 3);
    if (isThousandPattern) s = s.replace(/\./g, '');
  }

  const n = parseFloat(s);
  return Number.isFinite(n) ? n : null;
}

function bustCache() {
  cache = { ts: 0, rows: null };
}

async function getRowsCached() {
  const now = Date.now();
  if (cache.rows && now - cache.ts < CACHE_TTL_MS) return cache.rows;
  const rows = await google.getSheetData(config.sheet.name);
  cache = { ts: now, rows };
  return rows;
}

// Text columns can pick up stray checkbox values ("TRUE"/"FALSE") when columns
// are inserted next to a checkbox column in the sheet. Treat those as empty.
function cleanFreeText(v) {
  const s = String(v == null ? '' : v).trim();
  if (s === 'TRUE' || s === 'FALSE') return '';
  return s;
}

function projectRow(row, rowNum) {
  return {
    rowNum,
    kanFaktureres: row[COL.KAN_FAKTURERES - 1] === true || row[COL.KAN_FAKTURERES - 1] === 'TRUE',
    oppdragsnr: row[COL.OPPDRAGSNR - 1] || '',
    datoMottatt: row[COL.DATO_MOTTATT - 1] || '',
    kilde: row[COL.KILDE - 1] || '',
    scanIvit: row[COL.SCAN_IVIT - 1] === true || row[COL.SCAN_IVIT - 1] === 'TRUE',
    oppdragstype: row[COL.OPPDRAGSTYPE - 1] || '',
    adresse: row[COL.ADRESSE - 1] || '',
    oppdragsgiver: row[COL.OPPDRAGSGIVER - 1] || '',
    selger: row[COL.SELGER - 1] || '',
    selgerTlf: row[COL.SELGER_TLF - 1] || '',
    selgerEpost: row[COL.SELGER_EPOST - 1] || '',
    megler: row[COL.MEGLER - 1] || '',
    meglerEpost: row[COL.MEGLER_EPOST - 1] || '',
    fakturaRef: row[COL.FAKTURA_REF - 1] || '',
    status: row[COL.STATUS - 1] || '',
    fakturaSendesTil: row[COL.FAKTURA_SENDES_TIL - 1] || '',
    fakturamotaker: row[COL.FAKTURAMOTAKER - 1] || '',
    boligtype: row[COL.BOLIGTYPE - 1] || '',
    areal: row[COL.AREAL - 1] || '',
    antallTilleggsbygg: row[COL.ANTALL_TILLEGGSBYGG - 1] || '',
    rapporttype: row[COL.RAPPORTTYPE - 1] || '',
    medMarkedsverdi: row[COL.MED_MARKEDSVERDI - 1] === true || row[COL.MED_MARKEDSVERDI - 1] === 'TRUE',
    timer: row[COL.TIMER - 1] || '',
    prisInkl: parseSheetNumber(row[COL.PRIS_INKL - 1]),
    prisEks: parseSheetNumber(row[COL.PRIS_EKS - 1]),
    mvaBelop: parseSheetNumber(row[COL.MVA_BELOP - 1]),
    avstandKm: row[COL.AVSTAND_KM - 1] || '',
    reiseEks: parseSheetNumber(row[COL.REISE_EKS - 1]),
    reiseInkl: parseSheetNumber(row[COL.REISE_INKL - 1]),
    sumFergeBom: row[COL.SUM_FERGE_BOM - 1] || '',
    antallDeleReise: row[COL.ANTALL_DELE_REISE - 1] || '',
    befaringDato: row[COL.BEFARING_DATO - 1] || '',
    befaringKl: row[COL.BEFARING_KL - 1] || '',
    datoStatusendring: row[COL.DATO_STATUSENDRING - 1] || '',
    timestamp: row[COL.TIMESTAMP - 1] || '',
    linkMappe: row[COL.LINK_MAPPE - 1] || '',
    notater: row[COL.NOTATER - 1] || '',
    produktnummer: row[COL.PRODUKTNUMMER - 1] || '',
    kommentarRegnskap: row[COL.KOMMENTAR_REGNSKAP - 1] || '',
    kansellert: row[COL.KANSELLERT - 1] === true || row[COL.KANSELLERT - 1] === 'TRUE',
    fakturamotakerAdresse: cleanFreeText(row[COL.FAKTURAMOTAKER_ADRESSE - 1]),
    fakturamotakerEpost: cleanFreeText(row[COL.FAKTURAMOTAKER_EPOST - 1]),
    fakturamotakerInfo: cleanFreeText(row[COL.FAKTURAMOTAKER_INFO - 1]),
  };
}

async function listOppdrag({ status, type, q } = {}) {
  const rows = await getRowsCached();
  if (!rows || rows.length < 2) return [];
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const o = projectRow(rows[i], i + 1);
    if (status && o.status !== status) continue;
    if (type && o.oppdragstype !== type) continue;
    if (q) {
      const ql = q.toLowerCase();
      const hay = [o.adresse, o.oppdragsnr, o.megler, o.selger, o.oppdragstype].join(' ').toLowerCase();
      if (!hay.includes(ql)) continue;
    }
    out.push(o);
  }
  return out;
}

async function getOppdrag(rowNum) {
  if (config.demoMode) {
    const { getDemoOppdrag } = require('./demo-data');
    return getDemoOppdrag(rowNum);
  }
  const rows = await getRowsCached();
  if (!rows || rowNum < 2 || rowNum > rows.length) return null;
  return projectRow(rows[rowNum - 1], rowNum);
}

async function getDashboardData() {
  if (config.demoMode) {
    const { getDemoDashboardData } = require('./demo-data');
    return getDemoDashboardData();
  }
  const rows = await getRowsCached();
  if (!rows || rows.length < 2) {
    return {
      success: true,
      testMode: config.testMode,
      demoMode: config.demoMode,
      total: 0,
      active: 0,
      oppdrag: [],
      statusCounts: {},
    };
  }

  const now = new Date();
  const curMonth = now.getMonth();
  const curYear = now.getFullYear();
  const statusCounts = {};
  const oppdrag = [];
  let utestaendeInkl = 0;
  let omsMaaned = 0;
  let omsAar = 0;
  let sumPris = 0;
  let countPris = 0;

  for (let i = 1; i < rows.length; i++) {
    const o = projectRow(rows[i], i + 1);
    oppdrag.push(o);

    statusCounts[o.status] = (statusCounts[o.status] || 0) + 1;

    const prisInkl = Number(o.prisInkl) || 0;
    const reiseInkl = Number(o.reiseInkl) || 0;

    if (o.status === 'Kan faktureres') {
      utestaendeInkl += prisInkl + reiseInkl;
    }

    // Omsetning teller alle oppdrag som har gått gjennom systemet (har pris),
    // unntatt kansellerte. Inkluderer Befart/Utkast/Endelig rapport/Kan faktureres/
    // Fakturert/Oppdrag fullført.
    const isCancelled = o.status === 'Oppdrag kansellert';
    if (!isCancelled && prisInkl > 0) {
      omsAar += prisInkl;
      const sDato = parseDateString(o.datoStatusendring) || parseDateString(o.datoMottatt);
      if (sDato && sDato.getMonth() === curMonth && sDato.getFullYear() === curYear) {
        omsMaaned += prisInkl;
      }
      sumPris += prisInkl;
      countPris++;
    }
  }

  const total = rows.length - 1;
  const active = total
    - (statusCounts['Fakturert'] || 0)
    - (statusCounts['Oppdrag kansellert'] || 0)
    - (statusCounts['Oppdrag fullført'] || 0);

  return {
    success: true,
    testMode: config.testMode,
    demoMode: config.demoMode,
    total,
    active,
    utestaendeInkl,
    omsMaaned,
    omsAar,
    snittPrisInkl: countPris ? Math.round(sumPris / countPris) : 0,
    oppdrag,
    statusCounts,
  };
}

// Whitelist of fields the PATCH endpoint may write.
// Mapping: field name (from client) → { col: COL.X, kind: 'text'|'number'|'boolean'|'date' }
const PATCHABLE_FIELDS = {
  boligtype:           { col: COL.BOLIGTYPE,           kind: 'text' },
  areal:               { col: COL.AREAL,               kind: 'number' },
  antallTilleggsbygg:  { col: COL.ANTALL_TILLEGGSBYGG, kind: 'number' },
  medMarkedsverdi:     { col: COL.MED_MARKEDSVERDI,    kind: 'boolean' },
  rapporttype:         { col: COL.RAPPORTTYPE,         kind: 'text' },
  timer:               { col: COL.TIMER,               kind: 'text' },
  selger:              { col: COL.SELGER,              kind: 'text' },
  selgerTlf:           { col: COL.SELGER_TLF,          kind: 'text' },
  selgerEpost:         { col: COL.SELGER_EPOST,        kind: 'text' },
  megler:              { col: COL.MEGLER,              kind: 'text' },
  meglerEpost:         { col: COL.MEGLER_EPOST,        kind: 'text' },
  fakturaRef:          { col: COL.FAKTURA_REF,         kind: 'text' },
  fakturaSendesTil:    { col: COL.FAKTURA_SENDES_TIL,  kind: 'text' },
  fakturamotaker:      { col: COL.FAKTURAMOTAKER,      kind: 'text' },
  fakturamotakerAdresse: { col: COL.FAKTURAMOTAKER_ADRESSE, kind: 'text' },
  fakturamotakerEpost:   { col: COL.FAKTURAMOTAKER_EPOST,   kind: 'text' },
  fakturamotakerInfo:    { col: COL.FAKTURAMOTAKER_INFO,    kind: 'text' },
  notater:             { col: COL.NOTATER,             kind: 'text' },
  kommentarRegnskap:   { col: COL.KOMMENTAR_REGNSKAP,  kind: 'text' },
  befaringDato:        { col: COL.BEFARING_DATO,       kind: 'text' },
  befaringKl:          { col: COL.BEFARING_KL,         kind: 'text' },
  sumFergeBom:         { col: COL.SUM_FERGE_BOM,       kind: 'number' },
  antallDeleReise:     { col: COL.ANTALL_DELE_REISE,   kind: 'number' },
  scanIvit:            { col: COL.SCAN_IVIT,           kind: 'boolean' },
  kanFaktureres:       { col: COL.KAN_FAKTURERES,      kind: 'boolean' },
  // status is handled separately (routes through handleStatusChange)
};

// Fields that trigger price recalculation
const PRICING_FIELDS = new Set(['boligtype', 'areal', 'antallTilleggsbygg', 'medMarkedsverdi', 'rapporttype', 'timer']);
// Fields that trigger travel cost recalc
const TRAVEL_FIELDS = new Set(['sumFergeBom', 'antallDeleReise']);
// Fields that trigger befaring booked auto-status
const BEFARING_FIELDS = new Set(['befaringDato', 'befaringKl']);

function coerceValue(raw, kind) {
  if (raw === null || raw === undefined) return '';
  if (kind === 'boolean') {
    return raw === true || raw === 'true' || raw === 'TRUE' || raw === 1;
  }
  if (kind === 'number') {
    if (raw === '' || raw === '-') return '';
    const n = parseFloat(String(raw).replace(/\s/g, '').replace(',', '.'));
    return Number.isFinite(n) ? n : '';
  }
  return String(raw);
}

async function patchOppdrag(rowNum, partial) {
  const google = require('./google');
  const config = require('./config');

  const updates = [];
  const applied = {};
  let touchedPricing = false;
  let touchedTravel = false;
  let touchedBefaring = false;

  for (const [field, raw] of Object.entries(partial || {})) {
    const spec = PATCHABLE_FIELDS[field];
    if (!spec) continue; // ignore unknown / non-whitelisted
    const value = coerceValue(raw, spec.kind);
    updates.push({ col: spec.col, value });
    applied[field] = value;
    if (PRICING_FIELDS.has(field)) touchedPricing = true;
    if (TRAVEL_FIELDS.has(field)) touchedTravel = true;
    if (BEFARING_FIELDS.has(field)) touchedBefaring = true;
  }

  if (updates.length === 0) return { updated: applied, touched: { pricing: false, travel: false, befaring: false } };

  await google.updateCells(config.sheet.name, rowNum, updates);

  // Side effects
  if (touchedPricing) {
    const { calculatePriceForRow } = require('./oppdrag');
    await calculatePriceForRow(rowNum);
  }
  if (touchedTravel) {
    // Recalculate travel cost based on current AVSTAND_KM + new SUM_FERGE_BOM
    const rows = await google.getSheetData(config.sheet.name);
    const row = rows[rowNum - 1];
    if (row) {
      const kmStr = String(row[COL.AVSTAND_KM - 1] || '0').replace(',', '.');
      const km = parseFloat(kmStr);
      if (Number.isFinite(km) && km >= 0) {
        const { calculateTravelCost } = require('./oppdrag');
        const bomStr = String(row[COL.SUM_FERGE_BOM - 1] || '0').replace(/[^\d.,-]/g, '').replace(',', '.');
        const bom = parseFloat(bomStr) || 0;
        const tc = calculateTravelCost(km, bom);
        await google.updateCells(config.sheet.name, rowNum, [
          { col: COL.REISE_EKS, value: tc.kostnadEksMva },
          { col: COL.REISE_INKL, value: tc.kostnadInklMva },
        ]);
      }
    }
  }
  if (touchedBefaring) {
    const { handleBefaringBooked } = require('./oppdrag');
    await handleBefaringBooked(rowNum);
  }

  bustCache();
  return { updated: applied, touched: { pricing: touchedPricing, travel: touchedTravel, befaring: touchedBefaring } };
}

async function getDashboardStats() {
  if (config.demoMode) {
    const { OPPDRAG } = require('./demo-data');
    // Build a fake "raw rows" structure that matches what computeDashboardStats expects:
    // first row is header (placeholder), then 40 cells per oppdrag with values in COL positions.
    const headerRow = new Array(40).fill('');
    const rows = [headerRow];
    for (const o of OPPDRAG) {
      const r = new Array(40).fill('');
      r[COL.STATUS - 1] = o.status;
      r[COL.OPPDRAGSTYPE - 1] = o.oppdragstype;
      r[COL.BOLIGTYPE - 1] = ''; // demo doesn't set
      r[COL.AREAL - 1] = '';
      r[COL.PRIS_INKL - 1] = o.prisInkl || 0;
      r[COL.PRIS_EKS - 1] = o.prisInkl ? Math.round(o.prisInkl / 1.25) : 0;
      r[COL.REISE_INKL - 1] = o.reiseInkl || 0;
      r[COL.DATO_MOTTATT - 1] = o.datoMottatt;
      r[COL.DATO_STATUSENDRING - 1] = o.datoMottatt;
      rows.push(r);
    }
    const { computeDashboardStats } = require('./dashboard');
    return computeDashboardStats(rows, { monthsToShow: 12 });
  }
  const rows = await getRowsCached();
  const { computeDashboardStats } = require('./dashboard');
  return computeDashboardStats(rows, { monthsToShow: 12 });
}

module.exports = {
  listOppdrag,
  getOppdrag,
  getDashboardData,
  getDashboardStats,
  patchOppdrag,
  PATCHABLE_FIELDS,
  bustCache,
  projectRow,
};
