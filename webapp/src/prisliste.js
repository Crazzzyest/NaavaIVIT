const config = require('./config');
const google = require('./google');

const SHEET_NAME = 'Prisliste';
const HEADER = [
  'Kategori',
  'Størrelse',
  'Tilstandsrapport Pris (inkl. mva)',
  'Tilstandsrapport med teknisk- og markedsverdi (inkl. mva)',
  'Produktnummer',
];

// Default seed values from user spec.
const DEFAULTS = [
  { kategori: 'Leilighet',                                          storrelse: 'Under 80 m²',                    prisStandard: 10000, prisMarked: 12000, prodNr: '1' },
  { kategori: 'Leilighet',                                          storrelse: 'Over 80 m²',                     prisStandard: 12000, prisMarked: 14000, prodNr: '2' },
  { kategori: 'Rekkehus eller leilighet i 2-, 3- og 4-mannsbolig',  storrelse: 'Under 80 m²',                    prisStandard: 12000, prisMarked: 14000, prodNr: '3' },
  { kategori: 'Rekkehus eller leilighet i 2-, 3- og 4-mannsbolig',  storrelse: 'Over 80 m²',                     prisStandard: 14000, prisMarked: 16000, prodNr: '4' },
  { kategori: 'Enebolig eller fritidsbolig',                        storrelse: 'Under 150 m²',                   prisStandard: 16000, prisMarked: 18000, prodNr: '5' },
  { kategori: 'Enebolig eller fritidsbolig',                        storrelse: '150 m² - 250 m²',                prisStandard: 18000, prisMarked: 20000, prodNr: '6' },
  { kategori: 'Enebolig eller fritidsbolig',                        storrelse: 'Over 250 m²',                    prisStandard: 20000, prisMarked: 22000, prodNr: '7' },
  { kategori: 'Frittstående bygg',                                  storrelse: 'Garasje, carport, bod, anneks',  prisStandard: 1250,  prisMarked: 1250,  prodNr: '8' },
  { kategori: 'Timesats',                                           storrelse: '',                               prisStandard: 1500,  prisMarked: 1500,  prodNr: ''  },
  { kategori: 'Markedsverdi',                                       storrelse: '',                               prisStandard: 2000,  prisMarked: 0,     prodNr: '9' },
];

function parseSheetNumber(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  let s = String(v).replace(/[^\d.,-]/g, '').replace(/\s/g, '');
  if (!s) return 0;
  const hasComma = s.includes(',');
  const hasDot = s.includes('.');
  if (hasComma && hasDot) {
    s = s.replace(/,/g, '');
  } else if (hasComma) {
    const parts = s.split(',');
    const isThousand = parts.length > 1 && parts[parts.length - 1].length === 3 && parts.every((p, i) => i === 0 || p.length === 3);
    s = isThousand ? s.replace(/,/g, '') : s.replace(',', '.');
  } else if (hasDot) {
    const parts = s.split('.');
    const isThousand = parts.length > 1 && parts[parts.length - 1].length === 3 && parts.every((p, i) => i === 0 || p.length === 3);
    if (isThousand) s = s.replace(/\./g, '');
  }
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
}

async function getRows() {
  if (config.demoMode) return DEFAULTS.map(r => ({ ...r }));
  try {
    const data = await google.getSheetData(SHEET_NAME);
    if (!data || data.length < 2) return DEFAULTS.map(r => ({ ...r }));
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i] || [];
      const kategori = String(r[0] || '').trim();
      const storrelse = String(r[1] || '').trim();
      if (!kategori && !storrelse) continue; // skip blank rows
      rows.push({
        kategori,
        storrelse,
        prisStandard: parseSheetNumber(r[2]),
        prisMarked: parseSheetNumber(r[3]),
        prodNr: String(r[4] || '').trim(),
      });
    }
    return rows.length > 0 ? rows : DEFAULTS.map(r => ({ ...r }));
  } catch (e) {
    console.warn(`Prisliste sheet not accessible: ${e.message}. Returning defaults.`);
    return DEFAULTS.map(r => ({ ...r }));
  }
}

async function setRows(rows) {
  if (config.demoMode) {
    return { demoMode: true, updated: rows.length };
  }
  if (!Array.isArray(rows)) throw new Error('rows must be an array');

  // Build sheet data: header + each row
  const data = [HEADER];
  for (const r of rows) {
    data.push([
      String(r.kategori || ''),
      String(r.storrelse || ''),
      Number(r.prisStandard) || 0,
      Number(r.prisMarked) || 0,
      String(r.prodNr || ''),
    ]);
  }

  // Clear existing content (200 rows × 5 cols is plenty), then write fresh
  try {
    await google.clearRange(SHEET_NAME, 1, 200, 5);
  } catch (e) {
    console.warn(`Could not clear Prisliste sheet: ${e.message}`);
  }
  await google.writeRange(SHEET_NAME, 1, 1, data);

  return { updated: rows.length };
}

module.exports = { getRows, setRows, DEFAULTS, HEADER, SHEET_NAME };
