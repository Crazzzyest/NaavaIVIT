const config = require('./config');
const { COL } = require('./columns');
const google = require('./google');
const { formatCurrency, parseDateString, classifyBoligBucket, formatDate } = require('./utils');

const STATUS_LIST = ['Mottatt', 'Avtalt befaring', 'Befart', 'Utkast', 'Endelig rapport', 'Kan faktureres', 'Fakturert', 'Oppdrag kansellert', 'Oppdrag fullført'];

// ============================================================
// computeDashboardStats — pure function. Returns the stats object.
// Used by both updateDashboard() (writes to Sheet) and /api/dashboard-stats
// (returns JSON to frontend).
// ============================================================
function computeDashboardStats(data, opts = {}) {
  const monthsToShow = opts.monthsToShow || 12;
  const now = opts.now || new Date();
  const curMonth = now.getMonth();
  const curYear = now.getFullYear();

  const stats = {
    total: Math.max(0, (data?.length || 0) - 1),
    byStatus: {},
    byType: {},
    byBoligBucket: {},
    byTypeOms: {},        // omsetning per oppdragstype
    omsMåned: 0, omsÅr: 0,
    reiseMåned: 0, reiseÅr: 0,
    uteståendeInkl: 0, uteståendeEks: 0,
    snittPrisInkl: 0, snittPrisEks: 0, snittAntall: 0,
    trendMonths: [],
    omsAlle: 0,           // alle ikke-kansellerte med pris
    omsFakturert: 0,      // kun Fakturert (innkommet/realisert)
    countFakturert: 0,
    fakturertÅr: 0,
    timestamp: now.toISOString(),
  };

  STATUS_LIST.forEach(s => { stats.byStatus[s] = 0; });

  // Build last-N month buckets
  const monthMap = {};
  for (let k = monthsToShow - 1; k >= 0; k--) {
    const d = new Date(now.getFullYear(), now.getMonth() - k, 1);
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    const label = d.toLocaleDateString('nb-NO', { month: 'short', year: '2-digit' });
    const entry = { key, label, omsInkl: 0, reiseInkl: 0, totalInkl: 0, count: 0 };
    stats.trendMonths.push(entry);
    monthMap[key] = entry;
  }

  let sumPrisInkl = 0, sumPrisEks = 0, countPris = 0;

  if (!data || data.length < 2) {
    return stats;
  }

  // Robust numeric parse for sheet values like "16 000 kr" or "12,800 kr"
  function num(v) {
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

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[COL.STATUS - 1];
    const type = row[COL.OPPDRAGSTYPE - 1];
    const bolig = row[COL.BOLIGTYPE - 1];
    const areal = row[COL.AREAL - 1] ? Number(row[COL.AREAL - 1]) : 0;
    const prisInkl = num(row[COL.PRIS_INKL - 1]);
    const prisEks = num(row[COL.PRIS_EKS - 1]);
    const reiseInkl = num(row[COL.REISE_INKL - 1]);

    if (stats.byStatus[status] !== undefined) stats.byStatus[status]++;
    if (type) stats.byType[type] = (stats.byType[type] || 0) + 1;

    const bucket = classifyBoligBucket(bolig, areal);
    if (bucket) stats.byBoligBucket[bucket] = (stats.byBoligBucket[bucket] || 0) + 1;

    if (status === 'Kan faktureres') {
      stats.uteståendeInkl += (prisInkl + reiseInkl);
      stats.uteståendeEks += Math.round((prisInkl + reiseInkl) / (1 + config.mvaRate));
    }

    const isCancelled = status === 'Oppdrag kansellert';
    if (!isCancelled && prisInkl > 0) {
      stats.omsÅr += prisInkl;
      stats.omsAlle += prisInkl;
      stats.reiseÅr += reiseInkl;

      if (type) stats.byTypeOms[type] = (stats.byTypeOms[type] || 0) + prisInkl;

      const statusDato = parseDateString(row[COL.DATO_STATUSENDRING - 1]) || parseDateString(row[COL.DATO_MOTTATT - 1]);
      if (statusDato) {
        if (statusDato.getMonth() === curMonth && statusDato.getFullYear() === curYear) {
          stats.omsMåned += prisInkl;
          stats.reiseMåned += reiseInkl;
        }
        const mKey = `${statusDato.getFullYear()}-${String(statusDato.getMonth() + 1).padStart(2, '0')}`;
        if (monthMap[mKey]) {
          monthMap[mKey].omsInkl += prisInkl;
          monthMap[mKey].reiseInkl += reiseInkl;
          monthMap[mKey].totalInkl += (prisInkl + reiseInkl);
          monthMap[mKey].count++;
        }
      }

      sumPrisInkl += prisInkl;
      sumPrisEks += prisEks;
      countPris++;
    }

    if (status === 'Fakturert' && prisInkl > 0) {
      stats.omsFakturert += prisInkl;
      stats.fakturertÅr += prisInkl;
      stats.countFakturert++;
    }
  }

  stats.snittAntall = countPris;
  stats.snittPrisInkl = countPris ? Math.round(sumPrisInkl / countPris) : 0;
  stats.snittPrisEks = countPris ? Math.round(sumPrisEks / countPris) : 0;
  stats.active = stats.total
    - (stats.byStatus['Fakturert'] || 0)
    - (stats.byStatus['Oppdrag kansellert'] || 0)
    - (stats.byStatus['Oppdrag fullført'] || 0);

  return stats;
}

// ============================================================
// updateDashboard — writes a flattened table of stats into the
// Dashboard-tab in the Google Sheet (legacy mirror of Apps Script behavior).
// ============================================================
async function updateDashboard() {
  console.log('Updating dashboard...');

  const data = await google.getSheetData(config.sheet.name);
  if (!data || data.length < 2) {
    await google.writeRange(config.sheet.dashboardName, 1, 1, [['Ingen oppdrag ennå.']]);
    return;
  }

  const stats = computeDashboardStats(data, { monthsToShow: 6 });
  const now = new Date();

  const rows = [];
  rows.push(['NAAVA TAKST DASHBOARD' + (config.testMode ? ' TEST' : ''), '', '', '', '']);
  rows.push([formatDate(now, 'dd.MM.yyyy HH:mm'), '', '', '', '']);
  rows.push(['', '', '', '', '']);

  rows.push(['OVERSIKT', '', '', '', '']);
  rows.push(['Totalt antall oppdrag', stats.total, '', '', '']);
  rows.push(['Aktive (ikke Fakturert)', stats.active, '', '', '']);
  rows.push(['Utestående fordringer (Eks mva)', formatCurrency(stats.uteståendeEks), '', '', '']);
  rows.push(['Utestående fordringer (Inkl mva)', formatCurrency(stats.uteståendeInkl), '', '', '']);
  rows.push(['Gjennomsnitt pris per oppdrag (Eks mva)', formatCurrency(stats.snittPrisEks), '', '', '']);
  rows.push(['Gjennomsnitt pris per oppdrag (Inkl mva)', formatCurrency(stats.snittPrisInkl), '', '', '']);
  rows.push(['', '', '', '', '']);

  rows.push(['OMSETNING', 'Eks mva', 'Inkl mva', '', '']);
  rows.push(['Måned', formatCurrency(Math.round(stats.omsMåned / (1 + config.mvaRate))), formatCurrency(stats.omsMåned), '', '']);
  rows.push(['År', formatCurrency(Math.round(stats.omsÅr / (1 + config.mvaRate))), formatCurrency(stats.omsÅr), '', '']);
  rows.push(['', '', '', '', '']);

  rows.push(['REISE', 'Eks mva', 'Inkl mva', '', '']);
  rows.push(['Måned', formatCurrency(Math.round(stats.reiseMåned / (1 + config.mvaRate))), formatCurrency(stats.reiseMåned), '', '']);
  rows.push(['År', formatCurrency(Math.round(stats.reiseÅr / (1 + config.mvaRate))), formatCurrency(stats.reiseÅr), '', '']);
  rows.push(['', '', '', '', '']);

  rows.push(['STATUS', 'Antall', '', '', '']);
  STATUS_LIST.forEach(s => rows.push([s, stats.byStatus[s] || 0, '', '', '']));
  rows.push(['', '', '', '', '']);

  rows.push(['OPPDRAGSTYPER', 'Antall', '', '', '']);
  rows.push(['Totalt', stats.total, '', '', '']);
  Object.keys(stats.byType).sort().forEach(t => rows.push([t, stats.byType[t], '', '', '']));
  rows.push(['', '', '', '', '']);

  rows.push(['BOLIGSTØRRELSE', 'Antall', '', '', '']);
  Object.keys(stats.byBoligBucket).sort().forEach(k => rows.push([k, stats.byBoligBucket[k], '', '', '']));
  rows.push(['', '', '', '', '']);

  rows.push([`INNTJENING SISTE 6 MND (Inkl mva)`, 'Oppdrag', 'Reise', 'Total', '']);
  stats.trendMonths.forEach(m => {
    rows.push([m.key, m.omsInkl, m.reiseInkl, m.totalInkl, '']);
  });

  try {
    await google.clearRange(config.sheet.dashboardName, 1, 200, 5);
  } catch { /* ignore if empty */ }
  await google.writeRange(config.sheet.dashboardName, 1, 1, rows);

  console.log('Dashboard updated');
}

module.exports = { updateDashboard, computeDashboardStats, STATUS_LIST };
