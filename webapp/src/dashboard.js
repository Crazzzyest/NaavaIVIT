const config = require('./config');
const { COL } = require('./columns');
const google = require('./google');
const { formatCurrency, parseDateString, classifyBoligBucket, formatDate } = require('./utils');

async function updateDashboard() {
  console.log('Updating dashboard...');

  const data = await google.getSheetData(config.sheet.name);
  if (!data || data.length < 2) {
    await google.writeRange(config.sheet.dashboardName, 1, 1, [['Ingen oppdrag ennå.']]);
    return;
  }

  const now = new Date();
  const curMonth = now.getMonth();
  const curYear = now.getFullYear();

  const statusList = ['Mottatt', 'Avtalt befaring', 'Befart', 'Utkast', 'Endelig rapport', 'Kan faktureres', 'Fakturert', 'Oppdrag kansellert', 'Oppdrag fullført'];
  const stats = {
    total: data.length - 1,
    byStatus: {},
    byType: {},
    byBoligBucket: {},
    omsMåned: 0, omsÅr: 0,
    reiseMåned: 0, reiseÅr: 0,
    uteståendeInkl: 0, uteståendeEks: 0,
    snittPrisInkl: 0, snittPrisEks: 0, snittAntall: 0,
    trendMonths: {},
  };

  statusList.forEach(s => { stats.byStatus[s] = 0; });

  const monthsToShow = 6;
  const monthKeys = [];
  for (let k = monthsToShow - 1; k >= 0; k--) {
    const d = new Date(now.getFullYear(), now.getMonth() - k, 1);
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    monthKeys.push(key);
    stats.trendMonths[key] = { omsInkl: 0, reiseInkl: 0, totalInkl: 0 };
  }

  let sumPrisInkl = 0, sumPrisEks = 0, countPris = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[COL.STATUS - 1];
    const type = row[COL.OPPDRAGSTYPE - 1];
    const bolig = row[COL.BOLIGTYPE - 1];
    const areal = row[COL.AREAL - 1] ? Number(row[COL.AREAL - 1]) : 0;
    const prisInkl = Number(row[COL.PRIS_INKL - 1] || 0);
    const prisEks = Number(row[COL.PRIS_EKS - 1] || 0);
    const reiseInkl = Number(row[COL.REISE_INKL - 1] || 0);

    if (stats.byStatus[status] !== undefined) stats.byStatus[status]++;
    if (type) stats.byType[type] = (stats.byType[type] || 0) + 1;

    const bucket = classifyBoligBucket(bolig, areal);
    if (bucket) stats.byBoligBucket[bucket] = (stats.byBoligBucket[bucket] || 0) + 1;

    if (status === 'Kan faktureres') {
      stats.uteståendeInkl += (prisInkl + reiseInkl);
      stats.uteståendeEks += Math.round((prisInkl + reiseInkl) / (1 + config.mvaRate));
    }

    if (status === 'Fakturert' || status === 'Kan faktureres') {
      stats.omsÅr += prisInkl;
      stats.reiseÅr += reiseInkl;

      const statusDato = parseDateString(row[COL.DATO_STATUSENDRING - 1]);
      if (statusDato) {
        if (statusDato.getMonth() === curMonth && statusDato.getFullYear() === curYear) {
          stats.omsMåned += prisInkl;
          stats.reiseMåned += reiseInkl;
        }
        const mKey = `${statusDato.getFullYear()}-${String(statusDato.getMonth() + 1).padStart(2, '0')}`;
        if (stats.trendMonths[mKey]) {
          stats.trendMonths[mKey].omsInkl += prisInkl;
          stats.trendMonths[mKey].reiseInkl += reiseInkl;
          stats.trendMonths[mKey].totalInkl += (prisInkl + reiseInkl);
        }
      }

      if (prisInkl > 0) { sumPrisInkl += prisInkl; sumPrisEks += prisEks; countPris++; }
    }
  }

  stats.snittAntall = countPris;
  stats.snittPrisInkl = countPris ? Math.round(sumPrisInkl / countPris) : 0;
  stats.snittPrisEks = countPris ? Math.round(sumPrisEks / countPris) : 0;

  // Build dashboard rows
  const rows = [];
  rows.push(['NAAVA TAKST DASHBOARD' + (config.testMode ? ' TEST' : ''), '', '', '', '']);
  rows.push([formatDate(now, 'dd.MM.yyyy HH:mm'), '', '', '', '']);
  rows.push(['', '', '', '', '']);

  rows.push(['OVERSIKT', '', '', '', '']);
  rows.push(['Totalt antall oppdrag', stats.total, '', '', '']);
  rows.push(['Aktive (ikke Fakturert)', stats.total - (stats.byStatus['Fakturert'] || 0) - (stats.byStatus['Oppdrag kansellert'] || 0) - (stats.byStatus['Oppdrag fullført'] || 0), '', '', '']);
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
  statusList.forEach(s => rows.push([s, stats.byStatus[s] || 0, '', '', '']));
  rows.push(['', '', '', '', '']);

  rows.push(['OPPDRAGSTYPER', 'Antall', '', '', '']);
  rows.push(['Totalt', stats.total, '', '', '']);
  Object.keys(stats.byType).sort().forEach(t => rows.push([t, stats.byType[t], '', '', '']));
  rows.push(['', '', '', '', '']);

  rows.push(['BOLIGSTØRRELSE', 'Antall', '', '', '']);
  Object.keys(stats.byBoligBucket).sort().forEach(k => rows.push([k, stats.byBoligBucket[k], '', '', '']));
  rows.push(['', '', '', '', '']);

  rows.push([`INNTJENING SISTE ${monthsToShow} MND (Inkl mva)`, 'Oppdrag', 'Reise', 'Total', '']);
  monthKeys.forEach(k => {
    const v = stats.trendMonths[k];
    rows.push([k, v.omsInkl, v.reiseInkl, v.totalInkl, '']);
  });

  // Clear and write
  try {
    await google.clearRange(config.sheet.dashboardName, 1, 200, 5);
  } catch { /* ignore if empty */ }
  await google.writeRange(config.sheet.dashboardName, 1, 1, rows);

  console.log('Dashboard updated');
}

module.exports = { updateDashboard };
