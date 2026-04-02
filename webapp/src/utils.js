const config = require('./config');

function formatCurrency(amount) {
  if (!amount || isNaN(amount)) return '0 kr';
  return Number(amount).toLocaleString('nb-NO') + ' kr';
}

function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
}

function formatDate(date, fmt) {
  const pad = (n) => String(n).padStart(2, '0');
  const d = date instanceof Date ? date : new Date(date);
  const map = {
    'dd': pad(d.getDate()),
    'MM': pad(d.getMonth() + 1),
    'yyyy': String(d.getFullYear()),
    'HH': pad(d.getHours()),
    'mm': pad(d.getMinutes()),
  };
  let result = fmt;
  for (const [key, val] of Object.entries(map)) {
    result = result.replace(key, val);
  }
  return result;
}

function parseDateString(dateStr) {
  if (dateStr instanceof Date) return dateStr;
  if (typeof dateStr !== 'string' || !dateStr) return null;
  const parts = dateStr.split(' ')[0].split('.');
  if (parts.length >= 3) {
    return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
  }
  return null;
}

function sanitizeParsedData(parsed) {
  if (!config.testMode) return parsed;

  const allowed = [
    config.email.ownerEmail.toLowerCase(),
    config.email.accountantEmail.toLowerCase(),
  ];

  if (parsed.selgerEpost && !allowed.includes(parsed.selgerEpost.toLowerCase())) {
    parsed.selgerEpost = config.email.accountantEmail;
  }
  if (parsed.meglerEpost && !allowed.includes(parsed.meglerEpost.toLowerCase())) {
    parsed.meglerEpost = config.email.accountantEmail;
  }
  return parsed;
}

function extractDriveFolderId(value) {
  const s = String(value || '').trim();
  if (!s) return '';
  const m1 = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m1) return m1[1];
  const m2 = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2) return m2[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(s) && !s.startsWith('http')) return s;
  return '';
}

function extractAddressFromHyperlink(adresse) {
  const m = String(adresse || '').match(/=HYPERLINK\("[^"]+","([^"]+)"\)/);
  if (m) return m[1].replace(/""/g, '"');
  return String(adresse || '');
}

function classifyBoligBucket(boligtype, areal) {
  if (!boligtype) return '';
  const a = Number(areal || 0);
  if (boligtype === 'Leilighet') return a && a <= 80 ? 'Leilighet (0-80 m²)' : 'Leilighet (80+ m²)';
  if (boligtype === 'Rekkehus/leilighet 2-4-mannsbolig') return a && a <= 80 ? 'Rekkehus/2-4 (0-80 m²)' : 'Rekkehus/2-4 (80+ m²)';
  if (boligtype === 'Enebolig/fritidsbolig') {
    if (a && a <= 150) return 'Enebolig/fritid (0-150 m²)';
    if (a && a <= 250) return 'Enebolig/fritid (150-250 m²)';
    return 'Enebolig/fritid (250+ m²)';
  }
  if (boligtype === 'Frittstående bygg') return 'Frittstående bygg';
  return boligtype;
}

function mapBoligtype(raw) {
  const bt = String(raw || '').toLowerCase();
  if (bt.includes('enebolig') || bt.includes('fritid')) return 'Enebolig/fritidsbolig';
  if (bt.includes('rekkehus') || bt.includes('mannsbolig')) return 'Rekkehus/leilighet 2-4-mannsbolig';
  if (bt.includes('leilighet')) return 'Leilighet';
  if (bt.includes('næring')) return 'Næringsbygg';
  if (bt.includes('frittstående') || bt.includes('garasje')) return 'Frittstående bygg';
  return 'Annet';
}

module.exports = {
  formatCurrency,
  getWeekNumber,
  formatDate,
  parseDateString,
  sanitizeParsedData,
  extractDriveFolderId,
  extractAddressFromHyperlink,
  classifyBoligBucket,
  mapBoligtype,
};
