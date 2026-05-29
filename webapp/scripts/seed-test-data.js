/**
 * seed-test-data.js — Bulk-create test oppdrag in the live Google Sheet
 *
 * What it does:
 *   1. Logs in via /auth/dev-login (requires DEV_LOGIN_ENABLED=true on the running server)
 *   2. For each entry in src/demo-data.js OPPDRAG:
 *      a) POST /api/register-oppdrag with the manual-intake-compatible fields
 *      b) Look up the newly-appended row number
 *      c) PATCH /api/oppdrag/:row with the remaining fields the manual form doesn't expose
 *
 * Use this to seed a test sheet (or to repopulate prod with known test data — careful!).
 *
 * Prerequisites:
 *   - DEMO_MODE=false in .env (otherwise the POST/PATCH calls return "Demo: simulert" and
 *     nothing actually persists)
 *   - GOOGLE_REFRESH_TOKEN populated in .env
 *   - The Google Sheet pointed to by config.sheet.id is the one you want to seed (consider
 *     pointing at a copy before running this against prod)
 *   - The webapp is running locally on PORT (default 3002): `npm start` in another terminal
 *
 * Usage:
 *   node scripts/seed-test-data.js
 *   node scripts/seed-test-data.js --base http://localhost:3002
 *   node scripts/seed-test-data.js --skip-existing       (skips entries whose oppdragsnr already in sheet)
 */

const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '..', '.env') });

const { OPPDRAG } = require('../src/demo-data');

const args = process.argv.slice(2);
function arg(name, def) {
  const idx = args.indexOf(`--${name}`);
  if (idx === -1) return def;
  const next = args[idx + 1];
  return next && !next.startsWith('--') ? next : true;
}

const BASE = arg('base', `http://localhost:${process.env.PORT || 3002}`);
const SKIP_EXISTING = !!arg('skip-existing', false);

// ============================================================
// Helpers
// ============================================================

let cookie = '';

async function devLogin() {
  // Get session cookie via dev backdoor
  const res = await fetch(`${BASE}/auth/dev-login?email=jacob@naava.no`, { redirect: 'manual' });
  const setCookie = res.headers.get('set-cookie');
  if (!setCookie) throw new Error('No Set-Cookie header from /auth/dev-login — is DEV_LOGIN_ENABLED=true on the server?');
  cookie = setCookie.split(';')[0]; // "connect.sid=s%3A..."
  console.log(`[seed] Logged in (cookie: ${cookie.slice(0, 40)}...)`);
}

async function api(method, p, body) {
  const res = await fetch(`${BASE}${p}`, {
    method,
    headers: {
      'Content-Type': 'application/json',
      'Cookie': cookie,
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  let parsed;
  try { parsed = JSON.parse(text); } catch { parsed = { raw: text }; }
  if (!res.ok) {
    throw new Error(`${method} ${p} → ${res.status}: ${parsed.error || text.slice(0, 200)}`);
  }
  return parsed;
}

async function getCurrentOppdragsnrs() {
  const data = await api('GET', '/api/dashboard-data');
  return new Set((data.oppdrag || []).map(o => o.oppdragsnr));
}

// ============================================================
// Field maps
// ============================================================

// Fields handled by POST /api/register-oppdrag (manual intake form)
function intakeBodyFrom(o) {
  // Map rapporttype → form's rapporttype dropdown values
  // (Backend mapTypes in oppdrag.js then re-maps to oppdragstype + rapporttype)
  let rapporttype = '';
  if (o.rapporttype === 'Tilstandsrapport m/teknisk og markedsverdi' || o.medMarkedsverdi) {
    rapporttype = 'Tilstandsrapport m/teknisk og markedsverdi';
  } else if (o.rapporttype === 'Tilstandsrapport' || o.oppdragstype === 'Tilstandsrapport') {
    rapporttype = 'Tilstandsrapport';
  } else if (o.rapporttype === 'Skadetakstrapport' || o.oppdragstype === 'Skadetakst') {
    rapporttype = 'Skadetakstrapport';
  } else if (o.rapporttype === 'Reklamasjonsrapport' || o.oppdragstype === 'Reklamasjon') {
    rapporttype = 'Reklamasjonsrapport';
  } else if (o.rapporttype === 'Vurderingsrapport' || o.oppdragstype === 'Vurderingsoppdrag') {
    rapporttype = 'Vurderingsrapport';
  } else if (o.rapporttype === 'Overtagelsesrapport' || o.oppdragstype === 'Bistand overtagelse') {
    rapporttype = 'Overtagelsesrapport';
  } else {
    rapporttype = 'Annen rapport';
  }

  return {
    adresse: o.adresse,
    selger: o.selger,
    telefon: o.selgerTlf,
    epost: o.selgerEpost,
    megler: o.megler,
    boligtype: o.boligtype,
    rapporttype,
    areal: o.areal === '' || o.areal == null ? '' : String(o.areal),
    merknad: o.notater,
  };
}

// Fields not covered by intake — set via PATCH /api/oppdrag/:row
function patchBodyFrom(o) {
  const body = {
    fakturaRef: o.fakturaRef || '',
    meglerEpost: o.meglerEpost || '',
    fakturaSendesTil: o.fakturaSendesTil || '',
    fakturamotaker: o.fakturamotaker || '',
    antallTilleggsbygg: o.antallTilleggsbygg === '' ? '' : Number(o.antallTilleggsbygg || 0),
    medMarkedsverdi: !!o.medMarkedsverdi,
    befaringDato: o.befaringDato || '',
    befaringKl: o.befaringKl || '',
    sumFergeBom: o.sumFergeBom === '' ? '' : Number(o.sumFergeBom || 0),
    antallDeleReise: o.antallDeleReise === '' ? '' : Number(o.antallDeleReise || 0),
    kommentarRegnskap: o.kommentarRegnskap || '',
    scanIvit: !!o.scanIvit,
    kanFaktureres: !!o.kanFaktureres,
  };
  // Strip empty strings — PATCH treats undefined as "no change"
  for (const k of Object.keys(body)) {
    if (body[k] === '') delete body[k];
  }
  return body;
}

// ============================================================
// Main
// ============================================================

async function main() {
  console.log(`[seed] Target: ${BASE}`);
  console.log(`[seed] Entries: ${OPPDRAG.length}`);
  console.log(`[seed] Skip existing: ${SKIP_EXISTING}`);
  console.log('');

  await devLogin();

  // Check if /api/me reports demo mode (in which case POSTs are no-ops)
  const me = await api('GET', '/api/me');
  if (me.demoMode) {
    console.error('[seed] ⚠️  Server is in DEMO_MODE — POST/PATCH return simulert.');
    console.error('[seed]    Set DEMO_MODE=false in .env, restart the server, and retry.');
    process.exit(1);
  }
  console.log(`[seed] Server mode: testMode=${me.testMode}, demoMode=${me.demoMode}`);
  console.log('');

  const existing = SKIP_EXISTING ? await getCurrentOppdragsnrs() : new Set();
  let created = 0, skipped = 0, errors = 0;

  for (let i = 0; i < OPPDRAG.length; i++) {
    const o = OPPDRAG[i];
    const label = `[${i + 1}/${OPPDRAG.length}] ${o.adresse || '(no address)'}`;

    if (SKIP_EXISTING && existing.has(o.oppdragsnr)) {
      console.log(`${label} — SKIPPED (oppdragsnr ${o.oppdragsnr} already exists)`);
      skipped++;
      continue;
    }

    try {
      // 1. Manual intake (creates row + drive folder + price + travel calc)
      const intake = intakeBodyFrom(o);
      const intakeRes = await api('POST', '/api/register-oppdrag', intake);
      const newOppdragsnr = intakeRes.oppdragsnr;
      console.log(`${label}\n   → created: ${newOppdragsnr}`);

      // 2. Find row by oppdragsnr (latest in dashboard-data)
      const dash = await api('GET', '/api/dashboard-data');
      const row = (dash.oppdrag || []).find(x => x.oppdragsnr === newOppdragsnr);
      if (!row) {
        console.warn(`   ⚠️  Could not find row for ${newOppdragsnr} — skipping PATCH`);
        created++;
        continue;
      }

      // 3. PATCH remaining fields
      const patch = patchBodyFrom(o);
      if (Object.keys(patch).length > 0) {
        await api('PATCH', `/api/oppdrag/${row.rowNum}`, patch);
        console.log(`   → patched ${Object.keys(patch).length} additional fields`);
      }

      // 4. Status change if not Mottatt
      if (o.status && o.status !== 'Mottatt') {
        await api('POST', `/api/oppdrag/${row.rowNum}/status`, { status: o.status });
        console.log(`   → status: ${o.status}`);
      }

      created++;
    } catch (e) {
      console.error(`${label}\n   ❌ ${e.message}`);
      errors++;
    }

    // Be polite to Sheets API quota
    await new Promise(r => setTimeout(r, 500));
  }

  console.log('');
  console.log(`[seed] Done. Created: ${created}, Skipped: ${skipped}, Errors: ${errors}`);
}

main().catch(e => {
  console.error('[seed] Fatal:', e.message);
  process.exit(1);
});
