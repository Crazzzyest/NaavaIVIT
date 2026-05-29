/**
 * verify-credentials.js — Ping Google APIs to confirm refresh token works.
 * Does NOT modify the sheet. Read-only. Independent of DEMO_MODE.
 *
 * Run: node scripts/verify-credentials.js
 */

const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '..', '.env') });

const { google } = require('googleapis');

const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REFRESH_TOKEN = process.env.GOOGLE_REFRESH_TOKEN;
const SHEET_ID = '1VEzaCNEkvbWYZf0UOrj6IG5PUQB3hL1enB24PvsQ1MI'; // from config.js

if (!CLIENT_ID || !CLIENT_SECRET || !REFRESH_TOKEN) {
  console.error('❌ Mangler en eller flere av GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET / GOOGLE_REFRESH_TOKEN i .env');
  process.exit(1);
}

const auth = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET);
auth.setCredentials({ refresh_token: REFRESH_TOKEN });

async function main() {
  console.log('▶  Prøver å hente access-token fra refresh-token...');
  try {
    const t = await auth.getAccessToken();
    if (!t || !t.token) throw new Error('Ingen access-token returnert');
    console.log(`✓  Access-token OK (gyldig ${Math.round((t.res?.data?.expires_in || 3600) / 60)} min)\n`);
  } catch (e) {
    console.error(`❌ Klarte ikke å hente access-token: ${e.message}`);
    if (e.message.includes('invalid_grant')) {
      console.error('   → Refresh-tokenet er revokert eller utløpt. Kjør auth-setup.js på nytt.');
    }
    process.exit(1);
  }

  // --- Sheets: les overskrifts-rad ---
  console.log(`▶  Leser fra Sheet (ID: ${SHEET_ID})...`);
  try {
    const sheets = google.sheets({ version: 'v4', auth });
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const tabs = meta.data.sheets.map(s => s.properties.title);
    console.log(`✓  Sheet-tilgang OK. Faner: ${tabs.join(', ')}`);

    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Oppdragslogg!A1:E3',
    });
    const rows = res.data.values || [];
    console.log(`✓  Leste ${rows.length} rader fra Oppdragslogg!A1:E3:`);
    rows.forEach((r, i) => console.log(`   [${i}] ${r.join(' | ')}`));
    console.log('');
  } catch (e) {
    console.error(`❌ Sheets-feil: ${e.message}`);
    if (e.message.includes('not been used') || e.message.includes('disabled')) {
      console.error('   → Sjekk at Google Sheets API er aktivert i Google Cloud Console.');
    }
    process.exit(1);
  }

  // --- Drive: list rotmappen (bare for å sjekke at scope funker) ---
  console.log('▶  Sjekker Drive-tilgang...');
  try {
    const drive = google.drive({ version: 'v3', auth });
    const res = await drive.files.list({ pageSize: 1, fields: 'files(id, name)' });
    console.log(`✓  Drive OK (kan liste filer)\n`);
  } catch (e) {
    console.error(`❌ Drive-feil: ${e.message}`);
    process.exit(1);
  }

  // --- Gmail: hent profilen (bare for å sjekke at scope funker) ---
  console.log('▶  Sjekker Gmail-tilgang...');
  try {
    const gmail = google.gmail({ version: 'v1', auth });
    const res = await gmail.users.getProfile({ userId: 'me' });
    console.log(`✓  Gmail OK — innlogget som: ${res.data.emailAddress}`);
    console.log(`   (Antall meldinger i mailboxen: ${res.data.messagesTotal})\n`);
  } catch (e) {
    console.error(`❌ Gmail-feil: ${e.message}`);
    process.exit(1);
  }

  console.log('🎉 Alt funker! Refresh-tokenet har tilgang til Sheets + Drive + Gmail.');
  console.log('');
  console.log('Neste steg når du er klar:');
  console.log('  1. Sett DEMO_MODE=false i .env (men hold WEBAPP_CRON_ENABLED=false)');
  console.log('  2. Restart webappen (Ctrl+C, npm start)');
  console.log('  3. Webappen leser nå fra ekte sheet i stedet for demo-data');
}

main().catch(e => {
  console.error('❌ Uventet feil:', e.message);
  process.exit(1);
});
