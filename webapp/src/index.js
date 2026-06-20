const path = require('path');
const express = require('express');
const session = require('express-session');
const passport = require('passport');
const config = require('./config');
const { configurePassport } = require('./auth/google-oauth');
const requireAuth = require('./middleware/requireAuth');
const activity = require('./middleware/activityLog');
const cronMgr = require('./cron-manager');
const settingsModule = require('./settings');

const app = express();

// Trust Sliplane's TLS-terminating proxy so secure cookies work
app.set('trust proxy', 1);

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, '..', 'views'));

app.use(express.json());

app.use(session({
  secret: config.session.secret,
  resave: false,
  saveUninitialized: false,
  cookie: {
    httpOnly: true,
    sameSite: 'lax',
    secure: process.env.NODE_ENV === 'production',
    maxAge: config.session.maxAgeMs,
  },
}));

configurePassport();
app.use(passport.initialize());
app.use(passport.session());

// ============================================================
// PUBLIC ROUTES (no auth)
// ============================================================

app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    testMode: config.testMode,
    demoMode: config.demoMode,
    timestamp: new Date().toISOString(),
  });
});

// Public diagnostic — shows which OAuth config is loaded without exposing secrets.
app.get('/auth/debug', async (req, res) => {
  const out = {
    login_oauth: {
      callbackUrl: config.loginOAuth.callbackUrl,
      clientId_prefix: config.loginOAuth.clientId
        ? config.loginOAuth.clientId.slice(0, 25) + '…'
        : '(not set)',
      clientSecret_set: !!config.loginOAuth.clientSecret,
    },
    google_service_account: {
      clientId_prefix: config.google.clientId
        ? config.google.clientId.slice(0, 25) + '…'
        : '(not set)',
      clientSecret_set: !!config.google.clientSecret,
      refreshToken_set: !!config.google.refreshToken,
      refreshToken_length: config.google.refreshToken ? config.google.refreshToken.length : 0,
      access_token_test: null,
    },
    modes: {
      testMode: config.testMode,
      demoMode: config.demoMode,
      webappCronEnabled: config.webappCronEnabled,
      nodeEnv: process.env.NODE_ENV,
    },
  };

  // Try to actually exchange the refresh token for an access token
  if (config.google.refreshToken && config.google.clientId && config.google.clientSecret) {
    try {
      const { google } = require('googleapis');
      const auth = new google.auth.OAuth2(config.google.clientId, config.google.clientSecret);
      auth.setCredentials({ refresh_token: config.google.refreshToken });
      const t = await auth.getAccessToken();
      out.google_service_account.access_token_test = t && t.token ? 'OK — refresh-token virker' : 'FAILED — no token returned';
    } catch (e) {
      out.google_service_account.access_token_test = `FAILED — ${e.message}`;
    }
  } else {
    out.google_service_account.access_token_test = 'SKIPPED — missing creds';
  }

  res.json(out);
});

app.get('/login', (req, res) => {
  if (req.session && req.session.user) return res.redirect('/');
  res.render('login', { error: req.query.error || null });
});

app.get('/auth/google', passport.authenticate('google', {
  scope: ['openid', 'email', 'profile'],
  prompt: 'select_account',
}));

app.get('/auth/google/callback', (req, res, next) => {
  passport.authenticate('google', (err, user, info) => {
    if (err) {
      console.error('OAuth error:', err);
      return res.redirect('/login?error=failed');
    }
    if (!user) {
      const reason = info && info.message === 'denied' ? 'denied' : 'failed';
      return res.redirect(`/login?error=${reason}`);
    }
    req.session.user = user;
    activity.login(`Logget inn: ${user.email}`);
    return res.redirect('/');
  })(req, res, next);
});

app.post('/auth/logout', (req, res) => {
  const email = req.session.user?.email;
  if (email) activity.login(`Logget ut: ${email}`);
  req.session.destroy(() => res.redirect('/login'));
});

app.get('/auth/logout', (req, res) => {
  req.session.destroy(() => res.redirect('/login'));
});

// Dev-only: shortcut to create a session without going through Google.
// Active ONLY when NODE_ENV !== 'production' AND DEV_LOGIN_ENABLED=true.
if (process.env.NODE_ENV !== 'production' && process.env.DEV_LOGIN_ENABLED === 'true') {
  app.get('/auth/dev-login', (req, res) => {
    const email = (req.query.email || config.allowedEmails[0] || 'dev@example.com').toLowerCase();
    const { roleFor } = require('./auth/whitelist');
    req.session.user = { email, name: 'Dev User', picture: null, role: roleFor(email) };
    res.redirect('/');
  });
  console.log('DEV LOGIN ENABLED: /auth/dev-login is active. NEVER set this in production.');
}

// ============================================================
// AUTH GATE — everything below requires a session
// ============================================================

app.use(requireAuth);

// Accountant role is read-only: block any mutating /api request for them.
// They only need GET (oppdrag list + faktura-copy). Defense-in-depth on top of
// the frontend hiding edit/admin views.
app.use('/api', (req, res, next) => {
  if (req.session.user && req.session.user.role === 'accountant' && req.method !== 'GET') {
    return res.status(403).json({ success: false, error: 'Regnskap-rollen har kun lesetilgang' });
  }
  next();
});

// Static assets (public/) — gated: only authenticated users get the SPA shell + assets
app.use(express.static(path.join(__dirname, '..', 'public')));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, '..', 'public', 'index.html'));
});

// ============================================================
// API ROUTES (auth-protected)
// ============================================================

// Demo-mode short-circuit helper for write endpoints
function demoOk(res, message) {
  return res.json({ success: true, demoMode: true, message: `Demo: ${message}` });
}

// --- Session / current user ---
app.get('/api/me', (req, res) => {
  res.json({
    success: true,
    user: req.session.user,
    testMode: config.testMode,
    demoMode: config.demoMode,
    webappCronEnabled: config.webappCronEnabled,
  });
});

// --- Settings ---
app.get('/api/settings', async (req, res) => {
  try {
    const data = await settingsModule.get();
    // Mask secret values — never send actual secrets to the client.
    // The UI shows masked-input + placeholder "(satt)" if non-empty.
    const safe = { ...data };
    const masked = {};
    for (const key of Object.keys(safe)) {
      if (settingsModule.isSecretKey(key)) {
        masked[key] = !!safe[key];
        safe[key] = safe[key] ? '***' : '';
      }
    }
    res.json({
      success: true,
      settings: safe,
      schema: settingsModule.SCHEMA,
      defaults: settingsModule.DEFAULTS,
      secretsSet: masked,
    });
  } catch (e) {
    console.error('get-settings error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.patch('/api/settings', async (req, res) => {
  if (config.demoMode) {
    activity.admin('Demo: innstillinger oppdatert (ikke lagret)', { updates: req.body });
    return res.json({ success: true, demoMode: true, message: 'Demo: innstillinger simulert lagret' });
  }
  try {
    const updates = req.body || {};
    // Strip out empty values for secret keys — prevents accidental wipe by
    // submitting the masked '***' placeholder back as an empty/unchanged value.
    const cleaned = {};
    for (const [key, value] of Object.entries(updates)) {
      if (settingsModule.isSecretKey(key)) {
        if (!value || value === '***') continue; // skip — keep existing
      }
      cleaned[key] = value;
    }
    const result = await settingsModule.patch(cleaned);
    // Reapply cron schedules if any cron.* key changed
    if (Object.keys(cleaned).some(k => k.startsWith('cron.'))) {
      await cronMgr.applyScheduleFromSettings();
      activity.admin('Cron-jobber re-registrert etter endring av tidsplan', { keys: Object.keys(cleaned).filter(k => k.startsWith('cron.')) });
    }
    activity.admin('Innstillinger oppdatert', { updated: Object.keys(cleaned) });
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('patch-settings error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Activity log ---
app.get('/api/activity', (req, res) => {
  const limit = Math.min(parseInt(req.query.limit, 10) || 50, 200);
  res.json({ success: true, entries: activity.list(limit) });
});

// --- Fakturalogg ---
app.get('/api/fakturalogg', async (req, res) => {
  try {
    const limit = Math.min(parseInt(req.query.limit, 10) || 100, 500);
    const entries = await require('./fakturalogg').list(limit);
    res.json({ success: true, entries });
  } catch (e) {
    console.error('fakturalogg list error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/fakturalogg/:id/resend', async (req, res) => {
  if (config.demoMode) {
    activity.admin('Demo: faktura sendt på nytt (ikke sendt)', { id: req.params.id });
    return demoOk(res, 'faktura sendt på nytt (simulert)');
  }
  try {
    const result = await require('./fakturalogg').resend(req.params.id);
    activity.faktura(`Faktura sendt på nytt (oppdatering) — logg #${req.params.id}`, { by: req.session.user?.email });
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('fakturalogg resend error:', e);
    activity.error(`Faktura-resend feilet: ${e.message}`);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Prisliste ---
app.get('/api/prisliste', async (req, res) => {
  try {
    const { getRows, DEFAULTS } = require('./prisliste');
    const rows = await getRows();
    res.json({ success: true, rows, defaults: DEFAULTS });
  } catch (e) {
    console.error('get-prisliste error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.put('/api/prisliste', async (req, res) => {
  if (config.demoMode) {
    activity.admin('Demo: prisliste oppdatert (ikke lagret)', { count: (req.body?.rows || []).length });
    return demoOk(res, 'prisliste oppdatert (ikke lagret)');
  }
  try {
    const rows = req.body && req.body.rows;
    if (!Array.isArray(rows)) {
      return res.status(400).json({ success: false, error: 'body must be { rows: [...] }' });
    }
    const { setRows } = require('./prisliste');
    const result = await setRows(rows);
    activity.admin(`Prisliste oppdatert: ${result.updated} rader`, { by: req.session.user?.email });
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('put-prisliste error:', e);
    activity.error(`Prisliste-lagring feilet: ${e.message}`);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Cron job introspection ---
app.get('/api/admin/jobs', (req, res) => {
  res.json({ success: true, jobs: cronMgr.listJobs() });
});

// --- Manual trigger of a cron job ---
app.post('/api/admin/trigger/:job', async (req, res) => {
  if (config.demoMode) {
    activity.admin(`Demo: manuelt trigget "${req.params.job}"`);
    return res.json({ success: true, demoMode: true, message: `Demo: ${req.params.job} simulert` });
  }
  try {
    activity.admin(`Manuelt trigget "${req.params.job}"`, { by: req.session.user?.email });
    const result = await cronMgr.runJob(req.params.job);
    res.json({ success: true, ...result });
  } catch (e) {
    console.error(`manual-trigger error (${req.params.job}):`, e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Dashboard data (oppdrag list + KPIs) ---
app.get('/api/dashboard-data', async (req, res) => {
  try {
    const { getDashboardData } = require('./data');
    const data = await getDashboardData();
    res.json(data);
  } catch (e) {
    console.error('dashboard-data error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Dashboard stats (full breakdowns + trend, no oppdrag list) ---
app.get('/api/dashboard-stats', async (req, res) => {
  try {
    const { getDashboardStats } = require('./data');
    const stats = await getDashboardStats();
    res.json({ success: true, stats });
  } catch (e) {
    console.error('dashboard-stats error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Single oppdrag ---
app.get('/api/oppdrag/:row', async (req, res) => {
  try {
    const rowNum = parseInt(req.params.row, 10);
    if (!Number.isInteger(rowNum) || rowNum < 2) {
      return res.status(400).json({ success: false, error: 'invalid row' });
    }
    const { getOppdrag } = require('./data');
    const oppdrag = await getOppdrag(rowNum);
    if (!oppdrag) return res.status(404).json({ success: false, error: 'not found' });
    res.json({ success: true, oppdrag });
  } catch (e) {
    console.error('get-oppdrag error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Faktura copy text (PowerOffice paste, mirrors the email layout) ---
app.get('/api/faktura-copy/:row', async (req, res) => {
  try {
    const rowNum = parseInt(req.params.row, 10);
    if (!Number.isInteger(rowNum) || rowNum < 2) {
      return res.status(400).json({ success: false, error: 'invalid row' });
    }
    if (config.demoMode) {
      return res.json({ success: true, text: 'DEMO\nKunde: Demo Kunde\nProdukt 5 — Tilstandsrapport: 8000 kr eks. mva\nTOTAL EKS. MVA: 8000 kr' });
    }
    const google = require('./google');
    const data = await google.getSheetData(config.sheet.name);
    const rowData = data[rowNum - 1];
    if (!rowData) return res.status(404).json({ success: false, error: 'not found' });
    const { buildFakturaText } = require('./emails');
    const text = buildFakturaText(rowData);
    res.json({ success: true, text });
  } catch (e) {
    console.error('faktura-copy error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Patch oppdrag (whitelisted fields) ---
app.patch('/api/oppdrag/:row', async (req, res) => {
  if (config.demoMode) return demoOk(res, 'felt oppdatert (ikke lagret)');
  try {
    const rowNum = parseInt(req.params.row, 10);
    if (!Number.isInteger(rowNum) || rowNum < 2) {
      return res.status(400).json({ success: false, error: 'invalid row' });
    }
    if (!req.body || typeof req.body !== 'object') {
      return res.status(400).json({ success: false, error: 'body must be JSON object' });
    }
    const { patchOppdrag, getOppdrag } = require('./data');
    const result = await patchOppdrag(rowNum, req.body);
    const oppdrag = await getOppdrag(rowNum);
    activity.patch(`Rad ${rowNum} oppdatert (${Object.keys(req.body).join(', ')})`, { by: req.session.user?.email, rowNum, fields: Object.keys(req.body) });
    res.json({ success: true, ...result, oppdrag });
  } catch (e) {
    console.error('patch-oppdrag error:', e);
    activity.error(`PATCH rad ${req.params.row} feilet: ${e.message}`);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Change status (routes through handleStatusChange for side effects) ---
app.post('/api/oppdrag/:row/status', async (req, res) => {
  if (config.demoMode) return demoOk(res, 'status endret (ikke lagret)');
  try {
    const rowNum = parseInt(req.params.row, 10);
    const { status } = req.body || {};
    if (!Number.isInteger(rowNum) || rowNum < 2 || !status) {
      return res.status(400).json({ success: false, error: 'row and status required' });
    }
    const google = require('./google');
    const { COL } = require('./columns');
    await google.updateCell(config.sheet.name, rowNum, COL.STATUS, status);
    const { handleStatusChange } = require('./oppdrag');
    await handleStatusChange(rowNum, status);
    const { bustCache, getOppdrag } = require('./data');
    bustCache();
    const oppdrag = await getOppdrag(rowNum);
    activity.status(`Rad ${rowNum}: status → "${status}"`, { by: req.session.user?.email, rowNum, status });
    res.json({ success: true, oppdrag });
  } catch (e) {
    console.error('oppdrag-status error:', e);
    activity.error(`Status-endring rad ${req.params.row} feilet: ${e.message}`);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Email scanning ---
app.post('/api/scan-emails', async (req, res) => {
  if (config.demoMode) return demoOk(res, 'simulert e-postskanning fullført');
  try {
    const { scanIncomingEmails } = require('./scanner');
    const result = await scanIncomingEmails();
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('scan-emails error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- IVIT scraping ---
app.post('/api/process-ivit', async (req, res) => {
  if (config.demoMode) return demoOk(res, 'simulert IVIT-scraping fullført');
  try {
    const { processIVITScraping } = require('./ivit');
    const result = await processIVITScraping();
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('process-ivit error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Dashboard sheet write ---
app.post('/api/update-dashboard', async (req, res) => {
  if (config.demoMode) return demoOk(res, 'dashboard oppdatert (simulert)');
  try {
    const { updateDashboard } = require('./dashboard');
    await updateDashboard();
    res.json({ success: true });
  } catch (e) {
    console.error('update-dashboard error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Reminders ---
app.post('/api/check-reminders', async (req, res) => {
  if (config.demoMode) return demoOk(res, 'påminnelser sjekket (simulert)');
  try {
    const { checkReminders } = require('./reminders');
    const result = await checkReminders();
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('check-reminders error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Weekly report ---
app.post('/api/send-weekly-report', async (req, res) => {
  if (config.demoMode) return demoOk(res, 'ukerapport simulert (ingen e-post sendt)');
  try {
    const { sendWeeklyReport } = require('./reminders');
    await sendWeeklyReport();
    res.json({ success: true });
  } catch (e) {
    console.error('send-weekly-report error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Send faktura ---
app.post('/api/send-faktura', async (req, res) => {
  if (config.demoMode) return demoOk(res, 'faktura-batch simulert (ingen e-post sendt)');
  try {
    const { sendFakturaTilRegnskap } = require('./oppdrag');
    const count = await sendFakturaTilRegnskap();
    const { bustCache } = require('./data');
    bustCache();
    activity.faktura(`Faktura-batch sendt: ${count} oppdrag`, { by: req.session.user?.email, count });
    res.json({ success: true, count });
  } catch (e) {
    console.error('send-faktura error:', e);
    activity.error(`Faktura-batch feilet: ${e.message}`);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Manual registration ---
app.post('/api/register-oppdrag', async (req, res) => {
  if (config.demoMode) return demoOk(res, 'oppdrag registrert (ikke lagret)');
  try {
    const { adresse, selger, telefon, epost, boligtype, rapporttype, areal, megler, merknad } = req.body;
    if (!adresse) {
      return res.status(400).json({ success: false, error: 'Adresse er påkrevd' });
    }
    const { registerManualOppdrag } = require('./oppdrag');
    const oppdragsnr = await registerManualOppdrag({
      adresse, selger, telefon, epost, boligtype, rapporttype, areal, megler, merknad,
    });
    const { bustCache } = require('./data');
    bustCache();
    activity.register(`Nytt manuelt oppdrag: ${oppdragsnr} — ${adresse}`, { by: req.session.user?.email, oppdragsnr });
    res.json({ success: true, oppdragsnr });
  } catch (e) {
    console.error('register-oppdrag error:', e);
    activity.error(`Manuell registrering feilet: ${e.message}`);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Status change ---
app.post('/api/status-change', async (req, res) => {
  if (config.demoMode) return demoOk(res, 'status endret (ikke lagret)');
  try {
    const { row, status } = req.body;
    if (!row || !status) {
      return res.status(400).json({ success: false, error: 'row and status are required' });
    }
    const { handleStatusChange } = require('./oppdrag');
    await handleStatusChange(row, status);
    const { bustCache } = require('./data');
    bustCache();
    res.json({ success: true });
  } catch (e) {
    console.error('status-change error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Recalculate price ---
app.post('/api/recalculate-price', async (req, res) => {
  if (config.demoMode) return demoOk(res, 'pris rekalkulert (ikke lagret)');
  try {
    const { row } = req.body;
    if (!row) {
      return res.status(400).json({ success: false, error: 'row is required' });
    }
    const { calculatePriceForRow } = require('./oppdrag');
    await calculatePriceForRow(row);
    const { bustCache } = require('./data');
    bustCache();
    res.json({ success: true });
  } catch (e) {
    console.error('recalculate-price error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Recalculate travel ---
app.post('/api/recalculate-travel', async (req, res) => {
  if (config.demoMode) return demoOk(res, 'reisekostnad rekalkulert (ikke lagret)');
  try {
    const { row, address } = req.body;
    if (!row || !address) {
      return res.status(400).json({ success: false, error: 'row and address are required' });
    }
    const { recalculateTravel } = require('./oppdrag');
    await recalculateTravel(row, address);
    const { bustCache } = require('./data');
    bustCache();
    res.json({ success: true });
  } catch (e) {
    console.error('recalculate-travel error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ============================================================
// CRON JOBS
// ============================================================
// Schedules come from settings (Innstillinger-sheet, hot-reloadable via /api/settings PATCH).
// Gated by WEBAPP_CRON_ENABLED + !DEMO_MODE inside cron-manager.

cronMgr.registerDefaults();
// Apply initial schedule asynchronously (don't block server boot)
cronMgr.applyScheduleFromSettings().catch(e => {
  console.error('Initial cron registration failed:', e.message);
});

// ============================================================
// START SERVER
// ============================================================

const server = app.listen(config.port, () => {
  console.log(`Naava Takst webapp listening on port ${config.port}`);
  console.log(`Test mode: ${config.testMode}`);
  console.log(`Demo mode: ${config.demoMode}`);
  console.log(`Allowed emails: ${config.allowedEmails.length} configured`);
  if (config.webappCronEnabled && !config.demoMode) {
    console.log('Cron jobs: scan(5m), reminders(1h), dashboard(1h), ivit(15m), weekly(Fri 16:00)');
  } else if (config.webappCronEnabled && config.demoMode) {
    console.log('Cron jobs: SUPPRESSED (DEMO_MODE=true overrides WEBAPP_CRON_ENABLED).');
  } else {
    console.log('Cron jobs: DISABLED (WEBAPP_CRON_ENABLED=false). Apps Script handles automation.');
  }
});

async function shutdown(signal) {
  console.log(`\n${signal} received. Shutting down...`);
  server.close(() => {
    console.log('Server closed.');
    process.exit(0);
  });
  setTimeout(() => { process.exit(1); }, 10000);
}

process.on('SIGTERM', () => shutdown('SIGTERM'));
process.on('SIGINT', () => shutdown('SIGINT'));
