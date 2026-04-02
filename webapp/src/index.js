const express = require('express');
const cron = require('node-cron');
const config = require('./config');

const app = express();
app.use(express.json());

// ============================================================
// ROUTES
// ============================================================

app.get('/health', (req, res) => {
  res.json({ status: 'ok', testMode: config.testMode, timestamp: new Date().toISOString() });
});

// --- Email scanning ---
app.post('/api/scan-emails', async (req, res) => {
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
  try {
    const { processIVITScraping } = require('./ivit');
    const result = await processIVITScraping();
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('process-ivit error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Dashboard ---
app.post('/api/update-dashboard', async (req, res) => {
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
  try {
    const { sendFakturaTilRegnskap } = require('./oppdrag');
    const count = await sendFakturaTilRegnskap();
    res.json({ success: true, count });
  } catch (e) {
    console.error('send-faktura error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Manual registration ---
app.post('/api/register-oppdrag', async (req, res) => {
  try {
    const { adresse, selger, telefon, epost, boligtype, rapporttype, areal, megler, merknad } = req.body;
    if (!adresse) {
      return res.status(400).json({ success: false, error: 'Adresse er påkrevd' });
    }
    const { registerManualOppdrag } = require('./oppdrag');
    const oppdragsnr = await registerManualOppdrag({
      adresse, selger, telefon, epost, boligtype, rapporttype, areal, megler, merknad,
    });
    res.json({ success: true, oppdragsnr });
  } catch (e) {
    console.error('register-oppdrag error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Status change ---
app.post('/api/status-change', async (req, res) => {
  try {
    const { row, status } = req.body;
    if (!row || !status) {
      return res.status(400).json({ success: false, error: 'row and status are required' });
    }
    const { handleStatusChange } = require('./oppdrag');
    await handleStatusChange(row, status);
    res.json({ success: true });
  } catch (e) {
    console.error('status-change error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Recalculate price ---
app.post('/api/recalculate-price', async (req, res) => {
  try {
    const { row } = req.body;
    if (!row) {
      return res.status(400).json({ success: false, error: 'row is required' });
    }
    const { calculatePriceForRow } = require('./oppdrag');
    await calculatePriceForRow(row);
    res.json({ success: true });
  } catch (e) {
    console.error('recalculate-price error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// --- Recalculate travel ---
app.post('/api/recalculate-travel', async (req, res) => {
  try {
    const { row, address } = req.body;
    if (!row || !address) {
      return res.status(400).json({ success: false, error: 'row and address are required' });
    }
    const { recalculateTravel } = require('./oppdrag');
    await recalculateTravel(row, address);
    res.json({ success: true });
  } catch (e) {
    console.error('recalculate-travel error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ============================================================
// CRON JOBS
// ============================================================

// Scan emails every 5 minutes
cron.schedule('*/5 * * * *', async () => {
  try {
    const { scanIncomingEmails } = require('./scanner');
    await scanIncomingEmails();
  } catch (e) {
    console.error('Cron scan-emails error:', e.message);
  }
});

// Check reminders every hour
cron.schedule('0 * * * *', async () => {
  try {
    const { checkReminders } = require('./reminders');
    await checkReminders();
  } catch (e) {
    console.error('Cron check-reminders error:', e.message);
  }
});

// Update dashboard every hour
cron.schedule('30 * * * *', async () => {
  try {
    const { updateDashboard } = require('./dashboard');
    await updateDashboard();
  } catch (e) {
    console.error('Cron update-dashboard error:', e.message);
  }
});

// Process IVIT scraping every 15 minutes
cron.schedule('*/15 * * * *', async () => {
  try {
    const { processIVITScraping } = require('./ivit');
    await processIVITScraping();
  } catch (e) {
    console.error('Cron process-ivit error:', e.message);
  }
});

// Weekly report on Fridays at 16:00
cron.schedule('0 16 * * 5', async () => {
  try {
    const { sendWeeklyReport } = require('./reminders');
    await sendWeeklyReport();
  } catch (e) {
    console.error('Cron weekly-report error:', e.message);
  }
});

// ============================================================
// START SERVER
// ============================================================

const server = app.listen(config.port, () => {
  console.log(`Naava Takst webapp listening on port ${config.port}`);
  console.log(`Test mode: ${config.testMode}`);
  console.log('Cron jobs: scan(5m), reminders(1h), dashboard(1h), ivit(15m), weekly(Fri 16:00)');
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
