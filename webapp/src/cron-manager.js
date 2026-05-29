const cron = require('node-cron');
const config = require('./config');
const settings = require('./settings');
const activity = require('./middleware/activityLog');

// ============================================================
// Job registry
// ============================================================
// Each job has: name, settingsKey (e.g. 'cron.scanEmails'), handler (async fn)
// We track the active cron task per job so we can stop + recreate on settings change.

const jobs = new Map(); // name → { settingsKey, handler, task }

function defineJob(name, settingsKey, handler) {
  jobs.set(name, { name, settingsKey, handler, task: null });
}

async function runJob(name) {
  const job = jobs.get(name);
  if (!job) throw new Error(`Unknown job: ${name}`);
  activity.cron(`Starter ${name}`);
  try {
    const result = await job.handler();
    activity.cron(`Fullført ${name}`, result || {});
    return { success: true, result };
  } catch (e) {
    activity.cronErr(`Feil i ${name}: ${e.message}`);
    throw e;
  }
}

async function applyScheduleFromSettings() {
  // Stop all existing tasks
  for (const job of jobs.values()) {
    if (job.task) {
      job.task.stop();
      job.task = null;
    }
  }

  // If cron is disabled (or demo mode), don't register anything
  if (!config.webappCronEnabled || config.demoMode) return;

  const s = await settings.get();
  for (const job of jobs.values()) {
    const expression = s[job.settingsKey];
    if (!expression || typeof expression !== 'string') continue;
    if (!cron.validate(expression)) {
      activity.cronErr(`Ugyldig cron-uttrykk for ${job.name}: "${expression}"`);
      continue;
    }
    try {
      const task = cron.schedule(expression, async () => {
        try {
          await runJob(job.name);
        } catch (_) { /* already logged in runJob */ }
      });
      job.task = task;
    } catch (e) {
      activity.cronErr(`Kunne ikke registrere ${job.name}: ${e.message}`);
    }
  }
}

function listJobs() {
  return Array.from(jobs.values()).map(j => ({
    name: j.name,
    settingsKey: j.settingsKey,
    active: !!j.task,
  }));
}

// ============================================================
// Register all jobs (called from index.js at boot)
// ============================================================
function registerDefaults() {
  defineJob('scanEmails',      'cron.scanEmails',      async () => {
    const { scanIncomingEmails } = require('./scanner');
    return await scanIncomingEmails();
  });
  defineJob('processIvit',     'cron.processIvit',     async () => {
    const { processIVITScraping } = require('./ivit');
    return await processIVITScraping();
  });
  defineJob('checkReminders',  'cron.checkReminders',  async () => {
    const { checkReminders } = require('./reminders');
    return await checkReminders();
  });
  defineJob('updateDashboard', 'cron.updateDashboard', async () => {
    const { updateDashboard } = require('./dashboard');
    return await updateDashboard();
  });
  defineJob('weeklyReport',    'cron.weeklyReport',    async () => {
    const { sendWeeklyReport } = require('./reminders');
    return await sendWeeklyReport();
  });
}

module.exports = {
  defineJob,
  registerDefaults,
  applyScheduleFromSettings,
  runJob,
  listJobs,
};
