const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '..', '.env') });

const config = {
  testMode: process.env.TEST_MODE === 'true',
  demoMode: process.env.DEMO_MODE === 'true',
  webappCronEnabled: process.env.WEBAPP_CRON_ENABLED === 'true',

  google: {
    clientId: process.env.GOOGLE_CLIENT_ID,
    clientSecret: process.env.GOOGLE_CLIENT_SECRET,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
  },

  // Separate OAuth client for end-user login (whitelist enforced).
  // MUST NOT share credentials with the service-account refresh-token above.
  loginOAuth: {
    clientId: process.env.LOGIN_OAUTH_CLIENT_ID,
    clientSecret: process.env.LOGIN_OAUTH_CLIENT_SECRET,
    callbackUrl: process.env.LOGIN_OAUTH_CALLBACK_URL || 'http://localhost:3001/auth/google/callback',
  },

  session: {
    secret: process.env.SESSION_SECRET || 'dev-secret-change-me',
    maxAgeMs: 30 * 24 * 60 * 60 * 1000, // 30 days
  },

  allowedEmails: (process.env.ALLOWED_EMAILS || '')
    .split(',')
    .map(s => s.trim().toLowerCase())
    .filter(Boolean),

  // Accountant role — these emails log in but get a read-only "Regnskap" view
  // (PowerOffice copy buttons), not the full admin/edit interface.
  accountantEmails: (process.env.ACCOUNTANT_EMAILS || '')
    .split(',')
    .map(s => s.trim().toLowerCase())
    .filter(Boolean),

  sheet: {
    id: '1VEzaCNEkvbWYZf0UOrj6IG5PUQB3hL1enB24PvsQ1MI',
    name: 'Oppdragslogg',
    dashboardName: 'Dashboard',
  },

  drive: {
    rootFolderId: '1nDgJrHWnEdkkG1OR90vGmHby4CTYA7J_',
  },

  email: {
    ownerEmail: 'jacob@naava.no',
    accountantEmail: 'regnskap@naava.no',
    testSender: 'edsongreistad99@gmail.com',
  },

  openai: {
    apiKey: process.env.OPENAI_API_KEY || '',
    model: process.env.OPENAI_MODEL || 'gpt-5.2',
  },

  ivit: {
    webhookUrl: process.env.IVIT_WEBHOOK_URL || 'https://naavaivit.sliplane.app/webhook',
    webhookSecret: process.env.IVIT_WEBHOOK_SECRET || '',
  },

  travel: {
    baseAddress: 'Postveien 15, 6018 Ålesund',
    satsEksMva: 10,
    inkludertKm: 50,
  },

  triggerKeywords: [
    'tilstandsrapport', 'takst', 'befaring', 'verdivurdering',
    'markedsverdi', 'boligsalgsrapport', 'ordre i ivit',
    'skaderapport', 'reklamasjon', 'skadetakst', 'vurderingsoppdrag',
    'overtagelse', 'bistand',
  ],

  reminderHours: 2,
  urgentHours: 24,
  mvaRate: 0.25,

  port: parseInt(process.env.PORT, 10) || 3001,
};

module.exports = config;
