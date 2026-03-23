const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '..', '.env') });

const config = {
  ivit: {
    username: process.env.IVIT_USERNAME,
    password: process.env.IVIT_PASSWORD,
    baseUrl: 'https://www.ivit.no',
  },
  webhook: {
    port: parseInt(process.env.WEBHOOK_PORT, 10) || 3000,
    secret: process.env.WEBHOOK_SECRET || null,
  },
  debug: process.env.DEBUG === 'true',
  dryRun: process.env.DRY_RUN === 'true',
  debugDir: path.resolve(__dirname, '..', 'debug'),
  puppeteer: {
    headless: process.env.HEADLESS !== 'false' ? 'new' : false,
    timeout: 30000,
  },
};

// Validate required vars
if (!config.ivit.username || !config.ivit.password) {
  console.warn('WARNING: IVIT_USERNAME and IVIT_PASSWORD are not set in .env — scraping will fail until configured.');
}

module.exports = config;
