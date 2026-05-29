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

// Validate fallback credentials. Per-request creds (sent in webhook body)
// take precedence; env vars only act as fallback.
if (!config.ivit.username || !config.ivit.password) {
  console.log('INFO: IVIT_USERNAME/IVIT_PASSWORD env vars not set. Caller must supply ivitUsername+ivitPassword in each /webhook request body, or scraping will fail.');
}

module.exports = config;
