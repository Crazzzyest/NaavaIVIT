const express = require('express');
const config = require('./config');
const { scrapeOppdrag, closeBrowser, saveScreenshot } = require('./scraper');

const app = express();
app.use(express.json());

// Global request timeout — 2 minutes max per request
app.use((req, res, next) => {
  res.setTimeout(120000, () => {
    console.error('Request timed out (120s server limit)');
    if (!res.headersSent) {
      res.status(504).json({ success: false, error: 'Request timed out (120s)' });
    }
  });
  next();
});

/**
 * Validate webhook secret if configured.
 */
function validateRequest(req, res, next) {
  if (!config.webhook.secret) {
    return next();
  }

  const authHeader = req.headers['authorization'];
  if (!authHeader || authHeader !== `Bearer ${config.webhook.secret}`) {
    return res.status(401).json({
      success: false,
      error: 'Unauthorized: invalid or missing Authorization header',
    });
  }

  next();
}

/**
 * POST /webhook — main endpoint.
 */
app.post('/webhook', validateRequest, async (req, res) => {
  const { address } = req.body || {};

  if (!address || typeof address !== 'string' || address.trim().length === 0) {
    return res.status(400).json({
      success: false,
      error: 'Missing or invalid "address" field in request body',
    });
  }

  const trimmedAddress = address.trim();
  console.log(`\n=== Webhook received: "${trimmedAddress}" ===`);

  try {
    const data = await scrapeOppdrag(trimmedAddress);

    if (data.dryRun) {
      return res.json({
        success: true,
        dryRun: true,
        address: trimmedAddress,
        url: data.url,
        message: 'Dry run complete — screenshot saved to debug/',
      });
    }

    return res.json({
      success: true,
      address: trimmedAddress,
      data,
    });
  } catch (err) {
    console.error('Scraping error:', err.message);

    // Try to save a screenshot, but don't let it hang the response
    try {
      await Promise.race([
        saveScreenshot('webhook-error'),
        new Promise((_, reject) => setTimeout(() => reject(new Error('Screenshot timed out')), 5000)),
      ]);
    } catch (screenshotErr) {
      console.error('Could not save error screenshot:', screenshotErr.message);
    }

    return res.status(500).json({
      success: false,
      error: err.message,
      address: trimmedAddress,
    });
  }
});

/**
 * Health check.
 */
app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

/**
 * Start server.
 */
const server = app.listen(config.webhook.port, () => {
  console.log(`iVit Webhook Scraper listening on port ${config.webhook.port}`);
  console.log(`Debug mode: ${config.debug}`);
  console.log(`Dry run mode: ${config.dryRun}`);
  if (config.webhook.secret) {
    console.log('Webhook secret: configured');
  }
});

/**
 * Graceful shutdown.
 */
async function shutdown(signal) {
  console.log(`\n${signal} received. Shutting down...`);
  server.close(async () => {
    await closeBrowser();
    console.log('Server closed.');
    process.exit(0);
  });

  // Force exit after 10s
  setTimeout(() => {
    console.error('Forced shutdown after timeout.');
    process.exit(1);
  }, 10000);
}

process.on('SIGTERM', () => shutdown('SIGTERM'));
process.on('SIGINT', () => shutdown('SIGINT'));
