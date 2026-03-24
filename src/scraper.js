const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');
const config = require('./config');

let browser = null;
let page = null;
let loggedIn = false;

/**
 * Ensure the debug directory exists.
 */
function ensureDebugDir() {
  if (!fs.existsSync(config.debugDir)) {
    fs.mkdirSync(config.debugDir, { recursive: true });
  }
}

/**
 * Delete debug files older than maxAgeMs (default 24 hours).
 */
function cleanupDebugFiles(maxAgeMs = 24 * 60 * 60 * 1000) {
  try {
    if (!fs.existsSync(config.debugDir)) return;
    const now = Date.now();
    const files = fs.readdirSync(config.debugDir);
    let deleted = 0;
    for (const file of files) {
      const filePath = path.join(config.debugDir, file);
      const stat = fs.statSync(filePath);
      if (now - stat.mtimeMs > maxAgeMs) {
        fs.unlinkSync(filePath);
        deleted++;
      }
    }
    if (deleted > 0) {
      console.log(`Cleaned up ${deleted} old debug file(s).`);
    }
  } catch (err) {
    console.error('Debug cleanup error:', err.message);
  }
}

/**
 * Save a screenshot for debugging.
 */
async function saveScreenshot(label = 'error') {
  ensureDebugDir();
  const ts = new Date().toISOString().replace(/[:.]/g, '-');
  const filePath = path.join(config.debugDir, `${label}_${ts}.png`);
  try {
    if (page) {
      await page.screenshot({ path: filePath, fullPage: true });
      console.log(`Screenshot saved: ${filePath}`);
    }
  } catch (err) {
    console.error('Failed to save screenshot:', err.message);
  }
  return filePath;
}

/**
 * Dump page HTML to a file when DEBUG=true.
 */
async function dumpPageHtml(label = 'debug') {
  if (!config.debug) return;
  ensureDebugDir();
  const ts = new Date().toISOString().replace(/[:.]/g, '-');
  const filePath = path.join(config.debugDir, `${label}_${ts}.html`);
  try {
    if (page) {
      const html = await page.content();
      fs.writeFileSync(filePath, html, 'utf-8');
      console.log(`HTML dump saved: ${filePath}`);
    }
  } catch (err) {
    console.error('Failed to dump HTML:', err.message);
  }
}

/**
 * Wait for the Angular SPA to finish rendering.
 * Waits for the spinner to disappear and real content to appear.
 */
async function waitForSpaReady(p, timeout = config.puppeteer.timeout) {
  console.log('Waiting for SPA to render...');

  // Wait for the spinner to disappear (if present)
  try {
    await p.waitForFunction(() => {
      const spinner = document.querySelector('.spinner');
      if (!spinner) return true;
      return spinner.offsetParent === null; // hidden
    }, { timeout });
  } catch {
    // Spinner might not exist, that's ok
  }

  // Wait for Angular's app-root to have meaningful content
  await p.waitForFunction(() => {
    const appRoot = document.querySelector('app-root');
    if (!appRoot) return true; // not an Angular app
    // Check that app-root has more than just the spinner div
    const children = appRoot.children;
    if (children.length === 0) return false;
    // If only child is a spinner div, still loading
    if (children.length === 1 && children[0].classList.contains('spinner')) return false;
    // Check for meaningful text content (more than just whitespace)
    const text = appRoot.innerText?.trim() || '';
    return text.length > 20;
  }, { timeout });

  // Extra settle time for Angular change detection
  await new Promise(r => setTimeout(r, 1500));
  console.log('SPA content loaded.');
}

/**
 * Navigate to a URL and wait for SPA content to load.
 */
async function gotoAndWait(p, url) {
  await p.goto(url, { waitUntil: 'networkidle2' });
  await waitForSpaReady(p);
}

/**
 * Get or launch the browser instance.
 */
async function getBrowser() {
  if (browser && browser.connected) {
    return browser;
  }
  console.log('Launching browser...');
  const launchOptions = {
    headless: config.puppeteer.headless,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
  };
  // Use system Chromium in Docker (set via PUPPETEER_EXECUTABLE_PATH)
  if (process.env.PUPPETEER_EXECUTABLE_PATH) {
    launchOptions.executablePath = process.env.PUPPETEER_EXECUTABLE_PATH;
  }
  browser = await puppeteer.launch(launchOptions);
  loggedIn = false;
  page = null;
  return browser;
}

/**
 * Get or create the working page. Reuses a single page/tab.
 */
async function getPage() {
  const b = await getBrowser();
  if (page && !page.isClosed()) {
    return page;
  }
  const pages = await b.pages();
  page = pages[0] || await b.newPage();
  page.setDefaultTimeout(config.puppeteer.timeout);
  page.setDefaultNavigationTimeout(config.puppeteer.timeout);
  await page.setViewport({ width: 1280, height: 900 });
  loggedIn = false;
  return page;
}

/**
 * Check if we're on a login page (session expired).
 */
async function isOnLoginPage(p) {
  const url = p.url();
  // Check URL patterns that indicate login page
  if (url.includes('login') || url.includes('Login') || url.includes('signin') || url.includes('SignIn')) {
    return true;
  }
  // Check for login form elements
  const hasLoginForm = await p.evaluate(() => {
    const inputs = document.querySelectorAll('input[type="password"]');
    return inputs.length > 0;
  }).catch(() => false);
  return hasLoginForm;
}

/**
 * Log into iVit.
 */
async function login() {
  const p = await getPage();

  console.log('Navigating to iVit...');
  await gotoAndWait(p, config.ivit.baseUrl);

  await dumpPageHtml('pre-login');

  // Check if we're already logged in
  if (!(await isOnLoginPage(p))) {
    console.log('Already logged in (session reuse).');
    loggedIn = true;
    return;
  }

  console.log('Logging in...');

  // Wait for the login form — try common selectors
  // First, inspect what's actually on the page
  const loginSelectors = await p.evaluate(() => {
    const forms = document.querySelectorAll('form');
    const inputs = document.querySelectorAll('input');
    const buttons = document.querySelectorAll('button, input[type="submit"]');
    return {
      formCount: forms.length,
      inputs: Array.from(inputs).map(i => ({
        type: i.type,
        name: i.name,
        id: i.id,
        placeholder: i.placeholder,
        className: i.className,
      })),
      buttons: Array.from(buttons).map(b => ({
        type: b.type,
        text: b.textContent?.trim(),
        id: b.id,
        className: b.className,
      })),
    };
  });

  if (config.debug) {
    console.log('Login page elements:', JSON.stringify(loginSelectors, null, 2));
  }

  // Find username field — try multiple strategies
  const usernameSelectors = [
    'input[name="username"]',
    'input[name="Username"]',
    'input[name="email"]',
    'input[name="Email"]',
    'input[name="user"]',
    'input[name="brukernavn"]',
    'input[id="username"]',
    'input[id="Username"]',
    'input[id="email"]',
    'input[id="Email"]',
    'input[type="email"]',
    'input[type="text"]:not([type="hidden"])',
  ];

  let usernameField = null;
  for (const sel of usernameSelectors) {
    usernameField = await p.$(sel);
    if (usernameField) {
      console.log(`Found username field: ${sel}`);
      break;
    }
  }

  if (!usernameField) {
    await saveScreenshot('login-no-username-field');
    throw new Error('Could not find username input field on login page');
  }

  // Find password field
  const passwordField = await p.$('input[type="password"]');
  if (!passwordField) {
    await saveScreenshot('login-no-password-field');
    throw new Error('Could not find password input field on login page');
  }

  // Clear and type credentials
  await usernameField.click({ clickCount: 3 });
  await usernameField.type(config.ivit.username);

  await passwordField.click({ clickCount: 3 });
  await passwordField.type(config.ivit.password);

  // Find and click submit button
  const submitSelectors = [
    'button[type="submit"]',
    'input[type="submit"]',
    'button:has-text("Logg inn")',
    'button:has-text("Login")',
    'button:has-text("Sign in")',
    '#loginButton',
    '.login-button',
    'form button',
  ];

  let submitted = false;
  for (const sel of submitSelectors) {
    try {
      const btn = await p.$(sel);
      if (btn) {
        console.log(`Clicking submit button: ${sel}`);
        await Promise.all([
          p.waitForNavigation({ waitUntil: 'networkidle2', timeout: config.puppeteer.timeout }).catch(() => {}),
          btn.click(),
        ]);
        submitted = true;
        break;
      }
    } catch {
      continue;
    }
  }

  if (!submitted) {
    // Fallback: press Enter in password field
    console.log('No submit button found, pressing Enter...');
    await Promise.all([
      p.waitForNavigation({ waitUntil: 'networkidle2', timeout: config.puppeteer.timeout }).catch(() => {}),
      passwordField.press('Enter'),
    ]);
  }

  // Verify login success — wait for SPA to render post-login
  await waitForSpaReady(p);
  await dumpPageHtml('post-login');

  if (await isOnLoginPage(p)) {
    await saveScreenshot('login-failed');
    throw new Error('Login failed — still on login page after submitting credentials');
  }

  console.log('Login successful.');
  loggedIn = true;
}

/**
 * Ensure we're logged in, re-login if session expired.
 */
async function ensureLoggedIn() {
  const p = await getPage();
  if (!loggedIn || (await isOnLoginPage(p))) {
    await login();
  }
}

/**
 * Search for an oppdrag by address using the global search.
 */
async function findOppdrag(address) {
  await ensureLoggedIn();
  const p = await getPage();

  console.log(`Searching for oppdrag: "${address}"`);

  // The homepage (/) is the oppdrag list. Ensure we're there.
  const currentUrl = p.url();
  if (!currentUrl.startsWith(config.ivit.baseUrl)) {
    await gotoAndWait(p, config.ivit.baseUrl);
  }

  // Re-check login after navigation
  if (await isOnLoginPage(p)) {
    loggedIn = false;
    await login();
    return findOppdrag(address);
  }

  // Use the global search field: ivit-global-search input
  const searchSelector = 'ivit-global-search input[type="text"], input[placeholder*="søk" i]';
  await p.waitForSelector(searchSelector, { timeout: config.puppeteer.timeout });
  const searchField = await p.$(searchSelector);

  if (!searchField) {
    await saveScreenshot('no-search-field');
    throw new Error('Could not find the global search field on the page');
  }

  console.log('Found global search field, typing address...');

  // Clear and type the address
  await searchField.click({ clickCount: 3 });
  await searchField.type('', { delay: 0 }); // clear
  await searchField.evaluate(el => el.value = '');
  // Extract the main address part (street + number) for search
  const mainAddress = address.split(',')[0].trim();
  await searchField.type(mainAddress, { delay: 50 });

  // Wait for search results dropdown to appear
  console.log('Waiting for search results...');
  await p.waitForSelector('.search-result', { timeout: 10000 }).catch(() => {});
  await new Promise(r => setTimeout(r, 2000)); // let results settle

  await dumpPageHtml('search-results');
  await saveScreenshot('search-results');

  // Find a matching result in the search dropdown
  const matchResult = await p.evaluate((addr) => {
    const addrLower = addr.toLowerCase();
    // Look in search results for mission cards
    const cards = document.querySelectorAll('.search-result .mission-card, .search-result .mat-card, .search-result .row');
    for (const card of cards) {
      const text = card.textContent?.toLowerCase() || '';
      if (text.includes(addrLower)) {
        return { found: true, text: card.textContent.trim().substring(0, 300) };
      }
    }

    // Also try any clickable element in search results
    const results = document.querySelectorAll('.search-result a, .search-result [class*="card"], .search-result [class*="mission"]');
    for (const el of results) {
      const text = el.textContent?.toLowerCase() || '';
      if (text.includes(addrLower)) {
        return { found: true, text: el.textContent.trim().substring(0, 300) };
      }
    }

    // Fuzzy: match street name without number
    const streetName = addrLower.split(/\d/)[0].trim();
    if (streetName.length > 3) {
      const allResults = document.querySelectorAll('.search-result *');
      for (const el of allResults) {
        const text = el.textContent?.toLowerCase() || '';
        if (text.includes(streetName) && el.closest('.mission-card, .mat-card, a, [class*="mission"]')) {
          return { found: true, text: el.closest('.mission-card, .mat-card, a, [class*="mission"]').textContent.trim().substring(0, 300), fuzzy: true };
        }
      }
    }

    // Collect what results are visible for debugging
    const visibleResults = [];
    const allCards = document.querySelectorAll('.search-result .mission-card, .search-result .row:not(.mission-header)');
    allCards.forEach((c, i) => {
      if (i < 5) visibleResults.push(c.textContent?.trim().substring(0, 200));
    });

    return { found: false, visibleResults };
  }, mainAddress);

  if (!matchResult.found) {
    await saveScreenshot('no-match');
    const debugInfo = matchResult.visibleResults?.length
      ? ` Visible results: ${matchResult.visibleResults.join(' | ')}`
      : '';
    throw new Error(`No oppdrag found for address: ${address}.${debugInfo}`);
  }

  console.log('Found matching oppdrag:', matchResult.text);

  // Click the matching search result to navigate to detail page
  await p.evaluate((addr) => {
    const addrLower = addr.toLowerCase();
    const streetName = addrLower.split(/\\d/)[0].trim();

    const clickables = document.querySelectorAll('.search-result .mission-card, .search-result .mat-card, .search-result a, .search-result .row:not(.mission-header)');
    for (const el of clickables) {
      const text = el.textContent?.toLowerCase() || '';
      if (text.includes(addrLower) || (streetName.length > 3 && text.includes(streetName))) {
        el.click();
        return;
      }
    }
  }, mainAddress);

  // Wait for navigation to the detail page
  await new Promise(r => setTimeout(r, 1000));
  await p.waitForNavigation({ waitUntil: 'networkidle2', timeout: 10000 }).catch(() => {});
  await waitForSpaReady(p);

  await dumpPageHtml('oppdrag-detail');
  await saveScreenshot('oppdrag-detail');

  return p;
}

/**
 * Handle the "Endre status" popup by ALWAYS clicking "Nei".
 * This popup appears when navigating to the tilstandsrapport and asks
 * if you want to change the oppdrag status to "Under arbeid".
 * We MUST click Nei to avoid modifying the oppdrag status.
 */
async function handleEndreStatusPopup(p) {
  try {
    // Wait briefly for the popup to appear
    await p.waitForFunction(() => {
      const buttons = document.querySelectorAll('button');
      return Array.from(buttons).some(b => b.textContent?.trim() === 'Nei');
    }, { timeout: 5000 });

    console.log('"Endre status" popup detected — clicking Nei...');

    // Click the "Nei" button to dismiss without changing status
    await p.evaluate(() => {
      const buttons = document.querySelectorAll('button');
      for (const btn of buttons) {
        if (btn.textContent?.trim() === 'Nei') {
          btn.click();
          return;
        }
      }
    });

    // Wait for popup to close
    await new Promise(r => setTimeout(r, 1000));
    console.log('Popup dismissed (clicked Nei).');
  } catch {
    // No popup appeared — that's fine
    console.log('No "Endre status" popup appeared.');
  }
}

/**
 * Navigate to the tilstandsrapport (condition report) from the oppdrag detail page.
 * Clicks the report button and handles the "Endre status" popup.
 */
async function navigateToTilstandsrapport(p) {
  console.log('Navigating to tilstandsrapport...');

  // Click the report button: .md-mission-report-name-wrapper
  const reportButton = await p.$('.md-mission-report-name-wrapper');
  if (!reportButton) {
    console.log('Tilstandsrapport button not found on this oppdrag.');
    return false;
  }

  await reportButton.click();

  // Wait for potential navigation and SPA render
  await new Promise(r => setTimeout(r, 2000));

  // CRITICAL: Handle the "Endre status" popup — ALWAYS click Nei
  await handleEndreStatusPopup(p);

  // Wait for the report page to load
  await p.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }).catch(() => {});
  await waitForSpaReady(p);

  await dumpPageHtml('tilstandsrapport');
  await saveScreenshot('tilstandsrapport');

  console.log('Tilstandsrapport page loaded.');
  return true;
}

/**
 * Extract data from the tilstandsrapport page.
 * Reads the markedsverdi checkbox (mat-checkbox with aria-checked).
 */
async function extractTilstandsrapportData(p) {
  console.log('Extracting data from tilstandsrapport...');

  const reportData = await p.evaluate(() => {
    const result = {};

    // Look for a mat-checkbox near a "Markedsverdi" label.
    // The checkbox input has aria-checked="true" or "false".
    const allLabels = document.querySelectorAll('label, span, div, mat-label');
    let markedsverdiCheckbox = null;

    for (const el of allLabels) {
      const text = el.textContent?.trim().toLowerCase() || '';
      if (text.includes('markedsverdi') && !text.includes('uten') && text.length < 40) {
        // Look for a mat-checkbox in the same parent or nearby
        const parent = el.closest('mat-checkbox, .mat-checkbox') || el.parentElement?.closest('mat-checkbox, .mat-checkbox');
        if (parent) {
          markedsverdiCheckbox = parent.querySelector('input[type="checkbox"]');
          break;
        }
        // Also check siblings
        const container = el.parentElement;
        if (container) {
          markedsverdiCheckbox = container.querySelector('mat-checkbox input[type="checkbox"], .mat-checkbox input[type="checkbox"]');
          if (markedsverdiCheckbox) break;
        }
      }
    }

    // Fallback: find any mat-checkbox whose associated label contains "markedsverdi"
    if (!markedsverdiCheckbox) {
      const allCheckboxes = document.querySelectorAll('mat-checkbox, .mat-checkbox');
      for (const cb of allCheckboxes) {
        const cbText = cb.textContent?.trim().toLowerCase() || '';
        if (cbText.includes('markedsverdi')) {
          markedsverdiCheckbox = cb.querySelector('input[type="checkbox"]');
          break;
        }
      }
    }

    if (markedsverdiCheckbox) {
      const isChecked = markedsverdiCheckbox.getAttribute('aria-checked') === 'true'
                     || markedsverdiCheckbox.checked;
      result.med_markedsverdi = isChecked ? 'Ja' : 'Nei';
      console.log('Markedsverdi checkbox found, aria-checked=' + markedsverdiCheckbox.getAttribute('aria-checked'));
    } else {
      // No checkbox found — default to Nei (not checked)
      result.med_markedsverdi = 'Nei';
      console.log('No markedsverdi checkbox found — defaulting to Nei');
    }

    return result;
  });

  return reportData;
}

/**
 * Extract oppdrag data from the detail page.
 * Reads the Overview tab, then Befaringer tab, then the Tilstandsrapport.
 */
async function extractOppdragData() {
  const p = await getPage();

  if (config.dryRun) {
    console.log('DRY RUN: Taking screenshot of oppdrag page and skipping extraction.');
    await saveScreenshot('dry-run-oppdrag');
    return { dryRun: true, url: p.url() };
  }

  // --- Extract data from Overview tab ---
  console.log('Extracting data from Overview tab...');
  const overviewData = await p.evaluate(() => {
    const result = {};

    // "Deres ref" (fakturareferanse) — input with data-fieldname
    const deresRefInput = document.querySelector('input[data-fieldname="Deres ref"]');
    result.fakturareferanse = deresRefInput?.value?.trim() || null;

    // Markedsverdi — will be determined from tilstandsrapport checkbox
    result.med_markedsverdi = null;

    return result;
  });

  // --- Navigate to Befaringer tab for befaring data ---
  console.log('Navigating to Befaringer tab...');
  let befaringData = { befaring_dato: null, befaring_klokkeslett: null };

  const befaringerLink = await p.$('a[routerlink="inspections"], a[href*="inspections"]');
  if (befaringerLink) {
    await befaringerLink.click();
    await new Promise(r => setTimeout(r, 1000));
    await p.waitForNavigation({ waitUntil: 'networkidle2', timeout: 10000 }).catch(() => {});
    await waitForSpaReady(p);

    await dumpPageHtml('befaringer-tab');
    await saveScreenshot('befaringer-tab');

    befaringData = await p.evaluate(() => {
      const result = {};

      // The befaring date is in a mat-datepicker-input (readonly input set by Angular)
      const datePicker = document.querySelector('.mat-datepicker-input, input[matinput][readonly]');
      if (datePicker && datePicker.value) {
        result.befaring_dato = datePicker.value;
      }

      // The time values are in mat-chip elements or time inputs
      const chipTexts = [];
      const chips = document.querySelectorAll('mat-chip, .mat-chip, .mat-chip-list input');
      chips.forEach(c => {
        const text = c.textContent?.trim().replace(/[×x✕close]/gi, '').trim();
        if (text && text.match(/^\d{2}:\d{2}$/)) {
          chipTexts.push(text);
        }
      });

      // Also check all visible text for time patterns near the date
      if (chipTexts.length === 0) {
        const pageText = document.body.innerText || '';
        const timeMatches = pageText.match(/\d{2}:\d{2}/g);
        if (timeMatches) {
          chipTexts.push(...timeMatches.slice(0, 2));
        }
      }

      if (chipTexts.length >= 2) {
        result.befaring_klokkeslett = `${chipTexts[0]} - ${chipTexts[1]}`;
      } else if (chipTexts.length === 1) {
        result.befaring_klokkeslett = chipTexts[0];
      }

      // Read time inputs with HH:mm placeholder (Angular Material time pickers)
      if (!result.befaring_klokkeslett) {
        const timeInputs = document.querySelectorAll('input[placeholder="HH:mm"], input[data-placeholder="HH:mm"]');
        const times = [];
        timeInputs.forEach(input => {
          const val = input.value?.trim();
          if (val && val.match(/^\d{2}:\d{2}$/)) {
            times.push(val);
          }
        });
        if (times.length >= 2) {
          result.befaring_klokkeslett = `${times[0]} - ${times[1]}`;
        } else if (times.length === 1) {
          result.befaring_klokkeslett = times[0];
        }
      }

      if (!result.befaring_dato) result.befaring_dato = null;
      if (!result.befaring_klokkeslett) result.befaring_klokkeslett = null;

      return result;
    });
  } else {
    console.log('Befaringer tab not found.');
  }

  // --- Navigate back to Overview to access tilstandsrapport ---
  console.log('Navigating back to Overview tab...');
  const oversiktLink = await p.$('a[routerlink="overview"], a[href*="overview"]');
  if (oversiktLink) {
    await oversiktLink.click();
    await new Promise(r => setTimeout(r, 1000));
    await p.waitForNavigation({ waitUntil: 'networkidle2', timeout: 10000 }).catch(() => {});
    await waitForSpaReady(p);
  } else {
    // Try clicking the Oversikt tab by text
    await p.evaluate(() => {
      const links = document.querySelectorAll('a');
      for (const link of links) {
        if (link.textContent?.trim() === 'Oversikt') {
          link.click();
          return;
        }
      }
    });
    await new Promise(r => setTimeout(r, 2000));
    await waitForSpaReady(p);
  }

  // --- Navigate to Tilstandsrapport for markedsverdi ---
  let reportData = { med_markedsverdi: null };
  const reportOpened = await navigateToTilstandsrapport(p);
  if (reportOpened) {
    reportData = await extractTilstandsrapportData(p);
  }

  return {
    befaring_dato: befaringData.befaring_dato,
    befaring_klokkeslett: befaringData.befaring_klokkeslett,
    fakturareferanse: overviewData.fakturareferanse,
    med_markedsverdi: reportData.med_markedsverdi || 'Nei',
  };
}

/**
 * Wrap a promise with a timeout. Rejects if the promise doesn't resolve within `ms` milliseconds.
 */
function withTimeout(promise, ms, label = 'Operation') {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => {
      reject(new Error(`${label} timed out after ${Math.round(ms / 1000)}s`));
    }, ms);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}

/**
 * Main scraping function called by the webhook handler.
 * Wrapped in a 90-second timeout so it never hangs indefinitely.
 */
async function scrapeOppdrag(address) {
  cleanupDebugFiles();

  const SCRAPE_TIMEOUT = 90000; // 1.5 minutes

  return withTimeout(
    (async () => {
      await findOppdrag(address);
      const data = await extractOppdragData();
      return data;
    })(),
    SCRAPE_TIMEOUT,
    `Scraping "${address}"`
  );
}

/**
 * Close the browser instance.
 */
async function closeBrowser() {
  if (browser) {
    console.log('Closing browser...');
    try {
      await browser.close();
    } catch (err) {
      console.error('Error closing browser:', err.message);
    }
    browser = null;
    page = null;
    loggedIn = false;
  }
}

module.exports = {
  scrapeOppdrag,
  closeBrowser,
  saveScreenshot,
};
