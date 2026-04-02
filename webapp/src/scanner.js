const config = require('./config');
const google = require('./google');
const { assessAndParseEmail } = require('./ai');
const { sanitizeParsedData } = require('./utils');
const { createNewOppdrag } = require('./oppdrag');

function parseIVITEmail(subject, body, from) {
  const text = (body || '').replace(/\r/g, '');
  const result = {
    kilde: 'IVIT',
    oppdragstype: 'Tilstandsrapport',
    rapporttype: '',
    adresse: '',
    oppdragsgiver: '',
    selger: '',
    selgerTlf: '',
    selgerEpost: '',
    megler: '',
    meglerEpost: '',
    fakturaRef: '',
    fakturaSendesTil: '',
    notater: '',
  };

  const ordreMatch = text.match(/(?:^|\n)\s*[Oo]rdre\s*nummer\s*[:\t ]?\s*([A-Z0-9-]+)\s*(?:\n|$)/);
  if (ordreMatch) result.fakturaRef = ordreMatch[1].trim();

  let adresseMatch = text.match(/følgende eiendom[:\s]*\n\s*(.+?)(?:\n|,\s*gnr)/i);
  if (adresseMatch) result.adresse = adresseMatch[1].trim();
  if (!result.adresse) {
    adresseMatch = text.match(/[Aa]dress(?:e[n]?\s+er|e[:\s])\s*(.+?)(?:\n|$)/i);
    if (adresseMatch) result.adresse = adresseMatch[1].trim();
  }
  if (!result.adresse) {
    adresseMatch = text.match(/([A-ZÆØÅ][a-zæøåA-ZÆØÅ]*(?:ringen|veien|gata|gaten|vegen|stien|bakken|lia|haugen|åsen|berget|stranda|plassen|torget|brygga|bøen|øen|tunet|marka|jordet|løkka)\s+\d+[A-Za-z]?(?:\s*,?\s*\d{4}\s+[A-ZÆØÅa-zæøå]+)?)/);
    if (adresseMatch) result.adresse = adresseMatch[1].trim();
  }

  const hilsenBlock = text.match(/[Vv]ennlig hilsen\s*\n([\s\S]*?)(?:\n\s*\n|$)/);
  if (hilsenBlock) {
    const lines = hilsenBlock[1].split('\n').map(l => l.trim()).filter(Boolean);
    if (lines.length >= 1) result.megler = lines[0];
    if (lines.length >= 2) result.fakturaSendesTil = lines[1];
  }

  const emailMatch = (from || '').match(/<(.+?)>/);
  result.meglerEpost = emailMatch ? emailMatch[1] : (from || '').trim();

  return result;
}

async function scanIncomingEmails() {
  console.log('Starting email scan...');

  const labelId = await google.getOrCreateLabel('Takst-Behandlet');

  let query;
  if (config.testMode) {
    const keywords = config.triggerKeywords.map(kw => `"${kw}"`).join(' OR ');
    query = `from:${config.email.testSender} (${keywords}) -label:Takst-Behandlet after:2026/02/05`;
    console.log(`TEST MODE: Scanning from ${config.email.testSender}`);
  } else {
    const keywords = config.triggerKeywords.map(k => `"${k}"`).join(' OR ');
    const meglerQuery = `(${keywords})`;
    const ivitQuery = '(from:no-reply.takst@ivit.no)';
    query = `(${meglerQuery} OR ${ivitQuery}) -label:Takst-Behandlet newer_than:2d`;
  }

  console.log('Gmail query:', query);

  let threads;
  try {
    threads = await google.searchGmail(query, 20);
  } catch (e) {
    console.error('Gmail search failed:', e.message);
    return { processed: 0, error: e.message };
  }

  console.log(`Found ${threads.length} threads`);
  let processed = 0;

  for (const threadInfo of threads) {
    try {
      const thread = await google.getThread(threadInfo.id);
      const messages = thread.messages || [];
      if (messages.length === 0) continue;

      const msg = google.parseMessagePayload(messages[messages.length - 1]);

      console.log(`--- Thread: "${msg.subject}" from ${msg.from}`);

      if (config.testMode) {
        if (!msg.from.toLowerCase().includes(config.email.testSender.toLowerCase())) {
          console.log('  Skipping - wrong sender in test mode');
          continue;
        }
      }

      const isIVIT = (msg.from || '').toLowerCase().includes('no-reply.takst@ivit.no');

      if (isIVIT && (msg.body || '').includes('Jacob Engholm Holen')) {
        console.log('  Skipping IVIT - contains own signature');
        await google.addLabelToThread(threadInfo.id, labelId);
        await google.markThreadRead(threadInfo.id);
        continue;
      }

      let parsed = null;

      if (isIVIT) {
        parsed = parseIVITEmail(msg.subject, msg.body, msg.from);
      } else {
        parsed = await assessAndParseEmail(msg.subject, msg.body, msg.from);
        if (!parsed) {
          console.log('  AI assessed: not relevant');
          await google.addLabelToThread(threadInfo.id, labelId);
          await google.markThreadRead(threadInfo.id);
          continue;
        }
        console.log(`  AI assessed: relevant - ${parsed.oppdragstype}`);
      }

      if (config.testMode) {
        parsed = sanitizeParsedData(parsed);
      }

      console.log(`  Address: "${parsed.adresse}"`);
      await createNewOppdrag(parsed, msg.date);
      processed++;

      await google.addLabelToThread(threadInfo.id, labelId);
      await google.markThreadRead(threadInfo.id);
      console.log('  Done with thread');

    } catch (err) {
      console.error('Error processing thread:', err.message);
    }
  }

  console.log(`Scan complete. Processed ${processed} threads.`);
  return { processed, total: threads.length };
}

module.exports = { scanIncomingEmails, parseIVITEmail };
