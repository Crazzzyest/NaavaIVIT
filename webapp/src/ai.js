const config = require('./config');

async function callOpenAI(systemPrompt, userPrompt) {
  if (!config.openai.apiKey) return null;

  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${config.openai.apiKey}`,
    },
    body: JSON.stringify({
      model: config.openai.model,
      temperature: 0,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt },
      ],
    }),
  });

  if (!res.ok) {
    console.error(`OpenAI API error: HTTP ${res.status}`);
    return null;
  }

  const data = await res.json();
  return data.choices[0].message.content.trim();
}

function parseJsonResponse(text) {
  const match = text.match(/\{[\s\S]*\}/);
  if (!match) return null;
  try {
    return JSON.parse(match[0]);
  } catch {
    return null;
  }
}

async function assessAndParseEmail(subject, body, from) {
  const emailMatch = (from || '').match(/<(.+?)>/);
  const senderEmail = emailMatch ? emailMatch[1] : (from || '').trim();
  const senderName = ((from || '').match(/^"?([^"<]+)"?\s*</) || [])[1] || senderEmail;

  const systemPrompt =
    'Du er et system som vurderer innkommende e-poster for et takstfirma i Norge (Naava Takst). ' +
    'Du skal avgjøre om e-posten er en NY bestilling eller forespørsel om et takstoppdrag. ' +
    '\n\nSett "relevant": false for ALLE disse tilfellene:' +
    '\n- Oppfølging, statusoppdatering eller melding om et eksisterende oppdrag' +
    '\n- Oversending, vedlegg eller videresending av en ferdig eller tidligere rapport ' +
    '(signalord: "oppdatert", "tilbakemelding", "tidligere", "revidert")' +
    '\n- Avlysning, utsettelse eller omplanlegging av befaring' +
    '\n- Spørsmål om et pågående oppdrag' +
    '\n- Videresending av dokumenter til et oppdrag som allerede eksisterer' +
    '\n- Spam, nyhetsbrev, fakturaer, kvitteringer, automatiske varsler' +
    '\n- Intern kommunikasjon eller svar på e-post fra Naava Takst selv' +
    '\n\nSett "relevant": true KUN hvis e-posten er en tydelig ny forespørsel eller bestilling ' +
    'der noen ønsker at Naava Takst skal utføre en befaring eller rapport på en eiendom. ' +
    'Tvilstilfeller skal vurderes som IKKE relevante. ' +
    'Svar KUN med gyldig JSON.';

  const userPrompt =
    'E-post mottatt av Naava Takst:\n\n' +
    `Fra: ${from}\nEmne: ${subject}\n\nInnhold:\n${(body || '').substring(0, 3000)}\n\n` +
    'Svar med dette JSON-skjemaet:\n' +
    '{\n' +
    '  "relevant": true,\n' +
    '  "begrunnelse": "<en setning>",\n' +
    '  "oppdragstype": "<Tilstandsrapport | Tilstandsrapport m/markedsverdi | Skadetakst | Reklamasjon | Vurderingsoppdrag | Bistand overtagelse | Fukt-/fuktskadevurdering | Byggelånskontroll | Forhåndstakst | Verditakst | Annet>",\n' +
    '  "rapporttype": "<Tilstandsrapport m/teknisk og markedsverdi | Tilstandsrapport | Skadetakstrapport | Reklamasjonsrapport | Vurderingsrapport | Overtagelsesrapport | Annen rapport>",\n' +
    '  "adresse": "<gateadresse, postnr, poststed>",\n' +
    '  "oppdragsgiver": "<firmanavn eller personnavn>",\n' +
    '  "selger": "<selgers fulle navn>",\n' +
    '  "selgerTlf": "<telefonnummer, kun sifre>",\n' +
    '  "selgerEpost": "<e-post>",\n' +
    '  "megler": "<meglerens navn>",\n' +
    '  "meglerEpost": "<meglerens e-post>",\n' +
    '  "fakturaRef": "<referansenummer>",\n' +
    '  "fakturaSendesTil": "<hvem faktura sendes til>",\n' +
    '  "notater": "<andre relevante opplysninger>"\n' +
    '}\n' +
    'Feltene kan være tomme strenger. Ikke gjett.';

  const text = await callOpenAI(systemPrompt, userPrompt);
  if (!text) return null;

  const result = parseJsonResponse(text);
  if (!result || !result.relevant) return null;

  return {
    kilde: 'Megler-epost',
    oppdragstype: result.oppdragstype || 'Annet',
    rapporttype: result.rapporttype || 'Annen rapport',
    adresse: String(result.adresse || '').trim(),
    oppdragsgiver: String(result.oppdragsgiver || '').trim(),
    selger: String(result.selger || '').trim(),
    selgerTlf: String(result.selgerTlf || '').replace(/[^\d+]/g, ''),
    selgerEpost: String(result.selgerEpost || '').trim(),
    megler: String(result.megler || senderName).trim(),
    meglerEpost: String(result.meglerEpost || senderEmail).trim(),
    fakturaRef: String(result.fakturaRef || '').trim(),
    fakturaSendesTil: String(result.fakturaSendesTil || '').trim(),
    notater: String(result.notater || '').trim(),
  };
}

async function calculateDistance(destinationAddress) {
  const prompt =
    `Jeg trenger en estimert kjøreavstand (en vei) mellom disse to adressene i Norge.\n\n` +
    `Fra: ${config.travel.baseAddress}\nTil: ${destinationAddress}\n\n` +
    'Svar KUN med et JSON-objekt:\n' +
    '{"km_en_vei": <tall>, "km_tur_retur": <tall>, "estimert_tid_min": <tall>}\n\n' +
    'Gi realistisk kjøreavstand (ikke luftlinje).';

  const text = await callOpenAI('Du svarer kun med gyldig JSON.', prompt);
  if (!text) return null;

  const result = parseJsonResponse(text);
  if (!result) return null;

  return {
    kmEnVei: Number(result.km_en_vei) || 0,
    kmTurRetur: Number(result.km_tur_retur) || 0,
    estimertTidMin: Number(result.estimert_tid_min) || 0,
  };
}

module.exports = { assessAndParseEmail, calculateDistance };
