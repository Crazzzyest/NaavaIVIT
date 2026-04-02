const config = require('./config');
const { COL } = require('./columns');
const { formatCurrency, extractAddressFromHyperlink } = require('./utils');

function buildFakturaEmail(rowData, datoStr) {
  const oppdragsnr = rowData[COL.OPPDRAGSNR - 1] || '';
  let adresseFull = extractAddressFromHyperlink(rowData[COL.ADRESSE - 1]);

  const selger = rowData[COL.SELGER - 1] || '';
  const selgerTlf = rowData[COL.SELGER_TLF - 1] || '';
  const selgerEpost = rowData[COL.SELGER_EPOST - 1] || '';
  const megler = rowData[COL.MEGLER - 1] || '';
  const meglerEpost = rowData[COL.MEGLER_EPOST - 1] || '';
  const fakturaRef = rowData[COL.FAKTURA_REF - 1] || '';
  const fakturaTil = rowData[COL.FAKTURA_SENDES_TIL - 1] || '';
  const rapporttype = rowData[COL.RAPPORTTYPE - 1] || rowData[COL.OPPDRAGSTYPE - 1] || 'Oppdrag';
  const prisEks = parseFloat(rowData[COL.PRIS_EKS - 1]) || 0;
  const reiseEks = parseFloat(rowData[COL.REISE_EKS - 1]) || 0;
  const kommentarRegnskap = String(rowData[COL.KOMMENTAR_REGNSKAP - 1] || '').trim();

  let gate = adresseFull;
  let postnr = '';
  let poststed = '';
  const match = adresseFull.match(/(.+?)(?:,\s*|\s+)(\d{4})\s+(.+)/);
  if (match) {
    gate = match[1].trim();
    postnr = match[2].trim();
    poststed = match[3].trim();
  }

  const isMegler = fakturaTil && selger && !fakturaTil.toLowerCase().includes(selger.toLowerCase());
  const kundeNavn = fakturaTil || selger || 'Ikke oppgitt';
  const kundeEpost = (isMegler && meglerEpost) ? meglerEpost : (selgerEpost || meglerEpost || '');
  const kundeRef = megler || kundeNavn;

  let html = '<div style="font-family:Arial,sans-serif; max-width:700px; color:#333;">';
  html += '<div style="background:#1a5c2a;color:white;padding:15px;border-radius:6px 6px 0 0;">';
  html += `<h2 style="margin:0;">Klar til fakturering: ${oppdragsnr}</h2></div>`;
  html += '<div style="padding:25px; border:1px solid #ddd; border-top:none; border-radius:0 0 6px 6px;">';

  if (kommentarRegnskap) {
    html += '<div style="background-color: #fff3cd; color: #856404; padding: 15px; border-left: 5px solid #ffeeba; margin-bottom: 25px; border-radius: 4px;">';
    html += '<strong style="display:block; margin-bottom:5px; font-size:16px;">Melding til regnskap:</strong>';
    html += kommentarRegnskap.replace(/\n/g, '<br>');
    html += '</div>';
  }

  // Kunde
  html += '<h3 style="color:#1a5c2a; border-bottom:1px solid #ddd; padding-bottom:5px; margin-top:0;">Kunde</h3>';
  html += '<table style="width:100%; border-collapse:collapse; line-height:1.8;">';
  html += `<tr><td style="width:30px; color:#666;">1.</td><td style="width:180px; font-weight:bold;">Navn/Firma:</td><td>${kundeNavn}</td></tr>`;
  html += `<tr><td style="color:#666;">2.</td><td style="font-weight:bold;">Telefon:</td><td>${selgerTlf}</td></tr>`;
  html += `<tr><td style="color:#666;">3.</td><td style="font-weight:bold;">E-post:</td><td>${kundeEpost}</td></tr>`;
  html += `<tr><td style="color:#666;">4.</td><td style="font-weight:bold;">Adresse:</td><td>${gate}</td></tr>`;
  html += `<tr><td style="color:#666;">5.</td><td style="font-weight:bold;">Postnr.:</td><td>${postnr}</td></tr>`;
  html += `<tr><td style="color:#666;">6.</td><td style="font-weight:bold;">Poststed:</td><td>${poststed}</td></tr>`;
  html += '</table>';

  // Faktura
  html += '<h3 style="color:#1a5c2a; border-bottom:1px solid #ddd; padding-bottom:5px; margin-top:30px;">Faktura</h3>';
  html += '<table style="width:100%; border-collapse:collapse; line-height:1.8;">';
  html += `<tr><td style="width:30px; color:#666;">1.</td><td style="width:200px; font-weight:bold;">Kundens referanse:</td><td>${kundeRef}</td></tr>`;
  html += `<tr><td style="color:#666;">2.</td><td style="font-weight:bold;">PO-nr./Ordrenr.:</td><td>${fakturaRef || '-'}</td></tr>`;
  html += `<tr><td style="color:#666;">3.</td><td style="font-weight:bold;">Eiendom eier:</td><td>${selger}</td></tr>`;
  html += `<tr><td style="color:#666;">4.</td><td style="font-weight:bold;">Eiendom adresse:</td><td>${gate}</td></tr>`;
  html += `<tr><td style="color:#666;">5.</td><td style="font-weight:bold;">Postnr. + Poststed:</td><td>${postnr ? postnr + ' ' + poststed : poststed}</td></tr>`;
  html += '</table>';

  // Products
  const totalEks = prisEks;
  const inklMarked = rowData[COL.MED_MARKEDSVERDI - 1] === true || rowData[COL.MED_MARKEDSVERDI - 1] === 'TRUE';
  const antallTillegg = parseInt(rowData[COL.ANTALL_TILLEGGSBYGG - 1]) || 0;
  const boligtype = rowData[COL.BOLIGTYPE - 1] || '';
  const prodNrRaw = String(rowData[COL.PRODUKTNUMMER - 1] || '');

  const baseProd = prodNrRaw.split(',')[0].trim();
  const markedProd = prodNrRaw.includes(',') ? prodNrRaw.split(',')[1].trim() : '9';

  let markedEks = (inklMarked && boligtype !== 'Frittstående bygg') ? (2000 / 1.25) : 0;
  let tilleggEks = (antallTillegg * 1250) / 1.25;
  let baseEks = totalEks - markedEks - tilleggEks;
  if (baseEks < 0) { baseEks = totalEks; markedEks = 0; tilleggEks = 0; }

  html += '<h4 style="margin-top:25px; margin-bottom:10px; color:#666;">6. Produkter (Pris eks. mva)</h4>';
  html += '<table style="width:100%; border-collapse:collapse; border:1px solid #ddd; font-size:14px;">';
  html += '<tr style="background:#f1f1f1; color:#333; font-weight:bold; text-align:left;">';
  html += '<th style="padding:10px; border-bottom:1px solid #ddd;">Produktnr.</th>';
  html += '<th style="padding:10px; border-bottom:1px solid #ddd;">Produktnavn</th>';
  html += '<th style="padding:10px; border-bottom:1px solid #ddd; text-align:right;">Pris eks. mva</th></tr>';

  const addRow = (pNr, pNavn, pPris) => {
    html += `<tr style="background:#fff;">`;
    html += `<td style="padding:10px; border-bottom:1px solid #ddd;">${pNr}</td>`;
    html += `<td style="padding:10px; border-bottom:1px solid #ddd;">${pNavn}</td>`;
    html += `<td style="padding:10px; border-bottom:1px solid #ddd; text-align:right;">${formatCurrency(pPris)}</td></tr>`;
  };

  addRow(baseProd || '-', rapporttype + (boligtype ? ' - ' + boligtype : ''), baseEks);
  if (markedEks > 0) addRow(markedProd, 'Tilstandsrapport markedsverdi', markedEks);
  if (tilleggEks > 0) addRow('', `Tilleggsbygg (${antallTillegg} stk)`, tilleggEks);
  if (reiseEks > 0) addRow('', 'Reisekostnad (inkl. evt bom/ferge)', reiseEks);

  html += '<tr style="font-weight:bold; background:#e8f5e9; color:#1a5c2a;">';
  html += '<td colspan="2" style="padding:12px;">TOTAL EKS. MVA:</td>';
  html += `<td style="padding:12px; text-align:right;">${formatCurrency(totalEks + reiseEks)}</td></tr>`;
  html += '</table>';

  if (config.testMode) {
    html += '<div style="background:#fff3e0;padding:8px;border-radius:4px;margin-top:20px;">TESTMODUS</div>';
  }

  html += '</div></div>';
  return html;
}

function buildBefaringEmail(adresse, dato, tid, selgerNavn, oppdragstype) {
  const kundeNavn = selgerNavn || 'kunde';

  return '<div style="font-family:Arial,sans-serif;max-width:600px;">' +
    '<div style="background:#1a5c2a;color:white;padding:20px;border-radius:8px 8px 0 0;text-align:center;">' +
    '<h2 style="margin:0;">Bekreftelse på befaring</h2></div>' +
    '<div style="padding:24px;border:1px solid #ddd;border-top:none;">' +
    `<p>Hei ${kundeNavn},</p>` +
    '<p>Vi bekrefter herved avtalt befaring på følgende eiendom:</p>' +
    '<table style="width:100%;border-collapse:collapse;margin:16px 0;">' +
    `<tr style="background:#f5f5f5;"><td style="padding:10px;font-weight:bold;">Adresse:</td><td style="padding:10px;">${adresse}</td></tr>` +
    `<tr><td style="padding:10px;font-weight:bold;">Type oppdrag:</td><td style="padding:10px;">${oppdragstype}</td></tr>` +
    `<tr style="background:#f5f5f5;"><td style="padding:10px;font-weight:bold;">Dato:</td><td style="padding:10px;font-size:16px;font-weight:bold;">${dato}</td></tr>` +
    `<tr><td style="padding:10px;font-weight:bold;">Klokkeslett:</td><td style="padding:10px;font-size:16px;font-weight:bold;">${tid}</td></tr>` +
    '</table>' +

    '<div style="background:#f8f9fa;border:1px solid #e0e0e0;border-radius:6px;padding:16px;margin:16px 0;">' +
    '<h3 style="margin:0 0 12px;color:#1a5c2a;">Sjekkliste før befaring</h3>' +
    '<p style="font-weight:bold;margin:12px 0 4px;">Rydd fri adkomst:</p>' +
    '<ul style="margin:0;padding-left:20px;color:#333;">' +
    '<li>Alle rom, boder, garasje, teknisk rom</li>' +
    '<li>Klargjør loft, krypkjeller, takluke, legg frem stige ved behov</li>' +
    '<li>Sørg for tilgang til sikringsskap, hovedkran, vannstoppeventil, stoppekran i våtrom, varmtvannsbereder</li>' +
    '<li>Sørg for tilgang til sluk, rørskap, synlige rør, under servanter</li></ul>' +
    '<p style="font-weight:bold;margin:12px 0 4px;">Noter kjente forhold:</p>' +
    '<ul style="margin:0;padding-left:20px;color:#333;"><li>Fukt, lekkasjer, sprekker, setninger, lukt, støy, tidligere skader</li></ul>' +
    '<p style="font-weight:bold;margin:12px 0 4px;">Finn frem dokumentasjon:</p>' +
    '<ul style="margin:0;padding-left:20px;color:#333;"><li>Dokumentasjon som er lagret digitalt kan sendes på epost</li></ul>' +
    '<p style="font-weight:bold;margin:12px 0 4px;">Oppgraderinger / oppussing:</p>' +
    '<ul style="margin:0;padding-left:20px;color:#333;"><li>Skriv en punktliste med årstall over hva som er gjort og når</li></ul>' +
    '<p style="font-weight:bold;margin:12px 0 4px;">Hulltaking:</p>' +
    '<ul style="margin:0;padding-left:20px;color:#333;">' +
    '<li>Iht. Avhendingslova (NS 3600) skal det utføres hulltaking i alle våtrom og i rom under terreng</li>' +
    '<li>Etter hulltaking monteres plastlokk</li></ul>' +
    '<div style="background:#fff3e0;border:1px solid #ffcc80;border-radius:4px;padding:12px;margin-top:12px;">' +
    '<p style="margin:0;font-weight:bold;">Boligmappa</p>' +
    '<p style="margin:4px 0 0;">Logg inn på boligmappa.no → Velg riktig bolig → "Gi tilgang" → Legg inn vår e-post → Lesetilgang 1 måned</p></div>' +
    '</div>' +

    '<p style="margin-top:16px;color:#666;">Har du spørsmål? Ta gjerne kontakt.</p></div>' +
    '<div style="background:#f5f5f5;padding:16px;border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;font-size:13px;color:#666;">' +
    '<strong>Jacob Engholm Holen</strong><br>Takstingeniør<br>+47 469 49 615<br>jacob@naava.no<br>www.naava.no<br><br>' +
    '<em>Medlem av Norsk Takst og NITO</em></div></div>';
}

function buildNewOppdragEmail(parsed, dato, folderUrl, oppdragsnr, avstandKm, reiseEks, reiseInkl) {
  const fields = [
    ['Oppdragsnr', oppdragsnr], ['Type', parsed.oppdragstype], ['Kilde', parsed.kilde],
    ['Adresse', parsed.adresse], ['Oppdragsgiver', parsed.oppdragsgiver],
    ['Selger', parsed.selger], ['Selger tlf', parsed.selgerTlf],
    ['Megler', parsed.megler], ['Faktura ref', parsed.fakturaRef],
    ['Faktura til', parsed.fakturaSendesTil], ['Mottatt', dato],
  ];
  if (avstandKm) {
    fields.push(['Avstand t/r', avstandKm + ' km']);
    fields.push(['Reise eks mva', formatCurrency(reiseEks)]);
    fields.push(['Reise inkl mva', formatCurrency(reiseInkl)]);
  }

  let detailRows = '';
  fields.forEach((f, i) => {
    if (f[1]) {
      detailRows += `<tr style="background:${i % 2 === 0 ? '#fff' : '#f9f9f9'}"><td style="padding:8px;font-weight:bold;">${f[0]}</td><td style="padding:8px;">${f[1]}</td></tr>`;
    }
  });

  return '<div style="font-family:Arial,sans-serif;max-width:600px;">' +
    '<div style="background:#1a5c2a;color:white;padding:20px;border-radius:8px 8px 0 0;">' +
    `<h2 style="margin:0;">${parsed.oppdragstype}</h2></div>` +
    '<div style="padding:20px;border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;">' +
    (config.testMode ? '<div style="background:#fff3e0;padding:8px;border-radius:4px;margin-bottom:12px;">TESTMODUS</div>' : '') +
    `<table style="width:100%;border-collapse:collapse;">${detailRows}</table>` +
    `<p style="margin-top:16px;"><a href="${folderUrl}" style="background:#1a5c2a;color:white;padding:10px 20px;text-decoration:none;border-radius:4px;">Åpne mappe</a></p></div></div>`;
}

function buildWeeklyReportHtml(fakturerbare, fakturerte, ventende, totals) {
  const { getWeekNumber } = require('./utils');
  const ukeNr = getWeekNumber(new Date());
  let rows = '';
  fakturerbare.forEach(o => {
    rows += `<tr><td style="padding:5px;border-bottom:1px solid #eee;">${o.oppdragsnr}</td>` +
      `<td style="padding:5px;">${o.adresse}</td><td style="padding:5px;">${o.oppdragstype}</td>` +
      `<td style="padding:5px;">${o.fakturaRef || '-'}</td>` +
      `<td style="padding:5px;text-align:right;">${formatCurrency(o.prisEks)}</td>` +
      `<td style="padding:5px;text-align:right;">${formatCurrency(o.mvaBeløp)}</td>` +
      `<td style="padding:5px;text-align:right;font-weight:bold;">${formatCurrency(o.prisInkl)}</td>` +
      `<td style="padding:5px;text-align:right;">${formatCurrency(o.reiseInkl)}</td></tr>`;
  });

  return '<div style="font-family:Arial,sans-serif;max-width:900px;">' +
    '<div style="background:#1a5c2a;color:white;padding:20px;border-radius:8px 8px 0 0;">' +
    `<h2 style="margin:0;">Ukerapport uke ${ukeNr}</h2></div>` +
    '<div style="padding:20px;border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;">' +
    (config.testMode ? '<div style="background:#fff3e0;padding:8px;border-radius:4px;margin-bottom:12px;">TESTMODUS</div>' : '') +
    `<h3>Fakturerbart (${fakturerbare.length})</h3>` +
    (fakturerbare.length > 0 ?
      '<table style="width:100%;border-collapse:collapse;font-size:11px;"><tr style="background:#f5f5f5;">' +
      '<th style="padding:5px;text-align:left;">Nr</th><th>Adresse</th><th>Type</th><th>Ref</th>' +
      '<th style="text-align:right;">Eks mva</th><th style="text-align:right;">MVA</th>' +
      '<th style="text-align:right;">Inkl mva</th><th style="text-align:right;">Reise inkl</th></tr>' + rows +
      '<tr style="background:#1a5c2a;color:white;font-weight:bold;"><td colspan="4">TOTAL</td>' +
      `<td colspan="4" style="text-align:right;font-size:14px;">${formatCurrency(totals.totalInklMva + totals.totalReiseInkl)}</td></tr></table>` :
      '<p style="color:#666;">Ingen denne uken.</p>') +
    '</div></div>';
}

module.exports = {
  buildFakturaEmail,
  buildBefaringEmail,
  buildNewOppdragEmail,
  buildWeeklyReportHtml,
};
