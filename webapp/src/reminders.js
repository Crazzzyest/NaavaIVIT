const config = require('./config');
const { COL } = require('./columns');
const google = require('./google');
const { parseDateString, formatCurrency, getWeekNumber } = require('./utils');
const { buildWeeklyReportHtml } = require('./emails');

async function checkReminders() {
  console.log('Checking reminders...');

  const data = await google.getSheetData(config.sheet.name);
  if (!data || data.length < 2) return { alerts: 0 };

  const now = new Date();
  let alerts = 0;

  for (let i = 1; i < data.length; i++) {
    const status = data[i][COL.STATUS - 1];
    const timestamp = data[i][COL.TIMESTAMP - 1];
    const adresse = data[i][COL.ADRESSE - 1];
    const oppdragsnr = data[i][COL.OPPDRAGSNR - 1];

    if (status !== 'Mottatt' || !timestamp) continue;

    const mottattDato = new Date(timestamp);
    const timerSiden = (now - mottattDato) / (1000 * 60 * 60);

    if (timerSiden >= config.urgentHours && timerSiden < config.urgentHours + 1.5) {
      console.log(`URGENT: ${adresse} (${oppdragsnr}) - ${Math.round(timerSiden)} timer`);
      await google.sendEmail(
        config.email.ownerEmail,
        `HASTER: ${adresse} (${oppdragsnr}) - ${Math.round(timerSiden)} timer!`,
        `<p>Oppdrag <strong>${oppdragsnr}</strong> (${adresse}) har status "Mottatt" i ${Math.round(timerSiden)} timer.</p>`,
      );
      alerts++;
    }
  }

  console.log(`Reminders done. ${alerts} alerts sent.`);
  return { alerts };
}

async function sendWeeklyReport() {
  console.log('Sending weekly report...');

  const data = await google.getSheetData(config.sheet.name);
  if (!data || data.length < 2) return;

  const now = new Date();
  const monday = new Date(now);
  monday.setDate(now.getDate() - (now.getDay() === 0 ? 6 : now.getDay() - 1));
  monday.setHours(0, 0, 0, 0);

  const fakturerbare = [];
  const fakturerte = [];
  const ventende = [];
  let totalInklMva = 0, totalEksMva = 0, totalMva = 0, totalReiseEks = 0, totalReiseInkl = 0;

  for (let i = 1; i < data.length; i++) {
    const status = data[i][COL.STATUS - 1];
    const prisInkl = Number(data[i][COL.PRIS_INKL - 1] || 0);
    const prisEks = Number(data[i][COL.PRIS_EKS - 1] || 0);
    const mvaBeløp = Number(data[i][COL.MVA_BELOP - 1] || 0);
    const reiseEks = Number(data[i][COL.REISE_EKS - 1] || 0);
    const reiseInkl = Number(data[i][COL.REISE_INKL - 1] || 0);
    const statusDatoStr = data[i][COL.DATO_STATUSENDRING - 1];

    const sDato = parseDateString(statusDatoStr);
    if (!sDato) continue;

    const item = {
      oppdragsnr: data[i][COL.OPPDRAGSNR - 1],
      adresse: data[i][COL.ADRESSE - 1],
      oppdragstype: data[i][COL.OPPDRAGSTYPE - 1],
      prisInkl, prisEks, mvaBeløp, reiseEks, reiseInkl,
      fakturaRef: data[i][COL.FAKTURA_REF - 1],
      fakturaTil: data[i][COL.FAKTURA_SENDES_TIL - 1],
      statusDato: statusDatoStr,
    };

    if (status === 'Kan faktureres' && sDato >= monday) {
      fakturerbare.push(item);
      totalInklMva += prisInkl; totalEksMva += prisEks; totalMva += mvaBeløp;
      totalReiseEks += reiseEks; totalReiseInkl += reiseInkl;
    }
    if (status === 'Fakturert' && sDato >= monday) fakturerte.push(item);
    if (status === 'Kan faktureres') ventende.push(item);
  }

  const html = buildWeeklyReportHtml(fakturerbare, fakturerte, ventende, {
    totalInklMva, totalEksMva, totalMva, totalReiseEks, totalReiseInkl,
  });

  await google.sendEmail(
    `${config.email.ownerEmail},${config.email.accountantEmail}`,
    `Naava Takst - Ukerapport uke ${getWeekNumber(now)}`,
    html,
  );

  console.log('Weekly report sent');
}

module.exports = { checkReminders, sendWeeklyReport };
