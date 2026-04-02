const { google } = require('googleapis');
const config = require('./config');

let authClient = null;

function getAuth() {
  if (authClient) return authClient;
  authClient = new google.auth.OAuth2(
    config.google.clientId,
    config.google.clientSecret,
  );
  authClient.setCredentials({ refresh_token: config.google.refreshToken });
  return authClient;
}

function getSheets() {
  return google.sheets({ version: 'v4', auth: getAuth() });
}

function getDrive() {
  return google.drive({ version: 'v3', auth: getAuth() });
}

function getGmail() {
  return google.gmail({ version: 'v1', auth: getAuth() });
}

// ============================================================
// SHEETS HELPERS
// ============================================================

async function getSheetData(sheetName) {
  const res = await getSheets().spreadsheets.values.get({
    spreadsheetId: config.sheet.id,
    range: `${sheetName}!A:AN`,
  });
  return res.data.values || [];
}

async function appendRow(sheetName, row) {
  await getSheets().spreadsheets.values.append({
    spreadsheetId: config.sheet.id,
    range: `${sheetName}!A:AN`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [row] },
  });
}

async function updateCell(sheetName, row, col, value) {
  const colLetter = colToLetter(col);
  const range = `${sheetName}!${colLetter}${row}`;
  await getSheets().spreadsheets.values.update({
    spreadsheetId: config.sheet.id,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[value]] },
  });
}

async function updateCells(sheetName, row, updates) {
  const data = updates.map(({ col, value }) => ({
    range: `${sheetName}!${colToLetter(col)}${row}`,
    values: [[value]],
  }));
  await getSheets().spreadsheets.values.batchUpdate({
    spreadsheetId: config.sheet.id,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data,
    },
  });
}

async function getRowCount(sheetName) {
  const data = await getSheetData(sheetName);
  return data.length;
}

async function clearRange(sheetName, startRow, endRow, numCols) {
  const range = `${sheetName}!A${startRow}:${colToLetter(numCols)}${endRow}`;
  await getSheets().spreadsheets.values.clear({
    spreadsheetId: config.sheet.id,
    range,
  });
}

async function writeRange(sheetName, startRow, startCol, values) {
  const startLetter = colToLetter(startCol);
  const endLetter = colToLetter(startCol + (values[0] ? values[0].length - 1 : 0));
  const range = `${sheetName}!${startLetter}${startRow}:${endLetter}${startRow + values.length - 1}`;
  await getSheets().spreadsheets.values.update({
    spreadsheetId: config.sheet.id,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });
}

function colToLetter(col) {
  let letter = '';
  let c = col;
  while (c > 0) {
    c--;
    letter = String.fromCharCode(65 + (c % 26)) + letter;
    c = Math.floor(c / 26);
  }
  return letter;
}

// ============================================================
// DRIVE HELPERS
// ============================================================

async function createFolder(name, parentId) {
  const drive = getDrive();
  const res = await drive.files.create({
    requestBody: {
      name,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentId],
    },
    fields: 'id, webViewLink',
  });
  return { id: res.data.id, url: res.data.webViewLink };
}

async function createOppdragFolder(folderName) {
  const parent = config.drive.rootFolderId;
  const folder = await createFolder(folderName, parent);
  await Promise.all([
    createFolder('Bilder', folder.id),
    createFolder('Dokumenter fra megler', folder.id),
    createFolder('Rapport', folder.id),
  ]);
  return folder;
}

async function moveFolder(folderId, newParentId) {
  const drive = getDrive();
  const file = await drive.files.get({
    fileId: folderId,
    fields: 'parents',
  });
  const previousParents = (file.data.parents || []).join(',');
  await drive.files.update({
    fileId: folderId,
    addParents: newParentId,
    removeParents: previousParents,
  });
}

async function getOrCreateAvsluttedeFolder() {
  const drive = getDrive();
  const res = await drive.files.list({
    q: `'${config.drive.rootFolderId}' in parents and mimeType='application/vnd.google-apps.folder' and name='avsluttede oppdrag' and trashed=false`,
    fields: 'files(id)',
  });
  if (res.data.files && res.data.files.length > 0) {
    return res.data.files[0].id;
  }
  const folder = await createFolder('avsluttede oppdrag', config.drive.rootFolderId);
  return folder.id;
}

// ============================================================
// GMAIL HELPERS
// ============================================================

async function searchGmail(query, maxResults = 20) {
  const gmail = getGmail();
  const res = await gmail.users.threads.list({
    userId: 'me',
    q: query,
    maxResults,
  });
  return res.data.threads || [];
}

async function getThread(threadId) {
  const gmail = getGmail();
  const res = await gmail.users.threads.get({
    userId: 'me',
    id: threadId,
    format: 'full',
  });
  return res.data;
}

async function getOrCreateLabel(labelName) {
  const gmail = getGmail();
  const res = await gmail.users.labels.list({ userId: 'me' });
  const existing = (res.data.labels || []).find(l => l.name === labelName);
  if (existing) return existing.id;

  const created = await gmail.users.labels.create({
    userId: 'me',
    requestBody: { name: labelName, labelListVisibility: 'labelShow', messageListVisibility: 'show' },
  });
  return created.data.id;
}

async function addLabelToThread(threadId, labelId) {
  const gmail = getGmail();
  await gmail.users.threads.modify({
    userId: 'me',
    id: threadId,
    requestBody: { addLabelIds: [labelId] },
  });
}

async function markThreadRead(threadId) {
  const gmail = getGmail();
  await gmail.users.threads.modify({
    userId: 'me',
    id: threadId,
    requestBody: { removeLabelIds: ['UNREAD'] },
  });
}

async function sendEmail(to, subject, htmlBody) {
  const safeRecipient = getSafeRecipient(to, subject);

  const raw = buildRawEmail(safeRecipient.to, safeRecipient.subject, htmlBody);
  const gmail = getGmail();
  await gmail.users.messages.send({
    userId: 'me',
    requestBody: { raw },
  });
  console.log(`Email sent to: ${safeRecipient.to} | Subject: ${safeRecipient.subject}`);
}

function getSafeRecipient(to, subject) {
  if (!config.testMode) return { to, subject };

  const allowed = [
    config.email.ownerEmail.toLowerCase(),
    config.email.accountantEmail.toLowerCase(),
  ];
  const recipients = to.split(',').map(e => e.trim().toLowerCase());
  const safe = recipients.filter(e => allowed.includes(e));

  if (safe.length === 0) {
    return {
      to: config.email.ownerEmail,
      subject: `[TEST - Ville gått til: ${to}] ${subject}`,
    };
  }
  return { to: safe.join(','), subject };
}

function buildRawEmail(to, subject, htmlBody) {
  const lines = [
    `To: ${to}`,
    `Subject: =?utf-8?B?${Buffer.from(subject).toString('base64')}?=`,
    'MIME-Version: 1.0',
    'Content-Type: text/html; charset=utf-8',
    '',
    htmlBody,
  ];
  return Buffer.from(lines.join('\r\n')).toString('base64url');
}

function parseMessagePayload(message) {
  const headers = message.payload.headers || [];
  const getHeader = (name) => {
    const h = headers.find(h => h.name.toLowerCase() === name.toLowerCase());
    return h ? h.value : '';
  };

  let body = '';
  function extractBody(parts) {
    if (!parts) return;
    for (const part of parts) {
      if (part.mimeType === 'text/plain' && part.body && part.body.data) {
        body = Buffer.from(part.body.data, 'base64').toString('utf-8');
        return;
      }
      if (part.parts) extractBody(part.parts);
    }
  }

  if (message.payload.body && message.payload.body.data) {
    body = Buffer.from(message.payload.body.data, 'base64').toString('utf-8');
  } else {
    extractBody(message.payload.parts);
  }

  return {
    from: getHeader('From'),
    subject: getHeader('Subject'),
    date: new Date(parseInt(message.internalDate)),
    body,
  };
}

module.exports = {
  getSheetData,
  appendRow,
  updateCell,
  updateCells,
  getRowCount,
  clearRange,
  writeRange,
  createOppdragFolder,
  moveFolder,
  getOrCreateAvsluttedeFolder,
  searchGmail,
  getThread,
  getOrCreateLabel,
  addLabelToThread,
  markThreadRead,
  sendEmail,
  parseMessagePayload,
};
