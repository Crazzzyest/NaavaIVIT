/**
 * One-time OAuth2 setup script.
 *
 * 1. Go to Google Cloud Console → APIs & Services → Credentials
 * 2. Create an OAuth 2.0 Client ID (type: Web application)
 * 3. Add http://localhost:3333/callback as a redirect URI
 * 4. Enable Gmail API, Google Sheets API, Google Drive API
 * 5. Set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET in .env
 * 6. Run: node auth-setup.js
 * 7. Open the URL in your browser, sign in, and authorize
 * 8. Copy the refresh token into .env as GOOGLE_REFRESH_TOKEN
 */

const http = require('http');
const { URL } = require('url');
const { google } = require('googleapis');
const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '.env') });

const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REDIRECT_URI = 'http://localhost:3333/callback';

if (!CLIENT_ID || !CLIENT_SECRET) {
  console.error('Set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET in .env first.');
  process.exit(1);
}

const oauth2Client = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);

const SCOPES = [
  'https://www.googleapis.com/auth/gmail.modify',
  'https://www.googleapis.com/auth/gmail.send',
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive',
];

const authUrl = oauth2Client.generateAuthUrl({
  access_type: 'offline',
  prompt: 'consent',
  scope: SCOPES,
});

console.log('\nOpen this URL in your browser:\n');
console.log(authUrl);
console.log('\nWaiting for callback on http://localhost:3333/callback ...\n');

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost:3333');
  if (url.pathname !== '/callback') {
    res.writeHead(404);
    res.end('Not found');
    return;
  }

  const code = url.searchParams.get('code');
  if (!code) {
    res.writeHead(400);
    res.end('Missing code parameter');
    return;
  }

  try {
    const { tokens } = await oauth2Client.getToken(code);
    console.log('\n=== SUCCESS ===\n');
    console.log('Add this to your .env file:\n');
    console.log(`GOOGLE_REFRESH_TOKEN=${tokens.refresh_token}\n`);

    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end('<h1>Success!</h1><p>Refresh token printed in terminal. You can close this tab.</p>');

    setTimeout(() => process.exit(0), 1000);
  } catch (e) {
    console.error('Error exchanging code:', e.message);
    res.writeHead(500);
    res.end('Error: ' + e.message);
  }
});

server.listen(3333);
