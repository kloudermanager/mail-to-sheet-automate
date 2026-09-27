// Run this ONCE on your own computer: node get-refresh-token.js
// It opens a login URL, you sign in with the Google account that can see
// your sheet and inbox, and it prints a refresh token to paste into your
// host's environment variables. You never need to run this again unless
// you revoke access.
require('dotenv').config();
const http = require('http');
const { google } = require('googleapis');

const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const PORT = 3001;
const REDIRECT_URI = `http://localhost:${PORT}/oauth2callback`;

if (!CLIENT_ID || !CLIENT_SECRET) {
  console.error('Set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET in a local .env file first.');
  process.exit(1);
}

const oauth2Client = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);
const authUrl = oauth2Client.generateAuthUrl({
  access_type: 'offline',
  prompt: 'consent', // forces Google to issue a refresh token every time
  scope: [
    'https://www.googleapis.com/auth/spreadsheets',  // read + write, so it can append rows
    'https://www.googleapis.com/auth/gmail.modify',  // read inbox + apply a "processed" label
  ],
});

const server = http.createServer(async (req, res) => {
  if (!req.url.startsWith('/oauth2callback')) return res.end('Waiting for auth...');
  const code = new URL(req.url, REDIRECT_URI).searchParams.get('code');
  res.end('Success! You can close this tab and go back to your terminal.');
  server.close();
  const { tokens } = await oauth2Client.getToken(code);
  console.log('\nAdd this to your host\'s environment variables:\n');
  console.log('GOOGLE_REFRESH_TOKEN=' + tokens.refresh_token);
  if (!tokens.refresh_token) {
    console.log('\nNo refresh token returned. If you\'ve authorized this app before, revoke access at https://myaccount.google.com/permissions and run this script again.');
  }
  process.exit(0);
});
server.listen(PORT, () => {
  console.log('Open this URL, sign in, and grant access:\n');
  console.log(authUrl);
});
