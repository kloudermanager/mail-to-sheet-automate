require('dotenv').config();
const express = require('express');
const path = require('path');
const cron = require('node-cron');
const { google } = require('googleapis');
const { scanInbox, parseAttachmentBuffer } = require('./gmail-sync');

const app = express();
const PORT = process.env.PORT || 3000;
const PENDING_TAB = 'Pending Review';
const AUTO_APPLY = process.env.AUTO_APPLY_EMAIL_UPDATES === 'true';

// ---------- Basic Auth (shared team login) ----------
function checkAuth(req, res, next) {
  const auth = req.headers.authorization;
  if (!auth || !auth.startsWith('Basic ')) {
    res.set('WWW-Authenticate', 'Basic realm="Sheet Updater"');
    return res.status(401).send('Authentication required');
  }
  const [user, pass] = Buffer.from(auth.split(' ')[1], 'base64').toString().split(':');
  if (user === process.env.ADMIN_USER && pass === process.env.ADMIN_PASS) return next();
  res.set('WWW-Authenticate', 'Basic realm="Sheet Updater"');
  return res.status(401).send('Invalid credentials');
}
app.use(checkAuth);
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json({ limit: '25mb' }));

// ---------- Shared Google auth (Sheets write + Gmail) ----------
function getOAuthClient() {
  const auth = new google.auth.OAuth2(process.env.GOOGLE_CLIENT_ID, process.env.GOOGLE_CLIENT_SECRET);
  auth.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
  return auth;
}
function getSheets() { return google.sheets({ version: 'v4', auth: getOAuthClient() }); }

// ---------- Pending Review tab helpers ----------
async function ensurePendingTab(sheets, sheetId) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId, fields: 'sheets.properties' });
  let tab = meta.data.sheets.find(s => s.properties.title === PENDING_TAB);
  if (tab) return tab.properties.sheetId;
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: sheetId,
    requestBody: { requests: [{ addSheet: { properties: { title: PENDING_TAB } } }] },
  });
  const newTabId = res.data.replies[0].addSheet.properties.sheetId;
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId, range: `'${PENDING_TAB}'!A1:K1`, valueInputOption: 'RAW',
    requestBody: { values: [['ID', 'Country', 'MCC', 'MNC', 'Rate', 'Bind', 'Operator', 'Date', 'Vendor', 'Source', 'Detected At']] },
  });
  return newTabId;
}

async function appendCandidates(candidates) {
  if (candidates.length === 0) return;
  const sheets = getSheets();
  const sheetId = process.env.SHEET_ID;
  await ensurePendingTab(sheets, sheetId);
  const values = candidates.map(c => [c.id, c.country, c.mcc, c.mnc, c.rate, c.bind, c.operator, c.date, c.vendor, `${c.sourceSubject} <${c.sourceFrom}>`, c.detectedAt]);
  await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId, range: `'${PENDING_TAB}'!A:K`, valueInputOption: 'RAW', requestBody: { values },
  });
}

async function ensureVendorTab(sheets, sheetId, vendorName) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId, fields: 'sheets.properties.title' });
  const existing = meta.data.sheets.find(s => s.properties.title.toLowerCase() === vendorName.toLowerCase());
  if (existing) return existing.properties.title; // reuse exact existing title/casing
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: sheetId,
    requestBody: { requests: [{ addSheet: { properties: { title: vendorName } } }] },
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId, range: `'${vendorName}'!A1:G1`, valueInputOption: 'RAW',
    requestBody: { values: [['COUNTRY', 'MCC', 'MNC', 'RATE', 'BIND', 'Operator', 'UPDATE DATE']] },
  });
  return vendorName;
}

async function applyCandidatesDirectly(candidates) {
  const sheets = getSheets();
  const sheetId = process.env.SHEET_ID;
  const byVendor = {};
  candidates.forEach(c => { (byVendor[c.vendor] = byVendor[c.vendor] || []).push(c); });
  for (const vendor in byVendor) {
    const actualTab = await ensureVendorTab(sheets, sheetId, vendor);
    const values = byVendor[vendor].map(c => [c.country, c.mcc, c.mnc, c.rate, c.bind, c.operator, c.date]);
    await sheets.spreadsheets.values.append({
      spreadsheetId: sheetId, range: `'${actualTab}'!A:G`, valueInputOption: 'RAW', requestBody: { values },
    });
  }
}

async function getPendingRows() {
  const sheets = getSheets();
  const sheetId = process.env.SHEET_ID;
  await ensurePendingTab(sheets, sheetId);
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: sheetId, range: `'${PENDING_TAB}'!A2:K` });
  return (res.data.values || []).map(r => ({
    id: r[0], country: r[1], mcc: r[2], mnc: r[3], rate: r[4], bind: r[5],
    operator: r[6], date: r[7], vendor: r[8], source: r[9], detectedAt: r[10],
  }));
}

async function removePendingRowById(id) {
  const sheets = getSheets();
  const sheetId = process.env.SHEET_ID;
  const tabId = await ensurePendingTab(sheets, sheetId);
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: sheetId, range: `'${PENDING_TAB}'!A2:A` });
  const ids = (res.data.values || []).map(r => r[0]);
  const idx = ids.indexOf(id);
  if (idx === -1) return null;
  const rowNumber = idx + 2; // +1 for header, +1 for 0-index
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: sheetId,
    requestBody: { requests: [{ deleteDimension: { range: { sheetId: tabId, dimension: 'ROWS', startIndex: rowNumber - 1, endIndex: rowNumber } } }] },
  });
  return true;
}

// ---------- Daily email scan ----------
async function runEmailScan() {
  const candidates = await scanInbox(getOAuthClient());
  if (candidates.length === 0) { console.log('[email-scan] no new rate emails'); return; }
  if (AUTO_APPLY) {
    await applyCandidatesDirectly(candidates);
    console.log(`[email-scan] auto-applied ${candidates.length} rows`);
  } else {
    await appendCandidates(candidates);
    console.log(`[email-scan] staged ${candidates.length} rows for review`);
  }
}

// ---------- API ----------
app.get('/api/pending', async (req, res) => {
  try { res.json({ ok: true, rows: await getPendingRows(), autoApply: AUTO_APPLY }); }
  catch (e) { res.status(500).json({ ok: false, error: e.message }); }
});
app.post('/api/pending/:id/approve', async (req, res) => {
  try {
    const rows = await getPendingRows();
    const row = rows.find(r => r.id === req.params.id);
    if (!row) return res.status(404).json({ ok: false, error: 'Not found' });
    await applyCandidatesDirectly([row]);
    await removePendingRowById(req.params.id);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ ok: false, error: e.message }); }
});
app.post('/api/pending/:id/reject', async (req, res) => {
  try { await removePendingRowById(req.params.id); res.json({ ok: true }); }
  catch (e) { res.status(500).json({ ok: false, error: e.message }); }
});
app.post('/api/scan-now', async (req, res) => {
  try { await runEmailScan(); res.json({ ok: true }); }
  catch (e) { console.error('[scan-now] failed:', e); res.status(500).json({ ok: false, error: e.message }); }
});

// ---------- Manual import (CSV / XLSX upload) ----------
app.get('/api/tabs', async (req, res) => {
  try {
    const meta = await getSheets().spreadsheets.get({ spreadsheetId: process.env.SHEET_ID, fields: 'sheets.properties.title' });
    res.json({ ok: true, tabs: meta.data.sheets.map(s => s.properties.title).filter(t => t !== PENDING_TAB) });
  } catch (e) { res.status(500).json({ ok: false, error: e.message }); }
});

app.post('/api/import', async (req, res) => {
  try {
    const { vendor, filename, data } = req.body || {};
    const name = (vendor || '').trim();
    if (!name || /[\[\]:*?\/\\]/.test(name) || name.length > 100) {
      return res.status(400).json({ ok: false, error: 'Enter a valid company / tab name (no [ ] : * ? / \\).' });
    }
    if (!data) return res.status(400).json({ ok: false, error: 'No file received.' });
    const rows = parseAttachmentBuffer(Buffer.from(data, 'base64'));
    if (rows.length === 0) {
      return res.status(400).json({ ok: false, error: 'No rate rows found. The first row must be a header containing at least Country and Rate columns.' });
    }
    await applyCandidatesDirectly(rows.map(r => ({ ...r, vendor: name })));
    console.log(`[import] ${rows.length} rows from ${filename} into "${name}"`);
    res.json({ ok: true, count: rows.length });
  } catch (e) { console.error(e); res.status(500).json({ ok: false, error: e.message }); }
});

app.post('/api/add-item', async (req, res) => {
  try {
    const { vendor, country, mcc, mnc, rate, bind, operator, date } = req.body || {};
    const name = (vendor || '').trim();
    if (!name || /[\[\]:*?\/\\]/.test(name) || name.length > 100) {
      return res.status(400).json({ ok: false, error: 'Enter a valid company / tab name (no [ ] : * ? / \\).' });
    }
    if (!country || !String(country).trim()) return res.status(400).json({ ok: false, error: 'Country is required.' });
    if (rate === undefined || rate === null || String(rate).trim() === '') return res.status(400).json({ ok: false, error: 'Rate is required.' });
    const row = {
      country: String(country).trim(), mcc: (mcc || '').toString().trim(), mnc: (mnc || '').toString().trim(),
      rate: String(rate).trim(), bind: (bind || '').toString().trim(), operator: (operator || '').toString().trim(),
      date: (date && String(date).trim()) || new Date().toISOString().slice(0, 10), vendor: name,
    };
    await applyCandidatesDirectly([row]);
    console.log(`[add-item] 1 row into "${name}" (${row.country})`);
    res.json({ ok: true });
  } catch (e) { console.error(e); res.status(500).json({ ok: false, error: e.message }); }
});

// Inbox scan every morning at 5:30 AM Bangladesh time (Asia/Dhaka, UTC+6).
cron.schedule('30 5 * * *', () => { runEmailScan().catch(console.error); }, { timezone: 'Asia/Dhaka' });

app.listen(PORT, () => console.log(`Sheet Updater listening on port ${PORT}`));
