const { google } = require('googleapis');
const XLSX = require('xlsx');
const crypto = require('crypto');

const LABEL_NAME = 'SheetUpdater/Processed';
const SCHEMA_FIELDS = ['country', 'mcc', 'mnc', 'rate', 'bind', 'operator', 'date'];
const HEADER_ALIASES = {
  country: ['country', 'destination', 'dest'],
  mcc: ['mcc'],
  mnc: ['mnc'],
  rate: ['rate', 'price'],
  bind: ['bind'],
  operator: ['operator', 'carrier', 'vendor'],
  date: ['update date', 'date', 'updated', 'effective date'],
};

function mapHeaderRow(headerRow) {
  const map = {};
  (headerRow || []).forEach((h, i) => {
    const norm = (h || '').toString().trim().toLowerCase();
    for (const field of SCHEMA_FIELDS) {
      if (HEADER_ALIASES[field].includes(norm)) map[field] = i;
    }
  });
  return map;
}

function parseAttachmentBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const rows = [];
  wb.SheetNames.forEach((name) => {
    const data = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false });
    if (data.length < 2) return;
    const colMap = mapHeaderRow(data[0]);
    if (colMap.country === undefined || colMap.rate === undefined) return; // doesn't look like a rate sheet
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (!row || !row[colMap.country]) continue;
      const get = (k) => (colMap[k] !== undefined ? row[colMap[k]] || '' : '');
      rows.push({
        country: get('country'), mcc: get('mcc'), mnc: get('mnc'), rate: get('rate'),
        bind: get('bind'), operator: get('operator'),
        date: get('date') || new Date().toISOString().slice(0, 10),
      });
    }
  });
  return rows;
}

async function parseBodyWithAI(text) {
  if (!process.env.ANTHROPIC_API_KEY || !text.trim()) return [];
  const prompt = `Extract any telecom rate table rows from this email into JSON. Schema: [{"country":"","mcc":"","mnc":"","rate":"","bind":"","operator":"","date":""}]. Only include rows you are confident represent an actual rate update. If none, return []. Return ONLY the JSON array, nothing else.\n\nEmail content:\n${text.slice(0, 12000)}`;
  const r = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-api-key': process.env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
    body: JSON.stringify({ model: 'claude-sonnet-4-6', max_tokens: 2000, messages: [{ role: 'user', content: prompt }] }),
  });
  const data = await r.json();
  const textOut = ((data.content) || []).map((b) => b.text || '').join('');
  try { return JSON.parse(textOut.trim().replace(/^```json|```$/g, '')); } catch (e) { return []; }
}

function decodeBase64Url(str) { return Buffer.from(str.replace(/-/g, '+').replace(/_/g, '/'), 'base64'); }

function extractBodyText(payload) {
  let text = '';
  (function walk(part) {
    if (!part) return;
    if (part.mimeType === 'text/plain' && part.body && part.body.data) text += decodeBase64Url(part.body.data).toString('utf8') + '\n';
    if (part.parts) part.parts.forEach(walk);
  })(payload);
  return text;
}

function collectAttachments(payload) {
  const atts = [];
  (function walk(part) {
    if (!part) return;
    if (part.filename && /\.(xlsx|xls|csv)$/i.test(part.filename) && part.body) {
      atts.push({ filename: part.filename, attachmentId: part.body.attachmentId });
    }
    if (part.parts) part.parts.forEach(walk);
  })(payload);
  return atts;
}

async function ensureLabel(gmail) {
  const list = await gmail.users.labels.list({ userId: 'me' });
  const existing = (list.data.labels || []).find((l) => l.name === LABEL_NAME);
  if (existing) return existing.id;
  const created = await gmail.users.labels.create({ userId: 'me', requestBody: { name: LABEL_NAME, labelListVisibility: 'labelHide', messageListVisibility: 'hide' } });
  return created.data.id;
}

function deriveVendorName(domainOrSender) {
  // "cmi.chinamobile.com" -> "chinamobile" -> "Chinamobile"; "vertexsms.com" -> "vertexsms" -> "Vertexsms"
  const parts = domainOrSender.split('.');
  const core = parts.length >= 2 ? parts[parts.length - 2] : parts[0];
  return core.charAt(0).toUpperCase() + core.slice(1);
}

function parseSenderMap() {
  // GMAIL_SENDER_MAP="vertexsms.com,naliaglobal.net,vendor3.com=Custom Tab Name"
  // A bare domain auto-derives its tab name; add "=Name" only to override it.
  const map = {};
  (process.env.GMAIL_SENDER_MAP || '').split(',').map((s) => s.trim()).filter(Boolean).forEach((pair) => {
    const [sender, name] = pair.split('=').map((x) => x && x.trim());
    if (!sender) return;
    map[sender.toLowerCase()] = name || deriveVendorName(sender.toLowerCase());
  });
  return map;
}

function vendorForSender(fromHeader, senderMap) {
  const from = (fromHeader || '').toLowerCase();
  for (const sender in senderMap) if (from.includes(sender)) return senderMap[sender];
  return null;
}

async function scanInbox(oauth2Client) {
  const gmail = google.gmail({ version: 'v1', auth: oauth2Client });
  const senderMap = parseSenderMap();
  const senders = Object.keys(senderMap);
  if (senders.length === 0) return [];

  const labelId = await ensureLabel(gmail);
  const q = `(${senders.map((s) => `from:${s}`).join(' OR ')}) newer_than:3d -label:"${LABEL_NAME}"`;
  const list = await gmail.users.messages.list({ userId: 'me', q, maxResults: 50 });
  const messages = list.data.messages || [];
  const candidates = [];

  for (const m of messages) {
    const full = await gmail.users.messages.get({ userId: 'me', id: m.id, format: 'full' });
    const headers = full.data.payload.headers || [];
    const fromHeader = (headers.find((h) => h.name === 'From') || {}).value || '';
    const subject = (headers.find((h) => h.name === 'Subject') || {}).value || '';
    const vendor = vendorForSender(fromHeader, senderMap);
    if (!vendor) { continue; }

    let rows = [];
    for (const att of collectAttachments(full.data.payload)) {
      const attData = await gmail.users.messages.attachments.get({ userId: 'me', messageId: m.id, id: att.attachmentId });
      rows = rows.concat(parseAttachmentBuffer(decodeBase64Url(attData.data.data)));
    }
    if (rows.length === 0) rows = rows.concat(await parseBodyWithAI(extractBodyText(full.data.payload)));

    rows.forEach((r) => candidates.push({
      id: crypto.randomUUID(), ...r, vendor, sourceSubject: subject, sourceFrom: fromHeader,
      detectedAt: new Date().toISOString(),
    }));

    await gmail.users.messages.modify({ userId: 'me', id: m.id, requestBody: { addLabelIds: [labelId, 'UNREAD'] } });
  }
  return candidates;
}

module.exports = { scanInbox, parseAttachmentBuffer };
