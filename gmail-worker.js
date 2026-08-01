/**
 * gmail-worker.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Render.com-compatible Gmail report processor.
 * Reads credentials from environment variables (not local files).
 * Runs an Express HTTP server for health checks (keeps Render service alive).
 * Polls Gmail every 5 minutes for new "Raporti Ditor" report emails.
 *
 * Required environment variables on Render:
 *   GOOGLE_CREDENTIALS  — full contents of credentials.json (as a JSON string)
 *   GOOGLE_TOKEN        — full contents of token.json (as a JSON string)
 *   PORT                — set automatically by Render
 *
 * Optional:
 *   CHECK_INTERVAL_MIN  — how often to poll Gmail (default: 5 minutes)
 *   RENDER_EXTERNAL_URL — Render sets this automatically (used for self-ping)
 */

'use strict';

const { google }   = require('googleapis');
const XLSX         = require('xlsx');
const JSZip        = require('jszip');
const express      = require('express');
const https        = require('https');
const http         = require('http');
const { Readable } = require('stream');

// ── CONFIG ────────────────────────────────────────────────────────────────────
const XLSX_FILE_ID      = '1abLRrgklWeV3wx-KEmA0u4SgCH5ebw3s';
const TOTAL_ROOMS       = 110;
const CHECK_INTERVAL_MS = (parseInt(process.env.CHECK_INTERVAL_MIN) || 5) * 60 * 1000;
const PORT              = parseInt(process.env.PORT) || 3000;
// GAS web app URL (set in Render env vars as GAS_ENDPOINT_URL)
const GAS_ENDPOINT_URL  = process.env.GAS_ENDPOINT_URL || '';

const SCOPES = [
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/gmail.modify',
];

// Names to EXCLUDE from hotel occupancy (owner / out-of-service rooms)
const EXCLUDE_NAMES = [
  'ernest caci', 'olti caci', 'olti  caci', 'ahmet caci',
  'jasht pune', 'jashte pune',
  'bllok'
];

// Block/agency CLIENT names also excluded from occupancy for reports dated 2026-06-06
// onward (matched ONLY against the client name, never the source/Burimi).
const EXCLUDE_NAMES_NEW = ['itaka', 'saistours', 'w2m'];
const NEW_EXCL_FROM_DATE = '2026-06-06';

// Room names that count as House Use in F&B (matched against col[7] "Skonto Për")
// Dhoma VILA 1 - 13, Dhoma VILA2 - 14, Dhoma 313 - 64, Fature Qerasje
function isHouseUse(tableStr) {
  const s = String(tableStr).toLowerCase().replace(/\s+/g, ' ').trim();
  return (
    /dhoma\s*vila\s*1/.test(s) ||
    /dhoma\s*vila\s*2/.test(s) ||
    /dhoma\s*313/.test(s)      ||
    /fatur[eë]\s*qerasje/.test(s)
  );
}

// ── STATUS TRACKING (for /health endpoint) ────────────────────────────────────
// Bump BUILD on each deploy that matters so `GET /` can confirm what's actually live
// (Render's autodeploy has been known to lag/stick behind origin/master here).
const BUILD          = 'hotel-xlsx-endpoint-2026-08-01';
let lastCheckTime    = null;
let lastCheckStatus  = 'not started';
let lastProcessed    = null;
let totalProcessed   = 0;
let checkCount       = 0;

// ── EXPRESS HEALTH SERVER ─────────────────────────────────────────────────────
function startHealthServer() {
  const app = express();
  app.use(express.json({ limit: '256kb' }));

  app.get('/', (req, res) => {
    res.json({
      service:       'Flower Hotel Gmail Report Processor',
      status:        'running',
      build:         BUILD,
      checkInterval: `${CHECK_INTERVAL_MS / 60000} minutes`,
      lastCheck:     lastCheckTime,
      lastStatus:    lastCheckStatus,
      lastProcessed: lastProcessed,
      totalProcessed,
      checkCount,
      uptime:        Math.floor(process.uptime()) + 's',
    });
  });

  app.get('/health', (req, res) => {
    res.json({ status: 'ok', uptime: Math.floor(process.uptime()) + 's' });
  });

  // ── HOTEL DAILY PERFORMANCE push (from the FLOW dashboard's nightly
  //    KontrolloPrenotimet ingest). Writes the Sample Power BI xlsx directly,
  //    reusing the SAME tested patchXlsx()/patchHotelRow() the daily email
  //    pipeline uses — because Apps Script / the Sheets API cannot write an
  //    .xlsx. Body is the exact payload the dashboard already sends to
  //    KONTROLLO_PUSH_URLS: { action:'hotel_performance', date:'YYYY-MM-DD',
  //    nightsOccupied, nightsAvailable, revenue }. Upserts by date row; if the
  //    date row doesn't exist yet, hotelWritten=false (row must be pre-created,
  //    same rule as the email pipeline).
  app.post('/api/hotel-xlsx', async (req, res) => {
    try {
      const b = req.body || {};
      if (process.env.HOTEL_PUSH_TOKEN && b.token !== process.env.HOTEL_PUSH_TOKEN) {
        return res.status(403).json({ ok: false, error: 'bad token' });
      }
      const m = String(b.date || '').match(/\d{4}-\d{2}-\d{2}/);
      if (!m) return res.status(400).json({ ok: false, error: 'missing/invalid date (need YYYY-MM-DD)' });
      const isoDate = m[0];
      const hotelValues = {
        nightsOccupied:  Number(b.nightsOccupied),
        nightsAvailable: Number(b.nightsAvailable) || TOTAL_ROOMS,
        revenue:         Number(b.revenue) || 0,
      };
      if (!Number.isFinite(hotelValues.nightsOccupied)) {
        return res.status(400).json({ ok: false, error: 'missing/invalid nightsOccupied' });
      }
      const auth  = await authorize();
      const drive = google.drive({ version: 'v3', auth });
      const results = await patchXlsx(drive, [
        { isoDate, hotelValues, fnbValues: null, expValues: null,
          writeFnb: false, writeHotel: true, writeExp: false },
      ]);
      const r = results[0] || {};
      console.log(`[hotel-xlsx] ${isoDate} → ${hotelValues.nightsOccupied}/${hotelValues.nightsAvailable} rev ${hotelValues.revenue} · written=${!!r.hotelWritten}`);
      return res.json({ ok: !!r.hotelWritten, date: isoDate, hotelWritten: !!r.hotelWritten, hotelValues });
    } catch (e) {
      console.error('[hotel-xlsx] error:', e.message);
      return res.status(500).json({ ok: false, error: e.message });
    }
  });

  app.listen(PORT, () => {
    console.log(`[server] Health server listening on port ${PORT}`);
  });
}

// ── SELF-PING (prevents Render free tier spin-down) ───────────────────────────
function startSelfPing() {
  const baseUrl = process.env.RENDER_EXTERNAL_URL;
  if (!baseUrl) {
    console.log('[ping] RENDER_EXTERNAL_URL not set — self-ping disabled');
    return;
  }
  const pingUrl = baseUrl.replace(/\/$/, '') + '/health';
  console.log(`[ping] Self-ping every 10 min → ${pingUrl}`);

  setInterval(() => {
    const lib = pingUrl.startsWith('https') ? https : http;
    lib.get(pingUrl, (res) => {
      console.log(`[ping] ${new Date().toISOString()} → HTTP ${res.statusCode}`);
      res.resume();
    }).on('error', (e) => {
      console.log(`[ping] Error: ${e.message}`);
    });
  }, 10 * 60 * 1000);
}

// ── AUTH (from environment variables) ────────────────────────────────────────
async function authorize() {
  const credsJson  = process.env.GOOGLE_CREDENTIALS;
  const tokenJson  = process.env.GOOGLE_TOKEN;

  if (!credsJson) throw new Error('GOOGLE_CREDENTIALS environment variable is not set');
  if (!tokenJson) throw new Error('GOOGLE_TOKEN environment variable is not set');

  const creds = JSON.parse(credsJson);
  const { client_id, client_secret } = creds.installed || creds.web;

  // On Render we can't do the interactive OAuth flow, so we require a pre-existing token.
  // The refresh_token will be used automatically to get new access tokens.
  const auth = new google.auth.OAuth2(client_id, client_secret, 'http://localhost:3456');
  auth.setCredentials(JSON.parse(tokenJson));

  // Auto-refresh: update in-memory token when it expires
  auth.on('tokens', (tokens) => {
    console.log('[auth] Token refreshed');
    // We can't write to disk on Render, but the new tokens are held in memory for this session
    if (tokens.refresh_token) {
      // If a new refresh_token is issued, log it (admin can update env var)
      console.log('[auth] New refresh_token issued — update GOOGLE_TOKEN env var on Render!');
      console.log('[auth] New token:', JSON.stringify(tokens).slice(0, 80) + '...');
    }
  });

  return auth;
}

// ── GMAIL HELPERS ─────────────────────────────────────────────────────────────
async function findReportEmails(gmail) {
  // Match on BOTH the subject AND the attachment file names, so a send is picked up even
  // when the subject is unusual (e.g. a standalone expenses email titled just "Blerjet"):
  //  · subject variants the hotel uses: "Raporti Ditor", "Raport Ditor Data …",
  //    "Raport Data …" (no "i"), "shpenzime ditore", a "Blerjet …" send, "Prenotimet …".
  //  · filename: matches the attachment names — "Blerjet Data ….xls" (expenses),
  //    "Prenotimet ne recepsion.xls" (hotel) — regardless of what the subject says.
  // The content-based routing still decides what each attachment actually is; this only
  // makes sure the email is picked up promptly (not just via the catch-all fallback).
  // Gmail tokens are distinct, so list both Raport and Raporti.
  const primaryQuery = '(subject:Raporti OR subject:Raport OR subject:shpenzime OR subject:blerjet OR subject:Prenotimet OR filename:blerjet OR filename:shpenzime OR filename:prenotimet) has:attachment -label:processed-report in:inbox';
  const r1 = await gmail.users.messages.list({ userId: 'me', q: primaryQuery, maxResults: 20 });
  if ((r1.data.messages || []).length > 0) {
    console.log('  (matched subject: Raporti Ditor / shpenzime ditore)');
    return r1.data.messages.map(m => m.id);
  }
  const fallbackQuery = 'has:attachment -label:processed-report in:inbox';
  const r2 = await gmail.users.messages.list({ userId: 'me', q: fallbackQuery, maxResults: 20 });
  if ((r2.data.messages || []).length > 0) {
    console.log('  (fallback: any unprocessed attachment email)');
  }
  return (r2.data.messages || []).map(m => m.id);
}

async function getAttachments(gmail, messageId) {
  const msg = await gmail.users.messages.get({ userId: 'me', id: messageId });
  const attachments = [];

  const headers = Object.fromEntries(
    (msg.data.payload.headers || []).map(h => [String(h.name).toLowerCase(), h.value])
  );
  const subject = headers['subject'] || '';
  // Year the email was received — used to resolve filenames like "9 Qeshor" (no year).
  const year = msg.data.internalDate
    ? new Date(parseInt(msg.data.internalDate)).getUTCFullYear()
    : new Date().getUTCFullYear();

  function findParts(parts) {
    if (!parts) return;
    for (const part of parts) {
      if (part.parts) findParts(part.parts);
      const fname = part.filename || '';
      if (/\.(xls|xlsx)$/i.test(fname) && part.body) {
        attachments.push({ filename: fname, part });
      }
    }
  }
  findParts(msg.data.payload.parts);

  const results = [];
  for (const { filename, part } of attachments) {
    let data;
    if (part.body.data) {
      data = Buffer.from(part.body.data, 'base64');
    } else if (part.body.attachmentId) {
      const att = await gmail.users.messages.attachments.get({
        userId: 'me', messageId, id: part.body.attachmentId,
      });
      data = Buffer.from(att.data.data, 'base64');
    }
    if (data) {
      results.push({ filename, buffer: data });
      console.log(`    📎 ${filename} (${(data.length / 1024).toFixed(0)} KB)`);
    }
  }
  return { files: results, subject, year };
}

async function markProcessed(gmail, messageId) {
  let labelId;
  try {
    const labels   = await gmail.users.labels.list({ userId: 'me' });
    const existing = (labels.data.labels || []).find(l => l.name === 'processed-report');
    if (existing) {
      labelId = existing.id;
    } else {
      const created = await gmail.users.labels.create({
        userId: 'me',
        requestBody: { name: 'processed-report', labelListVisibility: 'labelShow', messageListVisibility: 'show' },
      });
      labelId = created.data.id;
    }
  } catch (e) {
    console.log('  (could not create label:', e.message + ')');
  }

  await gmail.users.messages.modify({
    userId: 'me',
    id: messageId,
    requestBody: {
      removeLabelIds: ['UNREAD'],
      addLabelIds:    labelId ? [labelId] : [],
    },
  });
}

// ── F&B PARSER ────────────────────────────────────────────────────────────────
function parseNum(v) {
  if (v === '' || v == null) return 0;
  const n = parseFloat(String(v).replace(/\s/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

// Convert Excel date serial (col[3]) to ISO date string "YYYY-MM-DD"
function excelSerialToISO(serial) {
  const s = parseFloat(String(serial).replace(',', '.'));
  if (!s || isNaN(s)) return null;
  const ms = (Math.floor(s) - 25569) * 86400000;
  const d  = new Date(ms);
  return d.toISOString().slice(0, 10);
}

// Reception/hotel files carry NO report-date cell — col[3] is the reservation's
// check-in date, not the report day. The actual day lives in the filename or subject
// (e.g. "Prenotimet ne recepsion 9 Qeshor.xls"). Parse it from there.
const SQ_MONTHS = {
  janar: 1, shkurt: 2, mars: 3, prill: 4, maj: 5, qershor: 6, qeshor: 6,
  korrik: 7, gusht: 8, shtator: 9, tetor: 10, nentor: 11, 'nëntor': 11, dhjetor: 12,
};
function parseReportDateFromName(text, fallbackYear) {
  if (!text) return null;
  const s = String(text).toLowerCase();

  // Numeric: dd-mm-yyyy / dd.mm.yyyy / dd/mm/yyyy (also 2-digit year)
  let m = s.match(/(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})/);
  if (m) {
    let d = +m[1], mo = +m[2], y = +m[3];
    if (y < 100) y += 2000;
    if (d >= 1 && d <= 31 && mo >= 1 && mo <= 12) {
      return `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    }
  }

  // Albanian "DD <muaji>" — e.g. "9 qeshor", "recepsion10 qeshor", "8 Gusht"
  const monthAlt = Object.keys(SQ_MONTHS).join('|');
  m = s.match(new RegExp(`(\\d{1,2})\\s*(${monthAlt})`, 'i'));
  if (m) {
    const d = +m[1], mo = SQ_MONTHS[m[2].toLowerCase()];
    const y = fallbackYear || new Date().getUTCFullYear();
    if (d >= 1 && d <= 31 && mo) {
      return `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    }
  }
  return null;
}

// Extract the report date from col[3] of the first valid invoice row
function extractDateFromBuffer(buffer) {
  const wb   = XLSX.read(buffer, { type: 'buffer', raw: false });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  for (let r = 1; r < rows.length; r++) {
    const col0 = String(rows[r][0]).trim();
    if (col0 === '' || isNaN(Number(col0))) continue;
    const iso = excelSerialToISO(rows[r][3]);
    if (iso) return iso;
  }
  return null;
}

// ── DAILY EXPENSES (purchases report → sheet5 "DAILY EXPENSES") ────────────────
// The purchases ("Blerjet") export lists one section per warehouse ("Magazina: X")
// whose daily total sits under the "Vlera sipas mon. standarte" column. Each
// magazine maps to a column of the DAILY EXPENSES sheet (dates are rows).
const EXP_COL_MAP = {
  'bufe':'G', 'restorant':'C', 'flower restorant':'C', 'pool bar':'D', 'pool bar garden':'F',
  'beach bar':'B', 'brutal':'E', 'garden brutal':'E', 'magazina qendrore':'I',
  'operacionale mikse':'J', 'familja':'N', 'spa':'K', 'magazina garden':'P',
  'magazina garden (investime)':'P', 'investime':'P', 'shpenzime hoteli':'O',
  'marketing':'M', 'mirembajtje':'L', 'mirembajtje dhe riparime':'L',
  'paga':'Q', 'paga & utilitete':'Q', 'overheads':'H', 'overheads f&b':'H',
};
const EXP_EXCLUDE = ['flower tirane', 'tirane']; // separate property — not on this sheet
const normMag = s => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim().replace(/[.:]+$/, '').trim();

// True if the workbook looks like a purchases report (has the standard-currency
// column AND "Magazina:" section rows) — used to route attachments to expenses.
function isExpenseReport(buffer) {
  try {
    const wb = XLSX.read(buffer, { type: 'buffer', raw: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
    let hasStd = false, hasMag = false;
    for (let i = 0; i < Math.min(rows.length, 80); i++) {
      const line = rows[i].map(c => String(c)).join(' ').toLowerCase();
      if (line.includes('vlera sipas mon')) hasStd = true;
      if (/magazina:/i.test(line)) hasMag = true;
      if (hasStd && hasMag) return true;
    }
    return false;
  } catch (e) { return false; }
}

// Parse a purchases report → { serial, isoDate, mags:{ NAME: value } }.
// Report date = the most common floored invoice-date serial ("Data Faturës" column).
function parseExpensesBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', raw: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  let hdr = -1, colStd = -1, colDate = -1;
  for (let i = 0; i < Math.min(rows.length, 12); i++) {
    const std = rows[i].findIndex(c => String(c).toLowerCase().includes('vlera sipas mon'));
    if (std >= 0) { hdr = i; colStd = std; colDate = rows[i].findIndex(c => String(c).toLowerCase().includes('data fatur')); break; }
  }
  if (colStd < 0) return null;
  const mags = {}; const serials = [];
  for (let i = hdr + 1; i < rows.length; i++) {
    const r = rows[i]; const c0 = String(r[0] || '').trim();
    if (/^Magazina:/i.test(c0)) mags[c0.replace(/^Magazina:\s*/i, '').trim()] = parseNum(r[colStd]);
    if (colDate >= 0 && typeof r[colDate] === 'number' && r[colDate] > 40000) serials.push(Math.floor(r[colDate]));
  }
  const freq = {}; serials.forEach(s => freq[s] = (freq[s] || 0) + 1);
  const serial = Object.keys(freq).sort((a, b) => freq[b] - freq[a])[0];
  return { serial: serial ? parseInt(serial) : null, isoDate: serial ? excelSerialToISO(serial) : null, mags };
}

// Map magazine totals → { expValues:{ COL: value }, unmapped:[names] }. 2-dec rounded.
function mapExpenseMagazines(mags) {
  const expValues = {}; const unmapped = [];
  for (const [name, val] of Object.entries(mags)) {
    const key = normMag(name);
    if (EXP_EXCLUDE.includes(key)) continue;
    // Exact match first (keeps literal-"&" keys like "overheads f&b" / "paga & utilitete"),
    // then a fallback that reads "&" as Albanian "dhe" so "mirembajtje & riparime" → …"dhe"… → L,
    // and any future "X & Y" warehouse resolves without a new key.
    const col = EXP_COL_MAP[key] || EXP_COL_MAP[key.replace(/\s*&\s*/g, ' dhe ')];
    if (!col) { unmapped.push(name); continue; }
    expValues[col] = Math.round(((expValues[col] || 0) + val) * 100) / 100;
  }
  return { expValues, unmapped };
}

function parseFnBBuffer(buffer, label) {
  const wb   = XLSX.read(buffer, { type: 'buffer', raw: false });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  console.log(`  [${label}] ${rows.length} rows`);

  // Revenue = last non-empty value in col[6] (total row at bottom of file)
  let revenue = 0;
  for (let r = rows.length - 1; r >= 1; r--) {
    const v = parseNum(rows[r][6]);
    if (v !== 0) { revenue = v; break; }
  }

  // House Use = sum of col[6] (Shuma) for invoice rows where col[7] matches house use names
  let houseUse = 0;
  for (let r = 1; r < rows.length; r++) {
    const col0 = String(rows[r][0]).trim();
    if (col0 === '' || isNaN(Number(col0))) continue;
    const tableName = String(rows[r][7] || '').trim();
    if (!tableName) continue;
    const shuma = parseNum(rows[r][6]);
    if (isHouseUse(tableName) && shuma !== 0) houseUse += shuma;
  }

  revenue  = Math.round(revenue  * 100) / 100;
  houseUse = Math.round(houseUse * 100) / 100;
  console.log(`    Revenue: ${revenue}  |  HouseUse: ${houseUse}`);
  return { revenue, houseUse };
}

function parseHotelBuffer(buffer, isoDate) {
  const wb   = XLSX.read(buffer, { type: 'buffer', raw: false });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // From 2026-06-06 onward also exclude block/agency client names (ITAKA, SAISTOURS, W2M).
  const excludeNames = (isoDate && isoDate >= NEW_EXCL_FROM_DATE)
    ? EXCLUDE_NAMES.concat(EXCLUDE_NAMES_NEW) : EXCLUDE_NAMES;

  console.log(`  [hotel] ${rows.length} rows`);
  let nightsOccupied = 0, totalRevenue = 0;

  for (let r = 1; r < rows.length; r++) {
    const row    = rows[r];
    const bookId = row[0];
    if (bookId === '' || bookId == null) continue;

    const klienti = String(row[5] || '').trim();
    const dite    = parseNum(row[10]);
    const totali  = parseNum(row[11]);
    const nameLo  = klienti.toLowerCase().replace(/\s+/g, ' ');

    if (!excludeNames.some(ex => nameLo.includes(ex))) {
      nightsOccupied += (dite > 0 ? dite : 1);
      totalRevenue   += totali;
    }
  }

  const occupancyPct = Math.round((nightsOccupied / TOTAL_ROOMS) * 10000) / 100;
  totalRevenue = Math.round(totalRevenue * 100) / 100;
  console.log(`    Nights: ${nightsOccupied}/${TOTAL_ROOMS} = ${occupancyPct}%  |  Revenue: ${totalRevenue}`);
  return { occupancyPct, nightsOccupied, nightsAvailable: TOTAL_ROOMS, revenue: totalRevenue };
}

// ── CLASSIFY ATTACHMENT ───────────────────────────────────────────────────────
function classifyFile(filename) {
  const n = filename.toLowerCase();
  if (n.includes('prenotim') || n.includes('recepsion') || n.includes('reception') || n.includes('hotel')) return 'hotel';
  if (n.includes('beach')) return 'beach_bar';
  if (n.includes('pool bar g') || n.includes('pool bar garden') || n.includes('pool_bar_g') || n.includes('poolbar_g') || n.includes('pool_garden') || n.includes('pool garden')) return 'poolbar_g';
  if (n.includes('brutal') || (n.includes('garden') && !n.includes('pool'))) return 'garden';
  if (n.includes('pool bar') || n.includes('pool_bar') || n.includes('poolbar') || n.includes('pool')) return 'pool_bar';
  if (n.includes('restorant') || n.includes('restaurant') || n.includes('restor')) return 'restorant';
  return null;
}

function classifyPoolBarByContent(buffer) {
  try {
    const wb   = XLSX.read(buffer, { type: 'buffer', raw: false });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    for (const row of rows) {
      for (const cell of row) {
        if (String(cell).toLowerCase().includes('garden')) return 'poolbar_g';
      }
    }
  } catch (e) { /* ignore */ }
  return 'pool_bar';
}

// ── XLSX SURGICAL PATCHER ─────────────────────────────────────────────────────
function findRowForDate(xml, isoDate) {
  const d      = new Date(isoDate + 'T00:00:00Z');
  const serial = Math.floor(d.getTime() / 86400000) + 25569;

  const rowRe = /<row r="(\d+)"[^>]*>[\s\S]*?<\/row>/g;
  let m;
  while ((m = rowRe.exec(xml)) !== null) {
    const rowNum = m[1];
    const rowXml = m[0];
    if (new RegExp(`<c r="A${rowNum}"[^>]*><v>${serial}(?:\\.0)?<\\/v><\\/c>`).test(rowXml)) {
      return { rowNum: parseInt(rowNum), serial };
    }
  }
  return null;
}

function getCellStyle(cellXml) {
  const m = cellXml.match(/\bs="(\d+)"/);
  return m ? m[1] : '';
}

function readRefStyles(xml, refRowNum, colLetters) {
  const rowMatch = xml.match(new RegExp(`<row r="${refRowNum}"[^>]*>[\\s\\S]*?<\\/row>`));
  if (!rowMatch) return null;
  const rowXml = rowMatch[0];
  const styles = {};
  for (const col of colLetters) {
    const cellMatch = rowXml.match(new RegExp(`<c r="${col}${refRowNum}"[^>]*>`));
    if (cellMatch) {
      styles[col] = { s: getCellStyle(cellMatch[0]), t: /\bt="([^"]+)"/.exec(cellMatch[0])?.[1] || '' };
    } else {
      styles[col] = { s: '', t: '' };
    }
  }
  return styles;
}

// Patch one F&B (sheet3) date row in-place. Returns the new XML, or null if the
// date row wasn't found (sheet left unchanged).
function patchFnbRow(fnbXml, isoDate, fnbValues) {
  const fnbInfo = findRowForDate(fnbXml, isoDate);
  if (!fnbInfo) {
    console.error(`  FNB: date ${isoDate} not found in sheet!`);
    return null;
  }
  const r   = fnbInfo.rowNum;
  const ser = fnbInfo.serial;
  console.log(`  FNB: date ${isoDate} → row ${r} (serial ${ser})`);

  const ref = readRefStyles(fnbXml, r - 1, ['A','B','C','D','E','F','G','H']);
  const sA  = ref?.A?.s || '23';
  const sB  = ref?.B?.s || '36';
  const sG  = ref?.G?.s || '34';
  const sH  = ref?.H?.s || '35';
  console.log(`    Styles from row ${r-1}: A=s${sA} B=s${sB} G=s${sG} H=s${sH}`);

  const { B, C, D, E, F, G, H } = fnbValues;
  const newRow =
    `<row r="${r}" ht="14.25" customHeight="1">` +
    `<c r="A${r}" s="${sA}"><v>${ser}.0</v></c>` +
    `<c r="B${r}" s="${sB}"><v>${B}</v></c>` +
    `<c r="C${r}" s="${sB}"><v>${C}</v></c>` +
    `<c r="D${r}" s="${sB}"><v>${D}</v></c>` +
    `<c r="E${r}" s="${sB}"><v>${E}</v></c>` +
    `<c r="F${r}" s="${sB}"><v>${F}</v></c>` +
    `<c r="G${r}" s="${sG}"><v>${G}</v></c>` +
    `<c r="H${r}" s="${sH}"><f t="shared" si="1"/><v>${H}</v></c>` +
    `</row>`;
  console.log(`  ✓ FNB row ${r} written`);
  return fnbXml.replace(new RegExp(`<row r="${r}"[^>]*>[\\s\\S]*?<\\/row>`), newRow);
}

// Patch one Hotel (sheet2) date row in-place. Returns the new XML, or null if the
// date row wasn't found (sheet left unchanged).
function patchHotelRow(hotelXml, isoDate, hotelValues) {
  const hotelInfo = findRowForDate(hotelXml, isoDate);
  if (!hotelInfo) {
    console.error(`  Hotel: date ${isoDate} not found in sheet!`);
    return null;
  }
  const r   = hotelInfo.rowNum;
  const ser = hotelInfo.serial;
  console.log(`  Hotel: date ${isoDate} → row ${r} (serial ${ser})`);

  const ref  = readRefStyles(hotelXml, r - 1, ['A','B','C','D','E','F']);
  const sA   = ref?.A?.s || '23';
  const sB   = ref?.B?.s || '17';
  const sBt  = ref?.B?.t || 's';
  const sC   = ref?.C?.s || '24';
  const sD   = ref?.D?.s || '27';
  const sE   = ref?.E?.s || '17';
  const sF   = ref?.F?.s || '28';
  const refRow = hotelXml.match(new RegExp(`<row r="${r-1}"[^>]*>[\\s\\S]*?<\\/row>`))?.[0] || '';
  const ssMatch = refRow.match(new RegExp(`<c r="B${r-1}"[^>]*><v>(\\d+)<\\/v><\\/c>`));
  const ssIdx = ssMatch ? ssMatch[1] : '35';
  console.log(`    Styles from row ${r-1}: A=s${sA} B=s${sB}(t=${sBt},ss=${ssIdx}) C=s${sC} D=s${sD} E=s${sE} F=s${sF}`);

  const { nightsOccupied, nightsAvailable, revenue } = hotelValues;
  // Preserve the per-day available rooms already in col E (maintained from Power BI;
  // varies during June 2026). Recompute occupancy against it; fall back to the parsed
  // default only when the cell is empty — never clobber a real availability value.
  const curHotelRow = hotelXml.match(new RegExp(`<row r="${r}"[^>]*>[\\s\\S]*?<\\/row>`))?.[0] || '';
  const curAvailM = curHotelRow.match(new RegExp(`<c r="E${r}"[^>]*><v>([\\d.]+)<\\/v>`));
  const availRooms = (curAvailM && parseFloat(curAvailM[1]) > 0) ? parseFloat(curAvailM[1]) : nightsAvailable;
  const occDecimal = availRooms > 0 ? nightsOccupied / availRooms : 0;
  const newRow =
    `<row r="${r}" ht="14.25" customHeight="1">` +
    `<c r="A${r}" s="${sA}"><v>${ser}.0</v></c>` +
    `<c r="B${r}" s="${sB}" t="${sBt}"><v>${ssIdx}</v></c>` +
    `<c r="C${r}" s="${sC}"><f t="shared" si="1"/><v>${occDecimal}</v></c>` +
    `<c r="D${r}" s="${sD}"><v>${nightsOccupied}.0</v></c>` +
    `<c r="E${r}" s="${sE}"><v>${availRooms}</v></c>` +
    `<c r="F${r}" s="${sF}"><v>${revenue}</v></c>` +
    `</row>`;
  console.log(`  ✓ Hotel row ${r} written`);
  return hotelXml.replace(new RegExp(`<row r="${r}"[^>]*>[\\s\\S]*?<\\/row>`), newRow);
}

// Read a numeric cell value from a row's XML (0 if empty / not present).
function readCellVal(rowXml, col, r) {
  const m = rowXml.match(new RegExp(`<c r="${col}${r}"[^>]*?>(?:<f[^>]*\\/?>(?:[^<]*<\\/f>)?)?<v>([\\d.\\-]+)<\\/v>`));
  return m ? parseFloat(m[1]) : 0;
}

// Patch one DAILY EXPENSES (sheet5) date row — SELECTIVE per-column update: only the
// magazine columns from the report are touched, every other cell (Beach Bar, SPA,
// Paga & Utilitete, manual entries…) is preserved. TOTAL (col R) is recomputed but
// keeps its shared formula. Returns the new XML, or null if the date row wasn't found.
function patchExpensesRow(expXml, isoDate, expValues) {
  const info = findRowForDate(expXml, isoDate);
  if (!info) { console.error(`  EXP: date ${isoDate} not found in DAILY EXPENSES sheet!`); return null; }
  const r = info.rowNum;
  const rowXml = expXml.match(new RegExp(`<row r="${r}"[^>]*>[\\s\\S]*?<\\/row>`))?.[0];
  if (!rowXml) { console.error(`  EXP: row ${r} XML not found`); return null; }
  console.log(`  EXP: date ${isoDate} → row ${r} (serial ${info.serial})`);

  let newRow = rowXml;
  for (const [col, val] of Object.entries(expValues)) {
    const existing = newRow.match(new RegExp(`<c r="${col}${r}"([^>]*?)(?:\\/>|>[\\s\\S]*?<\\/c>)`));
    let style = '90'; // filled-number style
    if (existing) { const sm = existing[1].match(/\bs="(\d+)"/); if (sm) style = (sm[1] === '22' ? '90' : sm[1]); }
    const cell = `<c r="${col}${r}" s="${style}"><v>${val}</v></c>`;
    newRow = newRow.replace(new RegExp(`<c r="${col}${r}"(?:[^>]*\\/>|[^>]*>[\\s\\S]*?<\\/c>)`), cell);
  }
  // Recompute TOTAL (col R) = sum of B..Q, preserving its shared formula + style.
  let total = 0;
  'BCDEFGHIJKLMNOPQ'.split('').forEach(c => { total += (expValues[c] != null ? expValues[c] : readCellVal(newRow, c, r)); });
  total = Math.round(total * 100) / 100;
  const rM = newRow.match(new RegExp(`<c r="R${r}"([^>]*?)>([\\s\\S]*?)<\\/c>`));
  const rStyle = (rM && rM[1].match(/\bs="(\d+)"/)) ? rM[1].match(/\bs="(\d+)"/)[1] : '27';
  const fM = rM ? rM[2].match(/<f[^>]*\/>|<f[^>]*>[\s\S]*?<\/f>/) : null;
  const rCell = `<c r="R${r}" s="${rStyle}">${fM ? fM[0] : '<f t="shared" si="1"/>'}<v>${total}</v></c>`;
  newRow = newRow.replace(new RegExp(`<c r="R${r}"(?:[^>]*\\/>|[^>]*>[\\s\\S]*?<\\/c>)`), rCell);

  console.log(`  ✓ EXP row ${r} updated: ${Object.entries(expValues).map(([c, v]) => c + '=' + v).join(' ')} → TOTAL ${total}`);
  return expXml.replace(rowXml, newRow);
}

// Resolve the worksheet XML path for a sheet by its DISPLAY NAME, via workbook.xml
// + rels. Hardcoding sheetN.xml is unsafe: a workbook re-save can renumber the
// physical files. On 2026-07-06 exactly this happened — DAILY EXPENSES moved from
// sheet5.xml to sheet4.xml and sheet5.xml became DAILY CASH FLOW, so expenses were
// being written into the cash-flow sheet. Resolving by name is re-save-proof.
async function resolveSheetPath(zip, wantName) {
  const wb   = await zip.file('xl/workbook.xml').async('string');
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const relMap = {};
  for (const m of rels.matchAll(/<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"/g)) relMap[m[1]] = m[2];
  const decode = s => String(s).replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  const norm   = s => decode(s).toUpperCase().replace(/\s+/g, ' ').trim();
  const want   = norm(wantName);
  for (const m of wb.matchAll(/<sheet\b[^>]*\/>/g)) {
    const tag  = m[0];
    const name = (tag.match(/\bname="([^"]*)"/) || [])[1];
    const rid  = (tag.match(/\br:id="([^"]*)"/) || [])[1];
    if (name && norm(name) === want && relMap[rid]) {
      const t = relMap[rid].replace(/^\/?xl\//, '');
      return 'xl/' + (t.startsWith('worksheets/') ? t : 'worksheets/' + t);
    }
  }
  throw new Error(`sheet "${wantName}" not found in workbook.xml`);
}

// Apply a batch of per-date updates to the workbook in ONE download + upload.
// updates: [{ isoDate, fnbValues, hotelValues, expValues, writeFnb, writeHotel, writeExp }]
// Each update targets its own date row, so one email can carry several hotel
// days (and/or an F&B day, and/or a daily-expenses day) and every one lands right.
async function patchXlsx(drive, updates) {
  const list = (updates || []).filter(u => u && (u.writeFnb || u.writeHotel || u.writeExp));
  if (list.length === 0) {
    console.log('  Nothing to write — no sheets flagged.');
    return [];
  }

  console.log('\nDownloading XLSX from Google Drive...');
  const res = await drive.files.get(
    { fileId: XLSX_FILE_ID, alt: 'media' },
    { responseType: 'arraybuffer' }
  );
  const buffer = Buffer.from(res.data);
  console.log(`  ${(buffer.length / 1024).toFixed(0)} KB`);

  const zip = await JSZip.loadAsync(buffer);
  // Resolve each sheet by name (re-save-proof), not by a hardcoded sheetN.xml.
  const fnbPath   = await resolveSheetPath(zip, 'DAILY F&B REVENUES');
  const hotelPath = await resolveSheetPath(zip, 'HOTEL DAILY PERFORMANCE');
  const expPath   = await resolveSheetPath(zip, 'DAILY EXPENSES');
  console.log(`  sheets → F&B:${fnbPath} Hotel:${hotelPath} Expenses:${expPath}`);
  let fnbXml   = await zip.file(fnbPath).async('string');
  let hotelXml = await zip.file(hotelPath).async('string');
  let expXml   = await zip.file(expPath).async('string');
  let fnbChanged = false, hotelChanged = false, expChanged = false;

  // Mirror each update with whether its row actually landed (date row found).
  const results = list.map(u => {
    let fnbWritten = false, hotelWritten = false, expWritten = false;
    if (u.writeFnb) {
      const next = patchFnbRow(fnbXml, u.isoDate, u.fnbValues);
      if (next) { fnbXml = next; fnbChanged = true; fnbWritten = true; }
    }
    if (u.writeHotel) {
      const next = patchHotelRow(hotelXml, u.isoDate, u.hotelValues);
      if (next) { hotelXml = next; hotelChanged = true; hotelWritten = true; }
    }
    if (u.writeExp) {
      const next = patchExpensesRow(expXml, u.isoDate, u.expValues);
      if (next) { expXml = next; expChanged = true; expWritten = true; }
    }
    return { isoDate: u.isoDate, fnbValues: u.fnbValues, hotelValues: u.hotelValues, expValues: u.expValues, fnbWritten, hotelWritten, expWritten };
  });

  if (!fnbChanged && !hotelChanged && !expChanged) {
    console.log('  No rows matched — nothing uploaded.');
    return results;
  }
  if (fnbChanged)   zip.file(fnbPath, fnbXml);
  if (hotelChanged) zip.file(hotelPath, hotelXml);
  if (expChanged)   zip.file(expPath, expXml);

  const outBuf = await zip.generateAsync({
    type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 }
  });
  console.log(`\n  Uploading ${(outBuf.length / 1024).toFixed(0)} KB...`);
  await drive.files.update({
    fileId: XLSX_FILE_ID,
    media: {
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      body: Readable.from(outBuf),
    },
  });
  console.log('  ✓ Uploaded');
  return results;
}

// ── GAS / FLOW DASHBOARD CALL ─────────────────────────────────────────────────
async function callGasEndpoint(isoDate, hotelValues, fnbValues, opts = {}) {
  const { writeFnb = true, writeHotel = true } = opts;
  if (!GAS_ENDPOINT_URL) {
    console.log('  [GAS] GAS_ENDPOINT_URL not set — skipping Flow Dashboard update');
    return;
  }
  const https = require('https');
  const http  = require('http');
  const lib   = GAS_ENDPOINT_URL.startsWith('https') ? https : http;
  const url   = require('url');

  async function post(body) {
    return new Promise((resolve, reject) => {
      const parsed  = url.parse(GAS_ENDPOINT_URL);
      const payload = JSON.stringify(body);
      const opts = {
        hostname: parsed.hostname,
        path:     parsed.path,
        method:   'POST',
        headers:  { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(payload) },
      };
      const req = lib.request(opts, (res) => {
        let data = '';
        res.on('data', c => data += c);
        res.on('end', () => resolve({ status: res.statusCode, body: data }));
      });
      req.on('error', reject);
      req.write(payload);
      req.end();
    });
  }

  try {
    if (writeHotel) {
      const r1 = await post({
        action:          'hotel_performance',
        date:            isoDate,
        nightsOccupied:  hotelValues.nightsOccupied,
        nightsAvailable: hotelValues.nightsAvailable,
        revenue:         hotelValues.revenue,
      });
      console.log(`  [GAS] hotel_performance → HTTP ${r1.status}`);
    } else {
      console.log('  [GAS] hotel_performance skipped — no hotel file in this email');
    }

    if (writeFnb) {
      const r2 = await post({
        action:           'fnb_revenues',
        date:             isoDate,
        flowerRestaurant: fnbValues.B,
        poolBar:          fnbValues.C,
        brutalGarden:     fnbValues.D,
        poolBarGarden:    fnbValues.E,
        beachBar:         fnbValues.F,
        houseUse:         fnbValues.G,
      });
      console.log(`  [GAS] fnb_revenues → HTTP ${r2.status}`);
    } else {
      console.log('  [GAS] fnb_revenues skipped — no F&B file in this email');
    }
  } catch (e) {
    console.error(`  [GAS] ERROR: ${e.message}`);
  }
}

// ── MAIN CHECK CYCLE ──────────────────────────────────────────────────────────
async function runCheck() {
  checkCount++;
  lastCheckTime   = new Date().toISOString();
  lastCheckStatus = 'running';

  console.log('\n' + '═'.repeat(62));
  console.log(` [check #${checkCount}] ${lastCheckTime}`);
  console.log('═'.repeat(62));

  try {
    const auth  = await authorize();
    const gmail = google.gmail({ version: 'v1', auth });
    const drive = google.drive({ version: 'v3', auth });

    console.log('\nSearching Gmail for unread report emails...');
    const messageIds = await findReportEmails(gmail);
    if (messageIds.length === 0) {
      console.log('  No new report emails found.');
      lastCheckStatus = 'ok — no new emails';
      return;
    }
    console.log(`  Found ${messageIds.length} email(s)\n`);

    // Group attachments by email, process each email independently
    // (each email may be a separate date update)
    const emailGroups = [];
    for (const msgId of messageIds) {
      console.log(`Processing email ${msgId}...`);
      const { files, subject, year } = await getAttachments(gmail, msgId);
      const collected   = { restorant: null, pool_bar: null, poolbar_g: null, garden: null, beach_bar: null };
      const hotelFiles  = [];   // keep ALL hotel/reception files — one per day, each its own date row
      const expenseFiles = [];  // purchases ("Blerjet") reports → DAILY EXPENSES sheet
      let hasReport = false;

      for (const { filename, buffer } of files) {
        // Purchases report → daily expenses. Detected by CONTENT (the "Vlera sipas mon" +
        // "Magazina:" structure) OR by the file NAME ("Blerjet …" / "Shpenzime …"), so a
        // format variant that the content check might miss still routes here. If the file
        // then turns out not to be a real purchases report, parseExpensesBuffer returns
        // null / no magazines and it's skipped downstream — safe either way.
        if (isExpenseReport(buffer) || /blerje|shpenzim/i.test(filename)) {
          expenseFiles.push({ filename, buffer });
          console.log(`    → expenses: ${filename}`);
          hasReport = true;
          continue;
        }

        let type = classifyFile(filename);

        // Hotel/reception files: collect every one. A single email may carry several
        // days of "Prenotimet ne recepsion" — each goes to its own date row later.
        if (type === 'hotel') {
          hotelFiles.push({ filename, buffer });
          console.log(`    → hotel: ${filename}`);
          hasReport = true;
          continue;
        }

        if (type === 'pool_bar' && collected['pool_bar']) {
          const contentType = classifyPoolBarByContent(buffer);
          type = contentType;
          if (contentType === 'pool_bar') {
            const existingType = classifyPoolBarByContent(collected['pool_bar'].buffer);
            if (existingType === 'poolbar_g') {
              collected['poolbar_g'] = collected['pool_bar'];
              collected['pool_bar']  = { filename, buffer };
              hasReport = true;
              continue;
            }
          }
        }

        if (type && !collected[type]) {
          collected[type] = { filename, buffer };
          console.log(`    → ${type}: ${filename}`);
          hasReport = true;
        } else if (!type) {
          console.log(`    → Could not classify: ${filename}`);
        }
      }
      if (hasReport) emailGroups.push({ msgId, collected, hotelFiles, expenseFiles, subject, year });
    }

    for (const { msgId, collected, hotelFiles, expenseFiles, subject, year } of emailGroups) {
      // Each email may carry an F&B day (one combined row) and/or several hotel days
      // (one row each). Build a list of per-date updates, then write them all at once.
      const updates = [];
      let fnbDate = null;   // F&B invoice date — also the fallback date for an undated hotel file

      // ── F&B: one combined row for one date ────────────────────────────────
      const hasFnb = ['restorant','pool_bar','poolbar_g','garden','beach_bar'].some(k => collected[k]);
      if (hasFnb) {
        for (const key of ['restorant', 'pool_bar', 'garden', 'poolbar_g', 'beach_bar']) {
          if (collected[key]) { fnbDate = extractDateFromBuffer(collected[key].buffer); if (fnbDate) break; }
        }
        if (!fnbDate) {
          console.error(`  ERROR: could not extract F&B date for email ${msgId}`);
        } else {
          console.log(`\n F&B date ${fnbDate} — parsing...`);
          const fnb = { restorant: { revenue: 0, houseUse: 0 }, pool_bar: { revenue: 0, houseUse: 0 }, poolbar_g: { revenue: 0, houseUse: 0 }, garden: { revenue: 0, houseUse: 0 }, beach_bar: { revenue: 0, houseUse: 0 } };
          for (const key of Object.keys(fnb)) {
            if (collected[key]) {
              try { fnb[key] = parseFnBBuffer(collected[key].buffer, key); }
              catch (e) { console.error(`  ERROR [${key}]: ${e.message}`); }
            } else {
              console.log(`  [${key}] not provided — using 0`);
            }
          }
          const totalHouseUse = Object.values(fnb).reduce((s, v) => s + v.houseUse, 0);
          const houseUseNeg   = -(Math.round(totalHouseUse * 100) / 100);
          const fnbValues = {
            B: fnb.restorant.revenue,
            C: fnb.pool_bar.revenue,
            D: fnb.garden.revenue,
            E: fnb.poolbar_g.revenue,
            F: fnb.beach_bar.revenue,
            G: houseUseNeg,
            H: Math.round((fnb.restorant.revenue + fnb.pool_bar.revenue + fnb.garden.revenue + fnb.poolbar_g.revenue + fnb.beach_bar.revenue + houseUseNeg) * 100) / 100,
          };
          console.log(` F&B: R=${fnbValues.B} PB=${fnbValues.C} G=${fnbValues.D} PBG=${fnbValues.E} BB=${fnbValues.F} HU=${fnbValues.G} T=${fnbValues.H}`);
          updates.push({ isoDate: fnbDate, fnbValues, hotelValues: null, writeFnb: true, writeHotel: false });
        }
      }

      // ── Hotel: one row per file, each at its own date ─────────────────────
      // The reception file has no report-date cell, so the date comes from the
      // filename ("...9 Qeshor.xls"), then the email subject, then the F&B date
      // (the all-6 bundle, where the hotel file is undated). Never col[3].
      if (hotelFiles.length) {
        console.log(`\n Parsing Hotel (${hotelFiles.length} file(s))...`);
        for (const hf of hotelFiles) {
          // Date precedence: the filename ("...9 Qeshor.xls") wins. Otherwise, in a
          // 6-doc bundle, use the F&B invoice date so the hotel row lines up with the
          // F&B row; finally fall back to the email subject ("Raporti Ditor 9 Qeshor").
          const fromName    = parseReportDateFromName(hf.filename, year);
          const fromSubject = parseReportDateFromName(subject, year);
          const hDate = fromName || fnbDate || fromSubject;
          if (!hDate) {
            console.error(`  ERROR: no date for hotel file "${hf.filename}" — not in filename or subject, and no F&B file to borrow it from. Skipped. Rename the file with the day, e.g. "Prenotimet ne recepsion 9 Qeshor.xls".`);
            continue;
          }
          const dateSrc = fromName ? 'filename' : (fnbDate ? 'F&B reports' : 'subject');
          console.log(`  date source for ${hf.filename}: ${dateSrc} → ${hDate}`);
          let hv;
          try { hv = parseHotelBuffer(hf.buffer, hDate); }
          catch (e) { console.error(`  ERROR [hotel ${hf.filename}]: ${e.message}`); continue; }
          console.log(`  ${hf.filename} → ${hDate}: ${hv.nightsOccupied}/${TOTAL_ROOMS} = ${hv.occupancyPct}%  Rev: ${hv.revenue}`);
          updates.push({ isoDate: hDate, fnbValues: null, hotelValues: hv, writeFnb: false, writeHotel: true });
        }
      }

      // ── Daily Expenses: one row per purchases report, at the report's own date ──
      if (expenseFiles.length) {
        console.log(`\n Parsing Daily Expenses (${expenseFiles.length} file(s))...`);
        for (const ef of expenseFiles) {
          let parsed;
          try { parsed = parseExpensesBuffer(ef.buffer); }
          catch (e) { console.error(`  ERROR [expenses ${ef.filename}]: ${e.message}`); continue; }
          if (!parsed || !parsed.isoDate) { console.error(`  ERROR: could not read date from "${ef.filename}" — skipped.`); continue; }
          const { expValues, unmapped } = mapExpenseMagazines(parsed.mags);
          if (unmapped.length) console.warn(`  ⚠ EXP unmapped magazines (skipped): ${unmapped.join(', ')}`);
          if (Object.keys(expValues).length === 0) { console.error(`  ERROR: no mappable magazines in "${ef.filename}" — skipped.`); continue; }
          console.log(`  ${ef.filename} → ${parsed.isoDate}: ${Object.entries(expValues).map(([c, v]) => c + '=' + v).join(' ')}`);
          updates.push({ isoDate: parsed.isoDate, fnbValues: null, hotelValues: null, expValues, writeFnb: false, writeHotel: false, writeExp: true });
        }
      }

      if (updates.length === 0) {
        console.error(`  ERROR: nothing usable to write for email ${msgId} — leaving it UNPROCESSED for retry`);
        continue;
      }
      const dateList = [...new Set(updates.map(u => u.isoDate))].sort();
      console.log(`\n Dates in this email: ${dateList.join(', ')}`);

      // Write every row to the Sample Power BI workbook in one download + upload
      console.log('\n Writing to Sample Power BI...');
      const results = await patchXlsx(drive, updates);
      const written = results.filter(r => r.fnbWritten || r.hotelWritten || r.expWritten);

      // If NOTHING landed (e.g. the date row wasn't found), do NOT mark the email
      // processed — leave it so a fix + retry can fill it later.
      if (written.length === 0) {
        console.error(`  ⚠ No rows matched for email ${msgId} (dates ${dateList.join(', ')} not in sheet?) — leaving it UNPROCESSED.`);
        lastCheckStatus = `warn — no rows matched for ${dateList.join(', ')}`;
        continue;
      }

      // Write each successfully-landed F&B/hotel date to the Flow Dashboard (GAS).
      // Daily-expenses rows live only in the workbook (no GAS handler) — skip those.
      const gasWritten = written.filter(r => r.fnbWritten || r.hotelWritten);
      if (gasWritten.length) {
        console.log('\n Writing to Flow Dashboard...');
        for (const r of gasWritten) {
          await callGasEndpoint(r.isoDate, r.hotelValues, r.fnbValues, { writeFnb: r.fnbWritten, writeHotel: r.hotelWritten });
        }
      }

      // Mark this email as processed
      try {
        await markProcessed(gmail, msgId);
        console.log(`  ✓ marked ${msgId}`);
      } catch (e) {
        console.log(`  (could not mark ${msgId}: ${e.message})`);
      }

      lastProcessed = new Date().toISOString();
      totalProcessed++;

      const writtenDates = [...new Set(written.map(r => r.isoDate))].sort();
      console.log('\n' + '═'.repeat(62));
      console.log(` ✅ ${writtenDates.join(', ')} u shkrua saktë.`);
      console.log('═'.repeat(62) + '\n');
    }

    lastCheckStatus = `ok — processed ${emailGroups.length} email(s)`;

  } catch (e) {
    lastCheckStatus = `error: ${e.message}`;
    console.error(`\n[check] ERROR: ${e.message}\n${e.stack}\n`);
  }
}

// ── STARTUP ───────────────────────────────────────────────────────────────────
console.log('\n' + '═'.repeat(62));
console.log(' Flower Hotel Gmail Report Processor');
console.log(` Poll interval: every ${CHECK_INTERVAL_MS / 60000} minute(s)`);
console.log(' Watching: flowreport26@gmail.com');
console.log('═'.repeat(62) + '\n');

// Start health server
startHealthServer();

// Start self-ping
startSelfPing();

// Run first check immediately, then on interval
runCheck();
setInterval(runCheck, CHECK_INTERVAL_MS);
