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
let lastCheckTime    = null;
let lastCheckStatus  = 'not started';
let lastProcessed    = null;
let totalProcessed   = 0;
let checkCount       = 0;

// ── EXPRESS HEALTH SERVER ─────────────────────────────────────────────────────
function startHealthServer() {
  const app = express();

  app.get('/', (req, res) => {
    res.json({
      service:       'Flower Hotel Gmail Report Processor',
      status:        'running',
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
  const primaryQuery = 'subject:"Raporti Ditor" has:attachment -label:processed-report in:inbox';
  const r1 = await gmail.users.messages.list({ userId: 'me', q: primaryQuery, maxResults: 20 });
  if ((r1.data.messages || []).length > 0) {
    console.log('  (matched subject: Raporti Ditor)');
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
  return results;
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

function parseHotelBuffer(buffer) {
  const wb   = XLSX.read(buffer, { type: 'buffer', raw: false });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

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

    if (!EXCLUDE_NAMES.some(ex => nameLo.includes(ex))) {
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

async function patchXlsx(drive, isoDate, fnbValues, hotelValues) {
  console.log('\nDownloading XLSX from Google Drive...');
  const res = await drive.files.get(
    { fileId: XLSX_FILE_ID, alt: 'media' },
    { responseType: 'arraybuffer' }
  );
  const buffer = Buffer.from(res.data);
  console.log(`  ${(buffer.length / 1024).toFixed(0)} KB`);

  const zip = await JSZip.loadAsync(buffer);

  // ── FNB (sheet3.xml) ──────────────────────────────────────────
  let fnbXml = await zip.file('xl/worksheets/sheet3.xml').async('string');
  const fnbInfo = findRowForDate(fnbXml, isoDate);
  if (!fnbInfo) {
    console.error(`  FNB: date ${isoDate} not found in sheet!`);
  } else {
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
    fnbXml = fnbXml.replace(new RegExp(`<row r="${r}"[^>]*>[\\s\\S]*?<\\/row>`), newRow);
    zip.file('xl/worksheets/sheet3.xml', fnbXml);
    console.log(`  ✓ FNB row ${r} written`);
  }

  // ── Hotel (sheet2.xml) ────────────────────────────────────────
  let hotelXml = await zip.file('xl/worksheets/sheet2.xml').async('string');
  const hotelInfo = findRowForDate(hotelXml, isoDate);
  if (!hotelInfo) {
    console.error(`  Hotel: date ${isoDate} not found in sheet!`);
  } else {
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

    const { occupancyPct, nightsOccupied, nightsAvailable, revenue } = hotelValues;
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
    hotelXml = hotelXml.replace(new RegExp(`<row r="${r}"[^>]*>[\\s\\S]*?<\\/row>`), newRow);
    zip.file('xl/worksheets/sheet2.xml', hotelXml);
    console.log(`  ✓ Hotel row ${r} written`);
  }

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
}

// ── GAS / FLOW DASHBOARD CALL ─────────────────────────────────────────────────
async function callGasEndpoint(isoDate, hotelValues, fnbValues) {
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
    const r1 = await post({
      action:          'hotel_performance',
      date:            isoDate,
      nightsOccupied:  hotelValues.nightsOccupied,
      nightsAvailable: hotelValues.nightsAvailable,
      revenue:         hotelValues.revenue,
    });
    console.log(`  [GAS] hotel_performance → HTTP ${r1.status}`);

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
      const attachments = await getAttachments(gmail, msgId);
      const collected   = { restorant: null, pool_bar: null, poolbar_g: null, garden: null, beach_bar: null, hotel: null };
      let hasReport = false;

      for (const { filename, buffer } of attachments) {
        let type = classifyFile(filename);

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
      if (hasReport) emailGroups.push({ msgId, collected });
    }

    for (const { msgId, collected } of emailGroups) {
      // Extract date from Excel col[3] — NOT from today's date
      let isoDate = null;
      for (const key of ['restorant', 'pool_bar', 'garden', 'poolbar_g', 'beach_bar']) {
        if (collected[key]) {
          isoDate = extractDateFromBuffer(collected[key].buffer);
          if (isoDate) break;
        }
      }
      if (!isoDate && collected.hotel) {
        isoDate = extractDateFromBuffer(collected.hotel.buffer);
      }
      if (!isoDate) {
        console.error(`  ERROR: could not extract date from Excel for email ${msgId}`);
        continue;
      }
      console.log(`\n  Data e raportit: ${isoDate}`);

      // Parse F&B
      console.log('\n Parsing F&B...');
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

      // Parse Hotel
      console.log('\n Parsing Hotel...');
      let hotelValues = { occupancyPct: 0, nightsOccupied: 0, nightsAvailable: TOTAL_ROOMS, revenue: 0 };
      if (collected.hotel) {
        try { hotelValues = parseHotelBuffer(collected.hotel.buffer); }
        catch (e) { console.error(`  ERROR [hotel]: ${e.message}`); }
      } else {
        console.log('  Hotel file not provided — using 0');
      }
      console.log(` Hotel: ${hotelValues.nightsOccupied}/${TOTAL_ROOMS} = ${hotelValues.occupancyPct}%  Rev: ${hotelValues.revenue}`);

      // Write to Sample Power BI (XLSX)
      console.log('\n Writing to Sample Power BI...');
      await patchXlsx(drive, isoDate, fnbValues, hotelValues);

      // Write to Flow Dashboard (GAS)
      console.log('\n Writing to Flow Dashboard...');
      await callGasEndpoint(isoDate, hotelValues, fnbValues);

      // Mark this email as processed
      try {
        await markProcessed(gmail, msgId);
        console.log(`  ✓ marked ${msgId}`);
      } catch (e) {
        console.log(`  (could not mark ${msgId}: ${e.message})`);
      }

      lastProcessed = new Date().toISOString();
      totalProcessed++;

      console.log('\n' + '═'.repeat(62));
      console.log(` ✅ ${isoDate} u shkrua saktë.`);
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
