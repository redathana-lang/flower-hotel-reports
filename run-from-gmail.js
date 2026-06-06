/**
 * run-from-gmail.js
 * ─────────────────────────────────────────────────────────────────────────────
 * Reads XLS report attachments sent to flowreport26@gmail.com,
 * parses the Trinosoft data, and writes it into the XLSX Google Sheet.
 *
 * Workflow:
 *   1. Search Gmail inbox for unread emails that have .xls/.xlsx attachments
 *   2. Download those attachments into memory
 *   3. Parse F&B outlets and Hotel occupancy (same logic as parse-trinosoft.js)
 *   4. Surgical-patch the correct date row in the XLSX on Google Drive
 *   5. Mark the processed emails as read + apply label "processed-report"
 *
 * Usage:
 *   node run-from-gmail.js              ← process any new report emails right now
 *   node run-from-gmail.js 2026-05-08   ← specify a date explicitly
 *   node run-from-gmail.js --watch      ← keep running, auto-check every 5 min
 *   node run-from-gmail.js --watch 10   ← check every 10 minutes
 *
 * Send reports multiple times per day — each new email you send will be
 * picked up automatically and will overwrite the previous values in the sheet.
 *
 * First run: opens browser for Google authorization (Drive + Gmail).
 * After that: fully automatic.
 */

'use strict';

const { google } = require('googleapis');
const XLSX       = require('xlsx');
const JSZip      = require('jszip');
const fs         = require('fs');
const path       = require('path');
const http       = require('http');
const { Readable } = require('stream');

// ── CONFIG ────────────────────────────────────────────────────────────────────
const XLSX_FILE_ID  = '1abLRrgklWeV3wx-KEmA0u4SgCH5ebw3s';
const CREDENTIALS   = path.join(__dirname, 'credentials.json');
const TOKEN_PATH    = path.join(__dirname, 'token.json');
const REDIRECT_PORT = 3456;
const REDIRECT_URI  = `http://localhost:${REDIRECT_PORT}`;
const TOTAL_ROOMS   = 110;

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

// Room/table patterns that map to House Use in F&B
const HOUSE_USE_PATTERNS = ['vila 1', 'vila2', 'vila 2', 'dhoma 313', 'fature qerasje'];

// ── AUTH ──────────────────────────────────────────────────────────────────────
async function authorize() {
  if (!fs.existsSync(CREDENTIALS)) {
    console.error('\nERROR: credentials.json not found.');
    process.exit(1);
  }
  const creds = JSON.parse(fs.readFileSync(CREDENTIALS));
  const { client_id, client_secret } = creds.installed || creds.web;
  const auth = new google.auth.OAuth2(client_id, client_secret, REDIRECT_URI);

  // Check cached token — if it lacks Gmail scope, delete and re-auth
  if (fs.existsSync(TOKEN_PATH)) {
    const saved = JSON.parse(fs.readFileSync(TOKEN_PATH));
    if (saved.scope && saved.scope.includes('gmail')) {
      auth.setCredentials(saved);
      return auth;
    }
    console.log('\nToken missing Gmail scope — re-authorizing...');
    fs.unlinkSync(TOKEN_PATH);
  }

  // First-time / scope-upgrade auth
  const authUrl = auth.generateAuthUrl({ access_type: 'offline', scope: SCOPES });
  const code = await new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      const c = new URL(req.url, REDIRECT_URI).searchParams.get('code');
      if (!c) { res.end('No code'); return; }
      res.end('<h2 style="font-family:sans-serif;color:green">Authorization successful! You can close this tab.</h2>');
      server.close();
      resolve(c);
    });
    server.on('error', reject);
    server.listen(REDIRECT_PORT, () => {
      console.log('\n' + '='.repeat(60));
      console.log(' AUTHORIZATION REQUIRED');
      console.log(' Opening browser — sign in with: flowreport26@gmail.com');
      console.log('='.repeat(60) + '\n');
      const { exec } = require('child_process');
      exec(`start "" "${authUrl}"`);
      console.log('If browser did not open, visit:\n' + authUrl + '\n');
    });
  });

  const { tokens } = await auth.getToken(code);
  auth.setCredentials(tokens);
  fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
  console.log('Token saved.\n');
  return auth;
}

// ── GMAIL HELPERS ─────────────────────────────────────────────────────────────

/**
 * Search Gmail for unread messages that have XLS/XLSX attachments.
 * Returns array of message IDs.
 */
async function findReportEmails(gmail, isoDate) {
  // When isoDate is given, restrict to emails received on that date.
  // When isoDate is null, search ALL unprocessed report emails (catch-up mode,
  // matches gmail-worker.js behavior).
  let dateFilter = '';
  if (isoDate) {
    const d      = new Date(isoDate + 'T00:00:00Z');
    const after  = isoDate.replace(/-/g, '/');
    const nextD  = new Date(d.getTime() + 86400000);
    const before = nextD.toISOString().slice(0,10).replace(/-/g, '/');
    dateFilter   = ` after:${after} before:${before}`;
  }

  // Primary: emails with subject "Raporti Ditor" + XLS attachment, not yet processed
  const primaryQuery = `subject:"Raporti Ditor" has:attachment -label:processed-report in:inbox${dateFilter}`;
  const r1 = await gmail.users.messages.list({ userId: 'me', q: primaryQuery, maxResults: 20 });
  if ((r1.data.messages || []).length > 0) {
    console.log(`  (matched subject: Raporti Ditor${isoDate ? `, date: ${isoDate}` : ''})`);
    return r1.data.messages.map(m => m.id);
  }
  // Fallback: any unprocessed email with attachments
  const fallbackQuery = `has:attachment -label:processed-report in:inbox${dateFilter}`;
  const r2 = await gmail.users.messages.list({ userId: 'me', q: fallbackQuery, maxResults: 20 });
  if ((r2.data.messages || []).length > 0) {
    console.log(`  (fallback: any unprocessed attachment email${isoDate ? `, date: ${isoDate}` : ''})`);
  }
  return (r2.data.messages || []).map(m => m.id);
}

/**
 * Download all XLS/XLSX attachments from a message.
 * Returns array of { filename, buffer }.
 */
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
      // Inline data (base64url)
      data = Buffer.from(part.body.data, 'base64');
    } else if (part.body.attachmentId) {
      // Fetch separate attachment
      const att = await gmail.users.messages.attachments.get({
        userId: 'me',
        messageId,
        id: part.body.attachmentId,
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

/**
 * Mark a message as read and add label "processed-report".
 */
async function markProcessed(gmail, messageId) {
  // Ensure the label exists
  let labelId;
  try {
    const labels = await gmail.users.labels.list({ userId: 'me' });
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
      addLabelIds: labelId ? [labelId] : [],
    },
  });
}

// ── F&B PARSER ────────────────────────────────────────────────────────────────
function parseNum(v) {
  if (v === '' || v == null) return 0;
  const n = parseFloat(String(v).replace(/\s/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

// Excel date serial → ISO date string (YYYY-MM-DD)
function excelSerialToISO(serial) {
  const s = parseFloat(String(serial).replace(',', '.'));
  if (!s || isNaN(s)) return null;
  const ms = (Math.floor(s) - 25569) * 86400000;
  return new Date(ms).toISOString().slice(0, 10);
}

// Report date lives in col[3] of the first valid invoice row of any Trinosoft sheet.
// Returns ISO date string, or null if not found.
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

function isHouseUse(tableStr) {
  const s = String(tableStr).toLowerCase().trim();
  return HOUSE_USE_PATTERNS.some(p => s.includes(p));
}

function parseFnBBuffer(buffer, label) {
  const wb   = XLSX.read(buffer, { type: 'buffer', raw: false });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  console.log(`  [${label}] ${rows.length} rows`);
  let revenue = 0, houseUse = 0;

  for (let r = 1; r < rows.length; r++) {
    const row  = rows[r];
    const col0 = String(row[0]).trim();
    if (col0 === '' || isNaN(Number(col0))) continue;

    const shuma     = parseNum(row[6]);
    const skonto    = parseNum(row[8]);
    const tableName = String(row[7] || '').trim();
    const net       = shuma - skonto;

    if (isHouseUse(tableName)) {
      houseUse += net;
    } else {
      revenue += net;
    }
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
/**
 * Map a filename to its report type.
 * Returns 'restorant' | 'pool_bar' | 'poolbar_g' | 'garden' | 'hotel' | null
 */
function classifyFile(filename) {
  const n = filename.toLowerCase();
  if (n.includes('prenotim') || n.includes('recepsion') || n.includes('reception') || n.includes('hotel')) return 'hotel';
  if (n.includes('beach')) return 'beach_bar';
  // "pool bar garden" / "pool bar g" must be checked BEFORE plain "pool bar"
  if (n.includes('pool bar g') || n.includes('pool bar garden') || n.includes('pool_bar_g') || n.includes('poolbar_g') || n.includes('pool_garden') || n.includes('pool garden')) return 'poolbar_g';
  if (n.includes('brutal') || (n.includes('garden') && !n.includes('pool'))) return 'garden';
  if (n.includes('pool bar') || n.includes('pool_bar') || n.includes('poolbar') || n.includes('pool')) return 'pool_bar';
  if (n.includes('restorant') || n.includes('restaurant') || n.includes('restor')) return 'restorant';
  return null;
}

/**
 * When two files both classify as 'pool_bar', use the XLS content to
 * tell apart pool_bar vs poolbar_g. The pool bar garden file usually has
 * table names containing "garden" or "pgarden" in column H.
 */
function classifyPoolBarByContent(buffer) {
  try {
    const wb   = XLSX.read(buffer, { type: 'buffer', raw: false });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    // Look for any cell in the sheet that references "garden"
    for (const row of rows) {
      for (const cell of row) {
        if (String(cell).toLowerCase().includes('garden')) return 'poolbar_g';
      }
    }
  } catch (e) { /* ignore */ }
  return 'pool_bar';
}

// ── XLSX SURGICAL PATCHER ─────────────────────────────────────────────────────
function encodeXml(str) {
  return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/**
 * Get the Excel row number and date serial for a given ISO date.
 * Searches sheet2.xml (Hotel) or sheet3.xml (FNB) for the correct row.
 */
function findRowForDate(xml, isoDate) {
  // Excel serial: days since Unix epoch (midnight UTC) + 25569
  // Must use midnight UTC + Math.floor to avoid rounding to the next day
  const d = new Date(isoDate + 'T00:00:00Z');
  const serial = Math.floor(d.getTime() / 86400000) + 25569;

  // Try to find the row by matching the date serial in column A
  const rowRe = /<row r="(\d+)"[^>]*>[\s\S]*?<\/row>/g;
  let m;
  while ((m = rowRe.exec(xml)) !== null) {
    const rowNum = m[1];
    const rowXml = m[0];
    // Look for A cell value that matches the serial (allow .0 suffix)
    if (new RegExp(`<c r="A${rowNum}"[^>]*><v>${serial}(?:\\.0)?<\\/v><\\/c>`).test(rowXml)) {
      return { rowNum: parseInt(rowNum), serial };
    }
  }
  return null;
}

/**
 * Extract the style index (s="N") from a cell element string.
 * Returns the numeric style index, or '' if not set.
 */
function getCellStyle(cellXml) {
  const m = cellXml.match(/\bs="(\d+)"/);
  return m ? m[1] : '';
}

/**
 * Read styles from a reference row (one row before the target).
 * Returns { sA, sB, sG, sH, sBtype, sC, sD, sE } for FNB
 * or { sA, sB, sBtype, sC, sD, sE, sF } for Hotel.
 */
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
    const r = fnbInfo.rowNum;
    const ser = fnbInfo.serial;
    console.log(`  FNB: date ${isoDate} → row ${r} (serial ${ser})`);

    // Read styles from the row immediately before (guaranteed to be a valid data row)
    const ref = readRefStyles(fnbXml, r - 1, ['A','B','C','D','E','F','G','H']);
    const sA = ref?.A?.s || '23';
    const sB = ref?.B?.s || '36';
    const sG = ref?.G?.s || '34';
    const sH = ref?.H?.s || '35';
    console.log(`    Reference styles from row ${r-1}: A=s${sA} B=s${sB} G=s${sG} H=s${sH}`);

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

    // Read styles from the row immediately before
    const ref = readRefStyles(hotelXml, r - 1, ['A','B','C','D','E','F']);
    const sA  = ref?.A?.s || '23';
    const sB  = ref?.B?.s || '17';
    const sBt = ref?.B?.t || 's';   // usually t="s" (shared string)
    const sC  = ref?.C?.s || '24';
    const sD  = ref?.D?.s || '27';
    const sE  = ref?.E?.s || '17';
    const sF  = ref?.F?.s || '28';
    // Find the shared string index for "FLOWER HOTEL & SPA" from reference row B
    const refRow = hotelXml.match(new RegExp(`<row r="${r-1}"[^>]*>[\\s\\S]*?<\\/row>`))?.[0] || '';
    const ssMatch = refRow.match(new RegExp(`<c r="B${r-1}"[^>]*><v>(\\d+)<\\/v><\\/c>`));
    const ssIdx = ssMatch ? ssMatch[1] : '35';
    console.log(`    Reference styles from row ${r-1}: A=s${sA} B=s${sB}(t=${sBt},ss=${ssIdx}) C=s${sC} D=s${sD} E=s${sE} F=s${sF}`);

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

  // Re-zip and upload
  const outBuf = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
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

// ── MAIN ──────────────────────────────────────────────────────────────────────
async function main() {
  // Date can be passed as first non-flag argument. If omitted, the date is
  // read from inside each Excel file (col[3] of the first invoice row).
  const dateArg = process.argv.slice(2).find(a => /^\d{4}-\d{2}-\d{2}$/.test(a));
  let DATE = dateArg || null;

  console.log('\n' + '═'.repeat(62));
  console.log(' Flower Hotel — Gmail Report Processor');
  console.log(` Date: ${DATE || '(auto — extracted from each Excel file)'}`);
  console.log('═'.repeat(62));

  const auth  = await authorize();
  const gmail = google.gmail({ version: 'v1', auth });
  const drive = google.drive({ version: 'v3', auth });

  // ── Find unread emails with XLS attachments ───────────────────
  console.log('\nSearching Gmail for unread report emails...');
  const messageIds = await findReportEmails(gmail, DATE);
  if (messageIds.length === 0) {
    console.log('  No unread emails with XLS attachments found.');
    console.log('  Send the Trinosoft XLS files to flowreport26@gmail.com and run again.\n');
    return;
  }
  console.log(`  Found ${messageIds.length} email(s)\n`);

  // ── Collect all attachments ───────────────────────────────────
  const collected = {
    restorant: null,
    pool_bar:  null,
    poolbar_g: null,
    garden:    null,
    beach_bar: null,
    hotel:     null,
  };

  const processedMsgIds = [];

  for (const msgId of messageIds) {
    console.log(`Processing email ${msgId}...`);
    const attachments = await getAttachments(gmail, msgId);
    let hasReport = false;

    for (const { filename, buffer } of attachments) {
      let type = classifyFile(filename);

      if (type === 'pool_bar' && collected['pool_bar']) {
        // Two pool-bar files — use content to distinguish pool_bar vs poolbar_g
        const contentType = classifyPoolBarByContent(buffer);
        console.log(`    → Second pool_bar file — content says: ${contentType}`);
        type = contentType;
        // Also re-check the already-collected pool_bar by content
        if (contentType === 'pool_bar') {
          const existingType = classifyPoolBarByContent(collected['pool_bar'].buffer);
          if (existingType === 'poolbar_g') {
            collected['poolbar_g'] = collected['pool_bar'];
            collected['pool_bar'] = { filename, buffer };
            console.log(`    → Swapped: existing pool_bar is actually poolbar_g`);
            hasReport = true;
            continue;
          }
        }
      }

      if (type && !collected[type]) {
        collected[type] = { filename, buffer };
        console.log(`    → Classified as: ${type}`);
        hasReport = true;
      } else if (type && collected[type]) {
        console.log(`    → ${type} already collected, skipping duplicate`);
      } else {
        console.log(`    → Could not classify: ${filename}`);
      }
    }
    if (hasReport) processedMsgIds.push(msgId);
  }

  // ── Check which files we have ─────────────────────────────────
  console.log('\n' + '─'.repeat(62));
  console.log(' Collected files:');
  for (const [key, val] of Object.entries(collected)) {
    console.log(`  ${key.padEnd(12)} : ${val ? '✓ ' + val.filename : '✗ missing'}`);
  }

  // ── Resolve report date from file content (when no CLI date given) ──
  if (!DATE) {
    const seen = {};
    for (const key of ['restorant', 'pool_bar', 'garden', 'poolbar_g', 'beach_bar', 'hotel']) {
      if (!collected[key]) continue;
      const iso = extractDateFromBuffer(collected[key].buffer);
      if (iso) seen[key] = iso;
    }
    const dates = Object.values(seen);
    if (dates.length === 0) {
      console.error('\n  ERROR: could not extract a date from any Excel file.');
      console.error('  Pass an explicit date:  node run-from-gmail.js YYYY-MM-DD\n');
      process.exit(1);
    }
    const unique = [...new Set(dates)];
    if (unique.length > 1) {
      console.error('\n  ERROR: files disagree on report date — refusing to write:');
      for (const [k, v] of Object.entries(seen)) console.error(`    ${k.padEnd(12)} : ${v}`);
      console.error('  Pass an explicit date to override:  node run-from-gmail.js YYYY-MM-DD\n');
      process.exit(1);
    }
    DATE = unique[0];
    console.log(`\n  Report date (from file content): ${DATE}`);
  }

  // ── Parse F&B ─────────────────────────────────────────────────
  console.log('\n' + '─'.repeat(62));
  console.log(' Parsing F&B...');
  const fnb = {
    restorant: { revenue: 0, houseUse: 0 },
    pool_bar:  { revenue: 0, houseUse: 0 },
    poolbar_g: { revenue: 0, houseUse: 0 },
    garden:    { revenue: 0, houseUse: 0 },
    beach_bar: { revenue: 0, houseUse: 0 },
  };

  for (const key of Object.keys(fnb)) {
    if (collected[key]) {
      try {
        fnb[key] = parseFnBBuffer(collected[key].buffer, key);
      } catch (e) {
        console.error(`  ERROR [${key}]: ${e.message}`);
      }
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
    H: Math.round((fnb.restorant.revenue + fnb.pool_bar.revenue + fnb.garden.revenue +
                   fnb.poolbar_g.revenue + fnb.beach_bar.revenue + houseUseNeg) * 100) / 100,
  };

  console.log('\n F&B Summary:');
  console.log(`   Flower Restaurant : ${fnbValues.B}`);
  console.log(`   Pool Bar          : ${fnbValues.C}`);
  console.log(`   Brutal Garden     : ${fnbValues.D}`);
  console.log(`   Pool Bar Garden   : ${fnbValues.E}`);
  console.log(`   Beach Bar         : ${fnbValues.F}`);
  console.log(`   House Use         : ${fnbValues.G}`);
  console.log(`   ─────────────────────`);
  console.log(`   TOTAL             : ${fnbValues.H}`);

  // ── Parse Hotel ───────────────────────────────────────────────
  console.log('\n' + '─'.repeat(62));
  console.log(' Parsing Hotel...');
  let hotelValues = { occupancyPct: 0, nightsOccupied: 0, nightsAvailable: TOTAL_ROOMS, revenue: 0 };
  if (collected.hotel) {
    try {
      hotelValues = parseHotelBuffer(collected.hotel.buffer, DATE);
    } catch (e) {
      console.error(`  ERROR [hotel]: ${e.message}`);
    }
  } else {
    console.log('  Hotel file not provided — using 0');
  }

  console.log('\n Hotel Summary:');
  console.log(`   Nights Occupied  : ${hotelValues.nightsOccupied} / ${hotelValues.nightsAvailable}`);
  console.log(`   Occupancy %      : ${hotelValues.occupancyPct}%`);
  console.log(`   Revenue          : ${hotelValues.revenue}`);

  // ── Write to XLSX ─────────────────────────────────────────────
  console.log('\n' + '─'.repeat(62));
  console.log(' Writing to Google Sheet...');
  await patchXlsx(drive, DATE, fnbValues, hotelValues);

  // ── Mark emails as processed ──────────────────────────────────
  console.log('\nMarking emails as processed...');
  for (const msgId of processedMsgIds) {
    try {
      await markProcessed(gmail, msgId);
      console.log(`  ✓ ${msgId}`);
    } catch (e) {
      console.log(`  (could not mark ${msgId}: ${e.message})`);
    }
  }

  console.log('\n' + '═'.repeat(62));
  console.log(' ✅ Done! Data written to Google Sheet.');
  console.log(' The FLOW dashboard will reflect the new data immediately.');
  console.log('═'.repeat(62) + '\n');
}

// ── ENTRY POINT ───────────────────────────────────────────────────────────────

const args = process.argv.slice(2);
const watchIdx = args.findIndex(a => a === '--watch');

if (watchIdx !== -1) {
  // Watch mode: run repeatedly every N minutes
  const intervalMin = parseInt(args[watchIdx + 1]) || 5;
  const intervalMs  = intervalMin * 60 * 1000;

  console.log('\n' + '═'.repeat(62));
  console.log(` 👁  Watch mode — checking every ${intervalMin} minute(s)`);
  console.log(' Send "Raporti Ditor" emails to flowreport26@gmail.com');
  console.log(' Press Ctrl+C to stop.');
  console.log('═'.repeat(62) + '\n');

  async function loop() {
    const now = new Date().toLocaleTimeString('en-GB');
    console.log(`[${now}] Checking for new reports...`);
    try {
      await main();
    } catch (e) {
      console.error(`[${now}] Error:`, e.message);
    }
    console.log(`[${new Date().toLocaleTimeString('en-GB')}] Next check in ${intervalMin} min. Waiting...\n`);
    setTimeout(loop, intervalMs);
  }

  loop();

} else {
  // Single run
  main().catch(e => { console.error('\nFATAL:', e.message, '\n', e.stack); process.exit(1); });
}
