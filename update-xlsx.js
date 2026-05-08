/**
 * update-xlsx.js
 * Downloads the Flower Hotel XLSX from Google Drive,
 * writes the Trinosoft data into the correct date row,
 * and re-uploads without touching anything else.
 *
 * First run: opens browser for one-time Google authorization.
 * After that: fully automatic, no browser needed.
 *
 * Usage: node update-xlsx.js <GAS_URL> [YYYY-MM-DD]
 *   or:  node update-xlsx.js --data '{"date":"2026-05-08","fnb":{...},"hotel":{...}}'
 */

const { google }  = require('googleapis');
const XLSX        = require('xlsx');
const fs          = require('fs');
const path        = require('path');
const http        = require('http');
const https       = require('https');

// ── CONFIG ────────────────────────────────────────────────────────────────────
const XLSX_FILE_ID   = '1abLRrgklWeV3wx-KEmA0u4SgCH5ebw3s';
const CREDENTIALS    = path.join(__dirname, 'credentials.json');
const TOKEN_PATH     = path.join(__dirname, 'token.json');
const SCOPES         = ['https://www.googleapis.com/auth/drive'];

// Tab names exactly as they appear in the XLSX
const TAB_FNB        = 'DAILY F&B REVENUES';
const TAB_HOTEL      = 'HOTEL DAILY PERFORMANCE';

// ── AUTH ──────────────────────────────────────────────────────────────────────

const REDIRECT_PORT = 3456;
const REDIRECT_URI  = `http://localhost:${REDIRECT_PORT}`;

async function authorize() {
  if (!fs.existsSync(CREDENTIALS)) {
    console.error('\nERROR: credentials.json not found in: ' + CREDENTIALS);
    process.exit(1);
  }

  const creds  = JSON.parse(fs.readFileSync(CREDENTIALS));
  const { client_id, client_secret } = creds.installed || creds.web;
  const auth   = new google.auth.OAuth2(client_id, client_secret, REDIRECT_URI);

  // Use cached token if available
  if (fs.existsSync(TOKEN_PATH)) {
    auth.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
    return auth;
  }

  // First time: start local server, open browser
  const authUrl = auth.generateAuthUrl({ access_type: 'offline', scope: SCOPES });

  const code = await new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      const qs = new URL(req.url, REDIRECT_URI).searchParams;
      const c  = qs.get('code');
      if (!c) { res.end('No code received'); return; }
      res.end('<h2 style="font-family:sans-serif;color:green">Authorization successful! You can close this tab.</h2>');
      server.close();
      resolve(c);
    });
    server.on('error', reject);
    server.listen(REDIRECT_PORT, () => {
      console.log('\n' + '='.repeat(60));
      console.log(' FIRST-TIME SETUP');
      console.log(' Opening browser for Google authorization...');
      console.log(' Sign in with: flowreport26@gmail.com');
      console.log('='.repeat(60) + '\n');
      const { exec } = require('child_process');
      exec('start "" "' + authUrl + '"');
      console.log('If browser did not open, visit:\n' + authUrl + '\n');
    });
  });

  const { tokens } = await auth.getToken(code);
  auth.setCredentials(tokens);
  fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
  console.log('Token saved. Authorization complete.\n');
  return auth;
}

// ── DRIVE HELPERS ─────────────────────────────────────────────────────────────

async function downloadFile(drive, fileId) {
  const res = await drive.files.get(
    { fileId, alt: 'media' },
    { responseType: 'arraybuffer' }
  );
  return Buffer.from(res.data);
}

async function uploadFile(drive, fileId, buffer) {
  const { Readable } = require('stream');
  await drive.files.update({
    fileId,
    media: {
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      body:     Readable.from(buffer)
    }
  });
}

// ── XLSX HELPERS ──────────────────────────────────────────────────────────────

function formatDate(isoDate) {
  const d = new Date(isoDate + 'T12:00:00');
  const days   = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const months = ['January','February','March','April','May','June','July','August',
                  'September','October','November','December'];
  return days[d.getDay()] + ', ' + months[d.getMonth()] + ' ' +
         String(d.getDate()).padStart(2,'0') + ', ' + d.getFullYear();
}

function parseNum(v) {
  if (v === '' || v == null) return 0;
  const n = parseFloat(String(v).replace(/,/g,''));
  return isNaN(n) ? 0 : n;
}

/**
 * Write values into the correct date row of a sheet.
 * Finds the row by date string in column A.
 * Only overwrites columns B onwards — date and any other columns stay untouched.
 */
function cellMatchesDate(cellA, isoDate) {
  if (!cellA) return false;
  const v = cellA.v;
  // Date object (cellDates:true converts Excel serials to JS Date)
  if (v instanceof Date) {
    // XLSX dates are stored as UTC but displayed in local time (UTC+3 Albania EEST).
    // Adding 12 hours before extracting the date string ensures correct local-date matching.
    const adjusted = new Date(v.getTime() + 12 * 60 * 60 * 1000);
    return adjusted.toISOString().slice(0, 10) === isoDate;
  }
  // Excel serial number — add 12h to compensate for UTC+3 (Albania EEST)
  if (typeof v === 'number') {
    const d = new Date((v - 25569) * 86400 * 1000 + 12 * 60 * 60 * 1000);
    return d.toISOString().slice(0, 10) === isoDate;
  }
  // String comparison
  const formatted = formatDate(isoDate);
  return String(v).trim() === formatted || String(v).trim() === isoDate;
}

function writeRow(ws, isoDate, values) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

  for (let r = range.s.r; r <= range.e.r; r++) {
    const cellA = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    if (!cellMatchesDate(cellA, isoDate)) continue;

    console.log('  Found date ' + isoDate + ' at row ' + (r + 1));

    // Find a nearby row with data to copy cell format from
    const refRow = r > range.s.r ? r - 1 : r + 1;

    values.forEach((val, i) => {
      const addr    = XLSX.utils.encode_cell({ r, c: i + 1 });       // target cell
      const refAddr = XLSX.utils.encode_cell({ r: refRow, c: i + 1 }); // reference cell
      const refCell = ws[refAddr];

      if (typeof val === 'number') {
        const cell = { t: 'n', v: val };
        // Preserve format string (z) and style (s) from a neighboring row
        if (refCell && refCell.z) cell.z = refCell.z;
        if (refCell && refCell.s) cell.s = refCell.s;
        ws[addr] = cell;
      } else {
        const cell = { t: 's', v: String(val) };
        if (refCell && refCell.s) cell.s = refCell.s;
        ws[addr] = cell;
      }
    });
    return true;
  }

  console.log('  Date ' + isoDate + ' not found in sheet — skipping');
  return false;
}

// ── MAIN ──────────────────────────────────────────────────────────────────────

async function main() {
  const GAS_URL = process.argv[2] || '';
  const DATE    = process.argv[3] || new Date().toISOString().slice(0, 10);

  // These are the computed values from parse-trinosoft.js
  // They are passed in via environment variables set by run-daily.js
  const fnbValues = [
    parseNum(process.env.FNB_RESTORANT   || '0'),  // B: Flower Restaurant
    parseNum(process.env.FNB_POOLBAR     || '0'),  // C: Pool Bar
    parseNum(process.env.FNB_GARDEN      || '0'),  // D: Brutal Garden
    parseNum(process.env.FNB_POOLBARG    || '0'),  // E: Pool Bar Garden
    parseNum(process.env.FNB_BEACHBAR    || '0'),  // F: Beach Bar
    parseNum(process.env.FNB_HOUSEUSE    || '0'),  // G: House Use
    parseNum(process.env.FNB_TOTAL       || '0'),  // H: Total
  ];

  const hotelValues = [
    process.env.HOTEL_COMPANY  || 'FLOWER HOTEL & SPA',   // B: Company
    process.env.HOTEL_OCC      || '0%',                    // C: Occupancy%
    parseNum(process.env.HOTEL_NIGHTS    || '0'),          // D: Nights Occupied
    parseNum(process.env.HOTEL_AVAIL     || '110'),        // E: Nights Available
    parseNum(process.env.HOTEL_REVENUE   || '0'),          // F: Revenue
  ];

  console.log('\n' + '='.repeat(60));
  console.log(' Flower Hotel XLSX Updater');
  console.log(' Date: ' + DATE);
  console.log('='.repeat(60));
  console.log('\n F&B:   ' + fnbValues.join(' | '));
  console.log(' Hotel: ' + hotelValues.join(' | '));

  const auth  = await authorize();
  const drive = google.drive({ version: 'v3', auth });

  console.log('\nDownloading XLSX from Google Drive...');
  const buffer = await downloadFile(drive, XLSX_FILE_ID);
  console.log('  Downloaded: ' + (buffer.length / 1024).toFixed(0) + ' KB');

  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false, cellStyles: true });
  console.log('  Sheets: ' + wb.SheetNames.join(', '));

  const wsFnB   = wb.Sheets[TAB_FNB];
  const wsHotel = wb.Sheets[TAB_HOTEL];

  if (!wsFnB)   { console.error('  Tab "' + TAB_FNB + '" not found!'); }
  if (!wsHotel) { console.error('  Tab "' + TAB_HOTEL + '" not found!'); }

  console.log('\n[1/2] DAILY F&B REVENUES...');
  if (wsFnB)   writeRow(wsFnB,   DATE, fnbValues);

  console.log('\n[2/2] HOTEL DAILY PERFORMANCE...');
  if (wsHotel) writeRow(wsHotel, DATE, hotelValues);

  console.log('\nConverting and uploading...');
  const out = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', compression: true });
  await uploadFile(drive, XLSX_FILE_ID, out);

  console.log('\n✅ Done! Both rows updated in your Google Sheet.\n');
}

main().catch(e => { console.error('\nFATAL:', e.message); process.exit(1); });
