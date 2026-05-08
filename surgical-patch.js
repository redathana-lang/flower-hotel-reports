/**
 * surgical-patch.js  (v2)
 * Patches ONLY row 860 (Excel row for 2026-05-08) in the XLSX file by
 * doing a zip-level XML edit — never re-serializes the whole workbook,
 * so Google Sheets GID mappings remain intact.
 *
 * Also clears any bad data from row 861 (May 9) that was accidentally
 * written during a previous run.
 *
 * Usage: node surgical-patch.js
 */

const { google } = require('googleapis');
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

// ── DATA FOR 2026-05-08 ───────────────────────────────────────────────────────
// Row 860 = Excel row for May 8, 2026 (header is row 1, Jan 1 2024 is row 2, ...)
const FNB_ROW = 860;
const FNB_DATA = {
  B: 38600,       // Flower Restaurant
  C: 5360,        // Pool Bar
  D: 210600,      // Brutal Garden
  E: 17380,       // Pool Bar Garden
  F: 0,           // Beach Bar
  G: -36930,      // House Use
  H: 235010,      // Total
};

const HOTEL_ROW = 860;
const HOTEL_DATA = {
  B: 'Flower Hotel & Spa',  // Company
  C: '35.45%',              // Occupancy %
  D: 39,                    // Nights Occupied
  E: 110,                   // Nights Available
  F: 4631.33,               // Revenue
};

// May 9 row that got polluted — needs to be cleaned
const BAD_ROW = 861;

// Tab names (exactly as in workbook.xml, decoded)
const TAB_FNB   = 'DAILY F&B REVENUES';
const TAB_HOTEL = 'HOTEL DAILY PERFORMANCE';

// ── AUTH ──────────────────────────────────────────────────────────────────────
async function authorize() {
  if (!fs.existsSync(CREDENTIALS)) {
    console.error('ERROR: credentials.json not found');
    process.exit(1);
  }
  const creds = JSON.parse(fs.readFileSync(CREDENTIALS));
  const { client_id, client_secret } = creds.installed || creds.web;
  const auth = new google.auth.OAuth2(client_id, client_secret, REDIRECT_URI);
  if (fs.existsSync(TOKEN_PATH)) {
    auth.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
    return auth;
  }
  const authUrl = auth.generateAuthUrl({ access_type: 'offline', scope: ['https://www.googleapis.com/auth/drive'] });
  const code = await new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      const c = new URL(req.url, REDIRECT_URI).searchParams.get('code');
      if (!c) { res.end('No code'); return; }
      res.end('<h2 style="color:green">Done! Close this tab.</h2>');
      server.close();
      resolve(c);
    });
    server.on('error', reject);
    server.listen(REDIRECT_PORT, () => {
      const { exec } = require('child_process');
      exec('start "" "' + authUrl + '"');
    });
  });
  const { tokens } = await auth.getToken(code);
  auth.setCredentials(tokens);
  fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
  return auth;
}

// ── DRIVE HELPERS ─────────────────────────────────────────────────────────────
async function downloadFile(drive, fileId) {
  const res = await drive.files.get({ fileId, alt: 'media' }, { responseType: 'arraybuffer' });
  return Buffer.from(res.data);
}

async function uploadFile(drive, fileId, buffer) {
  await drive.files.update({
    fileId,
    media: {
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      body: Readable.from(buffer)
    }
  });
}

// ── XML HELPERS ───────────────────────────────────────────────────────────────

function decodeXmlEntities(str) {
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function encodeXmlEntities(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function parseWorkbookSheets(workbookXml, relsXml) {
  const sheetMap = {};
  const sheetRegex = /<sheet[^>]+name="([^"]+)"[^>]+r:id="([^"]+)"/g;
  let m;
  while ((m = sheetRegex.exec(workbookXml)) !== null) {
    sheetMap[m[2]] = decodeXmlEntities(m[1]);
  }

  const relMap = {};
  const relRegex = /<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"/g;
  while ((m = relRegex.exec(relsXml)) !== null) {
    relMap[m[1]] = m[2];
  }

  const result = {};
  for (const [rId, name] of Object.entries(sheetMap)) {
    const target = relMap[rId];
    if (target) {
      result[name] = path.basename(target);
      console.log(`  Sheet: "${name}" → ${path.basename(target)}`);
    }
  }
  return result;
}

/**
 * Extract the row opening tag (attributes), the A-column cell XML, and any
 * style/span info we want to preserve from an existing row.
 */
function extractRowInfo(xmlStr, rowNum) {
  // Match <row r="860" ...> ... </row>
  const rowOpenRe = new RegExp(`(<row\\b[^>]*\\br="${rowNum}"[^>]*>)`, '');
  const rowOpenMatch = xmlStr.match(rowOpenRe);
  if (!rowOpenMatch) return null;

  const openTag = rowOpenMatch[1];
  const rowStart = xmlStr.indexOf(openTag);
  const rowEnd = xmlStr.indexOf('</row>', rowStart);
  const rowXml = xmlStr.slice(rowStart, rowEnd + 6);

  // Extract the A-column cell (keep it verbatim — it holds the date)
  const aCellMatch = rowXml.match(/<c r="A\d+"[^>]*>[\s\S]*?<\/c>/);
  const aCell = aCellMatch ? aCellMatch[0] : '';

  return { openTag, aCell, rowXml };
}

/**
 * Replace an entire row in the worksheet XML with a rebuilt version.
 * Keeps the row's opening-tag attributes and the A-column date cell.
 * Writes all colData entries in correct column order (B, C, D, ...).
 *
 * For "clean" rows (BAD_ROW), pass colData = null to remove all non-A cells.
 */
function replaceRow(xmlStr, rowNum, colData) {
  const info = extractRowInfo(xmlStr, rowNum);
  if (!info) {
    console.log(`  Row ${rowNum} not found in XML — skipping`);
    return xmlStr;
  }

  console.log(`  Row ${rowNum} found. Open tag: ${info.openTag.slice(0, 80)}`);

  // Build the new row content
  let newRow = info.openTag;

  // Keep the A cell (date)
  if (info.aCell) {
    newRow += info.aCell;
    console.log(`    A${rowNum}: preserved (date cell)`);
  }

  // Add data cells in column order (B, C, D, …)
  if (colData) {
    for (const [col, val] of Object.entries(colData)) {
      const cellRef = `${col}${rowNum}`;
      let cellXml;
      if (typeof val === 'string') {
        cellXml = `<c r="${cellRef}" t="str"><v>${encodeXmlEntities(val)}</v></c>`;
      } else {
        cellXml = `<c r="${cellRef}"><v>${val}</v></c>`;
      }
      newRow += cellXml;
      console.log(`    ${cellRef}: ${JSON.stringify(val)}`);
    }
  } else {
    console.log(`    (cleared — no data cells)`);
  }

  newRow += '</row>';

  // Replace old row XML with new row XML
  const before = xmlStr.slice(0, xmlStr.indexOf(info.rowXml));
  const after  = xmlStr.slice(xmlStr.indexOf(info.rowXml) + info.rowXml.length);
  return before + newRow + after;
}

// ── MAIN ──────────────────────────────────────────────────────────────────────
async function main() {
  console.log('\n' + '='.repeat(60));
  console.log(' Flower Hotel XLSX Surgical Patcher v2');
  console.log(' Target: row 860 (2026-05-08)');
  console.log(' Cleanup: row 861 (accidentally patched before)');
  console.log('='.repeat(60));

  const auth  = await authorize();
  const drive = google.drive({ version: 'v3', auth });

  console.log('\nDownloading XLSX from Google Drive...');
  const buffer = await downloadFile(drive, XLSX_FILE_ID);
  console.log(`  Downloaded: ${(buffer.length / 1024).toFixed(0)} KB`);

  const zip = await JSZip.loadAsync(buffer);

  // Parse workbook structure
  const wbXml  = await zip.file('xl/workbook.xml').async('string');
  const wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  console.log('\n  Sheet mapping:');
  const sheetMap = parseWorkbookSheets(wbXml, wbRels);

  const fnbFile   = sheetMap[TAB_FNB];
  const hotelFile = sheetMap[TAB_HOTEL];
  if (!fnbFile)   { console.error(`  ERROR: Tab "${TAB_FNB}" not found`);   }
  if (!hotelFile) { console.error(`  ERROR: Tab "${TAB_HOTEL}" not found`); }

  // ── FNB sheet ─────────────────────────────────────────────────
  if (fnbFile) {
    const fnbPath = `xl/worksheets/${fnbFile}`;
    console.log(`\n[FNB] ${fnbPath}`);
    let xml = await zip.file(fnbPath).async('string');

    console.log(`\n  [1] Writing data to row ${FNB_ROW} (May 8, 2026):`);
    xml = replaceRow(xml, FNB_ROW, FNB_DATA);

    console.log(`\n  [2] Cleaning bad data from row ${BAD_ROW} (May 9):`);
    xml = replaceRow(xml, BAD_ROW, null);

    zip.file(fnbPath, xml);
    console.log(`  ✓ FNB sheet updated`);
  }

  // ── Hotel sheet ───────────────────────────────────────────────
  if (hotelFile) {
    const hotelPath = `xl/worksheets/${hotelFile}`;
    console.log(`\n[Hotel] ${hotelPath}`);
    let xml = await zip.file(hotelPath).async('string');

    console.log(`\n  [1] Writing data to row ${HOTEL_ROW} (May 8, 2026):`);
    xml = replaceRow(xml, HOTEL_ROW, HOTEL_DATA);

    console.log(`\n  [2] Cleaning bad data from row ${BAD_ROW} (May 9):`);
    xml = replaceRow(xml, BAD_ROW, null);

    zip.file(hotelPath, xml);
    console.log(`  ✓ Hotel sheet updated`);
  }

  // Generate and upload
  const outBuffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });
  console.log(`\n  Output size: ${(outBuffer.length / 1024).toFixed(0)} KB`);

  console.log('\nUploading to Google Drive...');
  await uploadFile(drive, XLSX_FILE_ID, outBuffer);
  console.log('✅ Done! Surgical patch v2 complete.\n');

  // ── Verify ────────────────────────────────────────────────────
  console.log('Verifying via CSV export...');
  const https = require('https');
  function fetchFollow(url, maxR=5) {
    return new Promise((resolve, reject) => {
      https.get(url, res => {
        if ([301,302,307].includes(res.statusCode) && res.headers.location && maxR > 0) {
          res.resume();
          fetchFollow(res.headers.location, maxR-1).then(resolve, reject);
          return;
        }
        let d = '';
        res.on('data', c => d += c);
        res.on('end', () => resolve({ status: res.statusCode, body: d }));
      }).on('error', reject);
    });
  }

  const base = `https://docs.google.com/spreadsheets/d/${XLSX_FILE_ID}/export?format=csv&gid=`;

  const fnbR = await fetchFollow(base + '663170393');
  const fnbLines = fnbR.body.split('\n');
  console.log(`\nFNB CSV (GID 663170393): HTTP ${fnbR.status}`);
  // Print lines 858-863
  for (let i = 857; i <= 862; i++) {
    if (fnbLines[i]) console.log(`  line ${i+1}: ${fnbLines[i].slice(0,160)}`);
  }

  const hotelR = await fetchFollow(base + '398660926');
  const hotelLines = hotelR.body.split('\n');
  console.log(`\nHotel CSV (GID 398660926): HTTP ${hotelR.status}`);
  for (let i = 857; i <= 862; i++) {
    if (hotelLines[i]) console.log(`  line ${i+1}: ${hotelLines[i].slice(0,160)}`);
  }
}

main().catch(e => { console.error('\nFATAL:', e.message, '\n', e.stack); process.exit(1); });
