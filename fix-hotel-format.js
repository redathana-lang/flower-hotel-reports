/**
 * fix-hotel-format.js
 * Rewrites Hotel row 860 (May 8, 2026) to match the exact XML structure
 * of the surrounding rows — shared-string company, correct style indices,
 * occupancy stored as a decimal formula value, revenue with € format.
 */

const { google } = require('googleapis');
const JSZip      = require('jszip');
const fs         = require('fs');
const path       = require('path');
const { Readable } = require('stream');

const XLSX_FILE_ID = '1abLRrgklWeV3wx-KEmA0u4SgCH5ebw3s';
const CREDENTIALS  = path.join(__dirname, 'credentials.json');
const TOKEN_PATH   = path.join(__dirname, 'token.json');

// ── Values for May 8, 2026 ────────────────────────────────────────────────────
const HOTEL_ROW       = 860;
const DATE_SERIAL     = 46150.0;           // Excel serial for 2026-05-08
const SHARED_STR_IDX  = 35;               // "FLOWER HOTEL & SPA" in sharedStrings
const OCC_DECIMAL     = 39 / 110;         // 0.35454545...  (nights / available)
const NIGHTS_OCC      = 39.0;
const NIGHTS_AVAIL    = 110.0;
const REVENUE         = 4631.33;

// Style indices copied from neighboring rows (857-859, 862)
// A: s=23 (date), B: s=17 (string), C: s=24 (percent formula),
// D: s=27 (integer nights), E: s=17, F: s=28 (currency €)
const CORRECT_ROW_XML =
  `<row r="${HOTEL_ROW}" ht="14.25" customHeight="1">` +
  `<c r="A${HOTEL_ROW}" s="23"><v>${DATE_SERIAL}</v></c>` +
  `<c r="B${HOTEL_ROW}" s="17" t="s"><v>${SHARED_STR_IDX}</v></c>` +
  `<c r="C${HOTEL_ROW}" s="24"><f t="shared" si="1"/><v>${OCC_DECIMAL}</v></c>` +
  `<c r="D${HOTEL_ROW}" s="27"><v>${NIGHTS_OCC}</v></c>` +
  `<c r="E${HOTEL_ROW}" s="17"><v>${NIGHTS_AVAIL}</v></c>` +
  `<c r="F${HOTEL_ROW}" s="28"><v>${REVENUE}</v></c>` +
  `</row>`;

async function authorize() {
  const creds = JSON.parse(fs.readFileSync(CREDENTIALS));
  const { client_id, client_secret } = creds.installed || creds.web;
  const auth = new google.auth.OAuth2(client_id, client_secret, 'http://localhost:3456');
  auth.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return auth;
}

async function main() {
  console.log('\n' + '='.repeat(60));
  console.log(' Hotel Row 860 — Format Fix');
  console.log(' Matching styles of surrounding rows (s=17,24,27,28)');
  console.log('='.repeat(60));

  const auth  = await authorize();
  const drive = google.drive({ version: 'v3', auth });

  console.log('\nDownloading...');
  const res = await drive.files.get(
    { fileId: XLSX_FILE_ID, alt: 'media' },
    { responseType: 'arraybuffer' }
  );
  const buffer = Buffer.from(res.data);
  console.log(`  ${(buffer.length / 1024).toFixed(0)} KB`);

  const zip = await JSZip.loadAsync(buffer);
  let xml = await zip.file('xl/worksheets/sheet2.xml').async('string');

  // Find and replace row 860
  const rowRe = /<row r="860"[^>]*>[\s\S]*?<\/row>/;
  const match = xml.match(rowRe);
  if (!match) {
    console.error('Row 860 not found!');
    process.exit(1);
  }

  console.log('\nCurrent row 860:');
  console.log(' ', match[0].slice(0, 300));
  console.log('\nNew row 860:');
  console.log(' ', CORRECT_ROW_XML);

  xml = xml.replace(rowRe, CORRECT_ROW_XML);
  zip.file('xl/worksheets/sheet2.xml', xml);

  const outBuf = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });
  console.log(`\n  Output: ${(outBuf.length / 1024).toFixed(0)} KB`);

  console.log('\nUploading...');
  await drive.files.update({
    fileId: XLSX_FILE_ID,
    media: {
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      body: Readable.from(outBuf)
    }
  });

  console.log('\n✅ Done!\n');

  // Verify
  const https = require('https');
  function fetchFollow(url, n=5) {
    return new Promise((resolve, reject) => {
      https.get(url, r => {
        if ([301,302,307].includes(r.statusCode) && r.headers.location && n > 0) {
          r.resume(); fetchFollow(r.headers.location, n-1).then(resolve, reject); return;
        }
        let d = ''; r.on('data', c => d += c); r.on('end', () => resolve({ status: r.statusCode, body: d }));
      }).on('error', reject);
    });
  }

  const r = await fetchFollow(`https://docs.google.com/spreadsheets/d/${XLSX_FILE_ID}/export?format=csv&gid=398660926`);
  const lines = r.body.split('\n');
  console.log('Hotel CSV (GID 398660926):');
  for (let i = 857; i <= 862; i++) {
    if (lines[i]) console.log(`  line ${i+1}: ${lines[i].slice(0,160)}`);
  }
}

main().catch(e => { console.error('FATAL:', e.message); process.exit(1); });
