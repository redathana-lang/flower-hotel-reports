/**
 * fix-fnb-format.js
 * Rewrites FNB row 860 (May 8, 2026) to match the exact XML structure
 * of surrounding rows — correct style indices and shared formula for Total.
 *
 * From neighboring rows:
 *   B-F  → s="36"  (number format "15,990.00")
 *   G    → s="34"  (negative house-use format)
 *   H    → s="35" + shared formula (Total = SUM)
 */

const { google } = require('googleapis');
const JSZip      = require('jszip');
const fs         = require('fs');
const path       = require('path');
const { Readable } = require('stream');

const XLSX_FILE_ID = '1abLRrgklWeV3wx-KEmA0u4SgCH5ebw3s';
const CREDENTIALS  = path.join(__dirname, 'credentials.json');
const TOKEN_PATH   = path.join(__dirname, 'token.json');

const FNB_ROW  = 860;
const DATE_SER = 46150.0;

// Values
const B = 38600.0;    // Flower Restaurant
const C = 5360.0;     // Pool Bar
const D = 210600.0;   // Brutal Garden
const E = 17380.0;    // Pool Bar Garden
const F = 0.0;        // Beach Bar
const G = -36930.0;   // House Use
const H = 235010;     // Total

const CORRECT_ROW_XML =
  `<row r="${FNB_ROW}" ht="14.25" customHeight="1">` +
  `<c r="A${FNB_ROW}" s="23"><v>${DATE_SER}</v></c>` +
  `<c r="B${FNB_ROW}" s="36"><v>${B}</v></c>` +
  `<c r="C${FNB_ROW}" s="36"><v>${C}</v></c>` +
  `<c r="D${FNB_ROW}" s="36"><v>${D}</v></c>` +
  `<c r="E${FNB_ROW}" s="36"><v>${E}</v></c>` +
  `<c r="F${FNB_ROW}" s="36"><v>${F}</v></c>` +
  `<c r="G${FNB_ROW}" s="34"><v>${G}</v></c>` +
  `<c r="H${FNB_ROW}" s="35"><f t="shared" si="1"/><v>${H}</v></c>` +
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
  console.log(' FNB Row 860 — Format Fix');
  console.log(' Applying styles s=36 (B-F), s=34 (G), s=35+formula (H)');
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
  let xml = await zip.file('xl/worksheets/sheet3.xml').async('string');

  const rowRe = /<row r="860"[^>]*>[\s\S]*?<\/row>/;
  const match = xml.match(rowRe);
  if (!match) { console.error('Row 860 not found!'); process.exit(1); }

  console.log('\nCurrent row 860:');
  console.log(' ', match[0].slice(0, 400));
  console.log('\nNew row 860:');
  console.log(' ', CORRECT_ROW_XML);

  xml = xml.replace(rowRe, CORRECT_ROW_XML);
  zip.file('xl/worksheets/sheet3.xml', xml);

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

  // Verify via CSV
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

  const r = await fetchFollow(`https://docs.google.com/spreadsheets/d/${XLSX_FILE_ID}/export?format=csv&gid=663170393`);
  const lines = r.body.split('\n');
  console.log('FNB CSV (GID 663170393):');
  for (let i = 857; i <= 862; i++) {
    if (lines[i]) console.log(`  line ${i+1}: ${lines[i].slice(0,160)}`);
  }
}

main().catch(e => { console.error('FATAL:', e.message); process.exit(1); });
