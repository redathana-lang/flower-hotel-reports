'use strict';
/**
 * reprocess-dates.js
 * Heq etiketën "processed-report" nga emailet e datave të specifikuara
 * dhe riproceson me datën e saktë.
 *
 * Përdorim:
 *   node reprocess-dates.js 2026-05-09
 *   node reprocess-dates.js 2026-05-10
 */

const { google }   = require('googleapis');
const XLSX         = require('xlsx');
const JSZip        = require('jszip');
const fs           = require('fs');
const path         = require('path');
const { Readable } = require('stream');

const XLSX_FILE_ID = '1abLRrgklWeV3wx-KEmA0u4SgCH5ebw3s';
const CREDENTIALS  = path.join(__dirname, 'credentials.json');
const TOKEN_PATH   = path.join(__dirname, 'token.json');
const TOTAL_ROOMS  = 110;

const EXCLUDE_NAMES     = ['ernest caci', 'olti caci', 'olti  caci', 'ahmet caci', 'jasht pune', 'jashte pune', 'bllok'];
// Block/agency CLIENT names also excluded for reports dated 2026-06-06 onward (client name only).
const EXCLUDE_NAMES_NEW = ['itaka', 'saistours', 'w2m'];
const NEW_EXCL_FROM_DATE = '2026-06-06';
const HOUSE_USE_PATTERNS = ['vila 1', 'vila2', 'vila 2', 'dhoma 313', 'fature qerasje'];

async function authorize() {
  const creds = JSON.parse(fs.readFileSync(CREDENTIALS));
  const { client_id, client_secret } = creds.installed || creds.web;
  const auth = new google.auth.OAuth2(client_id, client_secret, 'http://localhost:3456');
  auth.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return auth;
}

function parseNum(v) {
  if (v === '' || v == null) return 0;
  const n = parseFloat(String(v).replace(/\s/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
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
    if (isHouseUse(tableName)) houseUse += net;
    else revenue += net;
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

function classifyFile(filename) {
  const n = filename.toLowerCase();
  if (n.includes('prenotim') || n.includes('recepsion') || n.includes('reception') || n.includes('hotel')) return 'hotel';
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
    for (const row of rows)
      for (const cell of row)
        if (String(cell).toLowerCase().includes('garden')) return 'poolbar_g';
  } catch (e) {}
  return 'pool_bar';
}

function findRowForDate(xml, isoDate) {
  const d      = new Date(isoDate + 'T00:00:00Z');
  const serial = Math.floor(d.getTime() / 86400000) + 25569;
  const rowRe  = /<row r="(\d+)"[^>]*>[\s\S]*?<\/row>/g;
  let m;
  while ((m = rowRe.exec(xml)) !== null) {
    const rowNum = m[1];
    const rowXml = m[0];
    if (new RegExp(`<c r="A${rowNum}"[^>]*><v>${serial}(?:\\.0)?<\\/v><\\/c>`).test(rowXml))
      return { rowNum: parseInt(rowNum), serial };
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
    styles[col] = cellMatch
      ? { s: getCellStyle(cellMatch[0]), t: /\bt="([^"]+)"/.exec(cellMatch[0])?.[1] || '' }
      : { s: '', t: '' };
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

  // FNB (sheet3.xml)
  let fnbXml  = await zip.file('xl/worksheets/sheet3.xml').async('string');
  const fnbInfo = findRowForDate(fnbXml, isoDate);
  if (!fnbInfo) {
    console.error(`  FNB: data ${isoDate} nuk u gjet!`);
  } else {
    const r = fnbInfo.rowNum, ser = fnbInfo.serial;
    console.log(`  FNB: ${isoDate} → rreshti ${r} (serial ${ser})`);
    const ref = readRefStyles(fnbXml, r - 1, ['A','B','C','D','E','F','G','H']);
    const sA = ref?.A?.s || '23', sB = ref?.B?.s || '36';
    const sG = ref?.G?.s || '34', sH = ref?.H?.s || '35';
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
    console.log(`  ✓ FNB rreshti ${r} u shkrua`);
  }

  // Hotel (sheet2.xml)
  let hotelXml = await zip.file('xl/worksheets/sheet2.xml').async('string');
  const hotelInfo = findRowForDate(hotelXml, isoDate);
  if (!hotelInfo) {
    console.error(`  Hotel: data ${isoDate} nuk u gjet!`);
  } else {
    const r = hotelInfo.rowNum, ser = hotelInfo.serial;
    console.log(`  Hotel: ${isoDate} → rreshti ${r} (serial ${ser})`);
    const ref  = readRefStyles(hotelXml, r - 1, ['A','B','C','D','E','F']);
    const sA   = ref?.A?.s || '23', sB = ref?.B?.s || '17';
    const sBt  = ref?.B?.t || 's', sC = ref?.C?.s || '24';
    const sD   = ref?.D?.s || '27', sE = ref?.E?.s || '17', sF = ref?.F?.s || '28';
    const refRow  = hotelXml.match(new RegExp(`<row r="${r-1}"[^>]*>[\\s\\S]*?<\\/row>`))?.[0] || '';
    const ssMatch = refRow.match(new RegExp(`<c r="B${r-1}"[^>]*><v>(\\d+)<\\/v><\\/c>`));
    const ssIdx   = ssMatch ? ssMatch[1] : '35';
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
    hotelXml = hotelXml.replace(new RegExp(`<row r="${r}"[^>]*>[\\s\\S]*?<\\/row>`), newRow);
    zip.file('xl/worksheets/sheet2.xml', hotelXml);
    console.log(`  ✓ Hotel rreshti ${r} u shkrua`);
  }

  const outBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  console.log(`\n  Po ngarkohet ${(outBuf.length / 1024).toFixed(0)} KB...`);
  await drive.files.update({
    fileId: XLSX_FILE_ID,
    media: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', body: Readable.from(outBuf) },
  });
  console.log('  ✓ Ngarkuar');
}

async function getAttachments(gmail, messageId) {
  const msg = await gmail.users.messages.get({ userId: 'me', id: messageId });
  const attachments = [];
  function findParts(parts) {
    if (!parts) return;
    for (const part of parts) {
      if (part.parts) findParts(part.parts);
      const fname = part.filename || '';
      if (/\.(xls|xlsx)$/i.test(fname) && part.body)
        attachments.push({ filename: fname, part });
    }
  }
  findParts(msg.data.payload.parts);
  const results = [];
  for (const { filename, part } of attachments) {
    let data;
    if (part.body.data) {
      data = Buffer.from(part.body.data, 'base64');
    } else if (part.body.attachmentId) {
      const att = await gmail.users.messages.attachments.get({ userId: 'me', messageId, id: part.body.attachmentId });
      data = Buffer.from(att.data.data, 'base64');
    }
    if (data) {
      results.push({ filename, buffer: data });
      console.log(`    📎 ${filename} (${(data.length / 1024).toFixed(0)} KB)`);
    }
  }
  return results;
}

async function main() {
  const dateArg = process.argv.slice(2).find(a => /^\d{4}-\d{2}-\d{2}$/.test(a));
  if (!dateArg) {
    console.error('Përdorim: node reprocess-dates.js YYYY-MM-DD');
    process.exit(1);
  }
  const DATE = dateArg;

  console.log('\n' + '═'.repeat(62));
  console.log(' Riprocesim i datës: ' + DATE);
  console.log('═'.repeat(62));

  const auth  = await authorize();
  const gmail = google.gmail({ version: 'v1', auth });
  const drive = google.drive({ version: 'v3', auth });

  // Gjej etiketën "processed-report"
  const labelsRes = await gmail.users.labels.list({ userId: 'me' });
  const procLabel = (labelsRes.data.labels || []).find(l => l.name === 'processed-report');
  if (!procLabel) { console.log('Etiketa "processed-report" nuk ekziston.'); process.exit(1); }
  console.log(`Etiketa "processed-report": ${procLabel.id}`);

  // Kërko emailet e datës (me ose pa etiketë)
  const d      = new Date(DATE + 'T00:00:00Z');
  const after  = DATE.replace(/-/g, '/');
  const nextD  = new Date(d.getTime() + 86400000);
  const before = nextD.toISOString().slice(0,10).replace(/-/g, '/');
  const query  = `subject:"Raporti Ditor" has:attachment after:${after} before:${before}`;
  console.log(`\nKërkesë: ${query}`);

  const r1 = await gmail.users.messages.list({ userId: 'me', q: query, maxResults: 20 });
  const msgIds = (r1.data.messages || []).map(m => m.id);
  if (msgIds.length === 0) {
    console.log('Asnjë email i gjetur për këtë datë.');
    process.exit(0);
  }
  console.log(`Gjendur ${msgIds.length} email(e)\n`);

  // Hiq etiketën "processed-report" që t'i rilexojmë
  for (const msgId of msgIds) {
    await gmail.users.messages.modify({
      userId: 'me', id: msgId,
      requestBody: { removeLabelIds: [procLabel.id] }
    });
    console.log(`  ✓ Hequr etiketa nga ${msgId}`);
  }

  // Mbledh attachmentet
  const collected = { restorant: null, pool_bar: null, poolbar_g: null, garden: null, hotel: null };
  const processedMsgIds = [];

  for (const msgId of msgIds) {
    console.log(`\nProcessing ${msgId}...`);
    const attachments = await getAttachments(gmail, msgId);
    let hasReport = false;

    for (const { filename, buffer } of attachments) {
      let type = classifyFile(filename);
      if (type === 'pool_bar' && collected['pool_bar']) {
        const ct = classifyPoolBarByContent(buffer);
        type = ct;
        if (ct === 'pool_bar') {
          const et = classifyPoolBarByContent(collected['pool_bar'].buffer);
          if (et === 'poolbar_g') {
            collected['poolbar_g'] = collected['pool_bar'];
            collected['pool_bar']  = { filename, buffer };
            hasReport = true; continue;
          }
        }
      }
      if (type && !collected[type]) {
        collected[type] = { filename, buffer };
        console.log(`    → ${type}: ${filename}`);
        hasReport = true;
      } else if (type && collected[type]) {
        console.log(`    → ${type} tashmë i mbledhur, anashkaluar`);
      } else {
        console.log(`    → Nuk u klasifikua: ${filename}`);
      }
    }
    if (hasReport) processedMsgIds.push(msgId);
  }

  console.log('\n Skedarët e mbledhur:');
  for (const [k, v] of Object.entries(collected))
    console.log(`  ${k.padEnd(12)}: ${v ? '✓ ' + v.filename : '✗ mungon'}`);

  // Parse F&B
  console.log('\n Parse F&B...');
  const fnb = { restorant: { revenue:0, houseUse:0 }, pool_bar: { revenue:0, houseUse:0 }, poolbar_g: { revenue:0, houseUse:0 }, garden: { revenue:0, houseUse:0 } };
  for (const key of Object.keys(fnb)) {
    if (collected[key]) {
      try { fnb[key] = parseFnBBuffer(collected[key].buffer, key); }
      catch (e) { console.error(`  GABIM [${key}]: ${e.message}`); }
    } else {
      console.log(`  [${key}] mungon — vlera 0`);
    }
  }
  const totalHouseUse = Object.values(fnb).reduce((s,v) => s + v.houseUse, 0);
  const houseUseNeg   = -(Math.round(totalHouseUse * 100) / 100);
  const fnbValues = {
    B: fnb.restorant.revenue, C: fnb.pool_bar.revenue,
    D: fnb.garden.revenue,    E: fnb.poolbar_g.revenue,
    F: 0, G: houseUseNeg,
    H: Math.round((fnb.restorant.revenue + fnb.pool_bar.revenue + fnb.garden.revenue + fnb.poolbar_g.revenue + houseUseNeg) * 100) / 100,
  };
  console.log(`\n F&B: R=${fnbValues.B}  PB=${fnbValues.C}  G=${fnbValues.D}  PBG=${fnbValues.E}  HU=${fnbValues.G}  TOTAL=${fnbValues.H}`);

  // Parse Hotel
  console.log('\n Parse Hotel...');
  let hotelValues = { occupancyPct:0, nightsOccupied:0, nightsAvailable:TOTAL_ROOMS, revenue:0 };
  if (collected.hotel) {
    try { hotelValues = parseHotelBuffer(collected.hotel.buffer, DATE); }
    catch (e) { console.error(`  GABIM [hotel]: ${e.message}`); }
  } else {
    console.log('  Hotel skedar mungon — vlera 0');
  }
  console.log(` Hotel: ${hotelValues.nightsOccupied}/${TOTAL_ROOMS} = ${hotelValues.occupancyPct}%  Rev: ${hotelValues.revenue}`);

  // Shkruaj në XLSX
  console.log('\n Po shkruhet në Google Sheet...');
  await patchXlsx(drive, DATE, fnbValues, hotelValues);

  // Rishëno si të procesuar
  console.log('\nRishënim si të procesuar...');
  for (const msgId of processedMsgIds) {
    try {
      await gmail.users.messages.modify({
        userId: 'me', id: msgId,
        requestBody: { removeLabelIds: ['UNREAD'], addLabelIds: [procLabel.id] }
      });
      console.log(`  ✓ ${msgId}`);
    } catch (e) { console.log(`  (nuk u shënua ${msgId}: ${e.message})`); }
  }

  console.log('\n' + '═'.repeat(62));
  console.log(` ✅ Gati! Të dhënat e ${DATE} janë shkruar saktë.`);
  console.log('═'.repeat(62) + '\n');
}

main().catch(e => { console.error('\nGABIM FATAL:', e.message, '\n', e.stack); process.exit(1); });
