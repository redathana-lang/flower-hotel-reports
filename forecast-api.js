'use strict';
/**
 * forecast-api.js — serves the Flower commercial forecast over HTTP so the
 * Commercial Analyst Agent can produce the weekly brief from anywhere, using
 * whatever was uploaded to the FLOW Dashboard that day.
 *
 * Mount it in gmail-worker.js with two lines:
 *
 *     require('./forecast-api').mount(app, authorize);
 *
 * Then:  GET /api/forecast            -> full JSON for the brief
 *        GET /api/forecast?as_of=YYYY-MM-DD
 *        GET /api/forecast/health     -> cheap check, no Drive download
 *
 * Why this exists: the current-year export is ~15 MB, far too large for an
 * agent session to download and parse. The server does it instead and hands
 * back a few KB of finished numbers. The closed baseline year never needs
 * downloading at all — baseline-2025.json holds it in 63 KB.
 *
 * Everything here is dependency-free except googleapis, which the worker
 * already has.
 */

const fs = require('fs');
const path = require('path');
const B = require('./lib/bookings');
const { forecastFrom } = require('./lib/forecast-core');

// Drive folder the FLOW Dashboard writes one export into per day.
const EXPORTS_FOLDER_ID = '12Gbgi50L6mLIddz2-LzR8bkEhtXEoytq';
const BASELINE = JSON.parse(fs.readFileSync(path.join(__dirname, 'baseline-2025.json'), 'utf8'));

// One download per day is plenty; the export only changes when a new one lands.
const cache = { key: null, at: 0, payload: null };
const CACHE_MS = 30 * 60 * 1000;

/** Newest Prenotimet_<date>.xls in the exports folder. */
async function findNewestExport(drive) {
  const res = await drive.files.list({
    q: `'${EXPORTS_FOLDER_ID}' in parents and trashed = false and name contains 'Prenotimet'`,
    orderBy: 'createdTime desc',
    pageSize: 5,
    fields: 'files(id,name,size,createdTime,modifiedTime)',
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
  });
  const files = (res.data && res.data.files) || [];
  if (!files.length) throw new Error(`No Prenotimet export found in Drive folder ${EXPORTS_FOLDER_ID}`);
  return files[0];
}

/** Download a Drive file to a temp path and return it. */
async function downloadToTmp(drive, file) {
  const dest = path.join(require('os').tmpdir(), `flower-${file.id}.xls`);
  if (fs.existsSync(dest) && fs.statSync(dest).size > 0) return dest;
  const res = await drive.files.get(
    { fileId: file.id, alt: 'media', supportsAllDrives: true },
    { responseType: 'stream' },
  );
  await new Promise((resolve, reject) => {
    const out = fs.createWriteStream(dest);
    res.data.on('error', reject);
    out.on('error', reject);
    out.on('finish', resolve);
    res.data.pipe(out);
  });
  return dest;
}

/**
 * Build the payload. `asOf` defaults to the export's own date, which is what
 * the brief wants: the picture as at the day the data was uploaded.
 */
function computePayload(exportPath, exportFile, asOfISO) {
  const loaded = B.load(exportPath);
  return forecastFrom({
    bookings: loaded.bookings,
    stats: loaded.stats,
    baseline: BASELINE,
    asOf: asOfISO,
    sourceFile: exportFile ? exportFile.name : path.basename(exportPath),
    sourceCreatedTime: exportFile ? exportFile.createdTime : null,
  });
}

/**
 * mount(app, authorize) — `authorize` is the worker's existing async function
 * that returns an OAuth2 client. Called per request, exactly as the other
 * routes here do, so token refresh keeps working the same way.
 */
function mount(app, authorize, opts = {}) {
  const { google } = require('googleapis');
  const driveFor = async () => google.drive({ version: 'v3', auth: await authorize() });
  const base = opts.basePath || '/api/forecast';

  app.get(`${base}/health`, (req, res) => {
    res.json({
      ok: true,
      baselineYear: BASELINE.baseYear,
      baselineRows: BASELINE.rows.length,
      basis: BASELINE.basis,
      cached: !!cache.payload,
      cachedFor: cache.key,
    });
  });

  app.get(base, async (req, res) => {
    try {
      const drive = await driveFor();
      const file = await findNewestExport(drive);

      // The export's own date is the natural as-of; allow an override.
      const fromName = /(\d{4}-\d{2}-\d{2})/.exec(file.name || '');
      const asOf = (req.query.as_of && /^\d{4}-\d{2}-\d{2}$/.test(req.query.as_of))
        ? req.query.as_of
        : (fromName ? fromName[1] : (file.createdTime || new Date().toISOString()).slice(0, 10));

      const key = `${file.id}|${asOf}`;
      if (cache.key === key && Date.now() - cache.at < CACHE_MS && !req.query.fresh) {
        res.setHeader('X-Forecast-Cache', 'hit');
        return res.json(cache.payload);
      }

      const localPath = await downloadToTmp(drive, file);
      const payload = computePayload(localPath, file, asOf);

      cache.key = key;
      cache.at = Date.now();
      cache.payload = payload;
      res.setHeader('X-Forecast-Cache', 'miss');
      return res.json(payload);
    } catch (e) {
      console.error('[forecast] failed:', e && e.message);
      return res.status(500).json({ ok: false, error: e && e.message });
    }
  });

  console.log(`[forecast] mounted at ${base} (baseline ${BASELINE.baseYear}, ${BASELINE.rows.length} rows)`);
}

module.exports = { mount, computePayload, findNewestExport, BASELINE };
