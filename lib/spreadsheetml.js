'use strict';
/**
 * Dependency-free reader for the Trinisoft "Prenotimet ne recepsion"
 * SpreadsheetML (.xls) export.
 *
 * The file is XML, not a real Excel binary, so no library is needed. Rows are
 * pulled out one at a time and cells are placed by position, honouring the
 * ss:Index attribute that Trinisoft emits when it skips empty cells.
 */

const fs = require('fs');

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };

function decodeEntities(s) {
  if (s.indexOf('&') === -1) return s;
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (m, body) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X'
        ? parseInt(body.slice(2), 16)
        : parseInt(body.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : m;
    }
    return Object.prototype.hasOwnProperty.call(ENTITIES, body) ? ENTITIES[body] : m;
  });
}

/**
 * Split one <Row>...</Row> chunk into an array of cell strings, positioned by
 * ss:Index where present. Empty / self-closing cells become ''.
 */
function parseRow(rowXml) {
  const cells = [];
  let pos = 0; // 0-based index of the next cell to write

  // Matches either a self-closing <Cell .../> or a full <Cell ...>...</Cell>.
  const cellRe = /<Cell\b([^>]*?)(?:\/>|>([\s\S]*?)<\/Cell>)/g;
  let m;
  while ((m = cellRe.exec(rowXml)) !== null) {
    const attrs = m[1] || '';
    const inner = m[2];

    const idx = /ss:Index\s*=\s*"(\d+)"/.exec(attrs);
    if (idx) pos = parseInt(idx[1], 10) - 1; // ss:Index is 1-based

    let value = '';
    if (inner) {
      const data = /<Data\b[^>]*>([\s\S]*?)<\/Data>/.exec(inner);
      if (data) value = decodeEntities(data[1]);
    }
    while (cells.length < pos) cells.push('');
    cells[pos] = value;
    pos += 1;
  }
  return cells;
}

/**
 * Iterate every <Row> in the workbook. Yields arrays of cell strings.
 * Reads the whole file into memory: the exports run ~15 MB, which is fine,
 * and streaming buys nothing once the rows have to be aggregated anyway.
 */
function* rows(filePath) {
  const xml = fs.readFileSync(filePath, 'utf8');
  const rowRe = /<Row\b[^>]*>([\s\S]*?)<\/Row>/g;
  let m;
  while ((m = rowRe.exec(xml)) !== null) {
    yield parseRow(m[1]);
  }
}

/**
 * Read the export as objects keyed by the header row.
 * The first row whose cells contain "Nr. Prenotimit" is treated as the header.
 */
function* records(filePath) {
  let header = null;
  for (const row of rows(filePath)) {
    if (!header) {
      if (row.includes('Nr. Prenotimit')) header = row.map((h) => h.trim());
      continue;
    }
    if (!row.length) continue;
    const rec = {};
    for (let i = 0; i < header.length; i += 1) {
      if (header[i]) rec[header[i]] = row[i] === undefined ? '' : row[i];
    }
    yield rec;
  }
  if (!header) throw new Error(`No header row found in ${filePath}`);
}

module.exports = { rows, records, parseRow, decodeEntities };
