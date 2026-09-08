'use strict';
/**
 * Turns raw Trinisoft export rows into cleaned booking records and stay-night
 * aggregates. Every rule here comes from definitions.json in the Drive folder
 * "Flower Data Layer" — do not invent rules locally.
 */

const path = require('path');
const { records } = require('./spreadsheetml');

const DEFS = require(path.join(__dirname, '..', 'definitions.json'));

const DAY_MS = 86400000;

/** Parse "2026-06-14T00:00:00" into a UTC date at midnight. Returns null if unusable. */
function parseDate(s) {
  if (!s) return null;
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s).trim());
  if (!m) return null;
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  return Number.isNaN(d.getTime()) ? null : d;
}

function parseNum(s) {
  if (s === undefined || s === null || s === '') return 0;
  const n = Number(String(s).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : 0;
}

function monthKey(d) {
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
}

/** Channel string as it appears in "Burimi Info.", lightly normalised. */
function normaliseChannel(raw) {
  const c = (raw || '').trim();
  if (!c) return '';
  if (/^booking(\.com)?$/i.test(c)) return 'Booking.com';
  if (/^website/i.test(c)) return 'Website/Booking Engine';
  if (/^itaka\s+poloni/i.test(c)) return 'ITAKA POLONI';
  if (/^itaka\s+hungary/i.test(c)) return 'ITAKA HUNGARY';
  return c;
}

/**
 * Segment for a row. Returns { segment, unmappedChannel } so the caller can
 * report channel codes that definitions.json does not yet know about.
 */
function classifySegment(channel, kompania, nationality, arrival) {
  // Any row whose Kompania mentions ITAKA is Itaka regardless of channel.
  if (/itaka/i.test(kompania || '')) return { segment: 'Itaka', unmappedChannel: null };
  if (/^itaka/i.test(channel)) return { segment: 'Itaka', unmappedChannel: null };

  const mapped = DEFS.channel_to_segment[channel];
  if (mapped) {
    // 2025 only: "Agjensi" rows with UK nationality were contracted UK business.
    if (
      mapped === 'Agency'
      && arrival && arrival.getUTCFullYear() === 2025
      && /^(gb|united kingdom|uk)$/i.test((nationality || '').trim())
    ) {
      return { segment: 'Wholesaler', unmappedChannel: null };
    }
    return { segment: mapped, unmappedChannel: null };
  }
  return { segment: 'Other', unmappedChannel: channel || '(blank)' };
}

function classifyWing(roomType) {
  const t = (roomType || '').toUpperCase();
  if (/^GARDEN|HOT-?TUB/.test(t)) return 'Garden';
  if (/LOFT SUITE|CLASSIC ROOM/.test(t)) return 'Premium Wing';
  return 'Flower';
}

/**
 * Load and clean one export file.
 * Returns { bookings, stats } where stats explains everything that was dropped.
 */
function load(filePath) {
  const bookings = [];
  const stats = {
    file: path.basename(filePath),
    rowsRead: 0,
    droppedNoDates: 0,
    droppedNightsOutOfRange: 0,
    droppedNoCreatedDate: 0,
    unmappedChannels: new Map(),
  };

  for (const r of records(filePath)) {
    stats.rowsRead += 1;

    const arrival = parseDate(r['Data Fillimit']);
    const departure = parseDate(r['Data Mbarimit']);
    const created = parseDate(r['Data Krijimit']);

    if (!arrival || !departure) { stats.droppedNoDates += 1; continue; }

    const nights = Math.round((departure - arrival) / DAY_MS);
    // valid = 1..30 nights; anything longer is house use / staff / a placeholder
    if (!(nights >= 1 && nights <= 30)) { stats.droppedNightsOutOfRange += 1; continue; }

    if (!created) { stats.droppedNoCreatedDate += 1; continue; }

    const channel = normaliseChannel(r['Burimi Info.']);
    const kompania = (r['Kompania'] || '').trim();
    const nationality = (r['Nacionaliteti'] || '').trim();
    const { segment, unmappedChannel } = classifySegment(channel, kompania, nationality, arrival);
    if (unmappedChannel) {
      stats.unmappedChannels.set(
        unmappedChannel,
        (stats.unmappedChannels.get(unmappedChannel) || 0) + 1,
      );
    }

    bookings.push({
      ref: r['Nr. Prenotimit'],
      arrival,
      departure,
      created,
      nights,
      roomType: (r['LLoji Dhomes'] || '').trim(),
      wing: classifyWing(r['LLoji Dhomes']),
      channel,
      segment,
      kompania,
      nationality,
      revenuePms: parseNum(r['Totali']),
      // No guest name, email, phone or card data is carried past this point.
    });
  }

  return { bookings, stats };
}

/**
 * Stay-nights per calendar month. Each booking is expanded night by night, so a
 * stay crossing a month boundary contributes to both months. This is what
 * occupancy has to be measured on; it is NOT the same as bucketing by arrival
 * month, which is what baseline_2025_pace_weekly.csv does.
 *
 * opts.createdOnOrBefore — count only bookings created on or before this date
 *                          (the "on the books as at" cut).
 */
function nightsByMonth(bookings, opts = {}) {
  const cut = opts.createdOnOrBefore || null;
  const out = new Map();
  for (const b of bookings) {
    if (cut && b.created > cut) continue;
    for (let i = 0; i < b.nights; i += 1) {
      const night = new Date(b.arrival.getTime() + i * DAY_MS);
      const k = monthKey(night);
      out.set(k, (out.get(k) || 0) + 1);
    }
  }
  return out;
}

/**
 * Same expansion, split by sales segment. The segment comes from "Burimi Info."
 * in the export, mapped through channel_to_segment in definitions.json, with
 * the Itaka and 2025-UK-agency overrides applied in classifySegment.
 *
 * Returns Map<'YYYY-MM', Map<segment, nights>>.
 */
function nightsByMonthAndSegment(bookings, opts = {}) {
  const cut = opts.createdOnOrBefore || null;
  const out = new Map();
  for (const b of bookings) {
    if (cut && b.created > cut) continue;
    for (let i = 0; i < b.nights; i += 1) {
      const night = new Date(b.arrival.getTime() + i * DAY_MS);
      const k = monthKey(night);
      let seg = out.get(k);
      if (!seg) { seg = new Map(); out.set(k, seg); }
      seg.set(b.segment, (seg.get(b.segment) || 0) + 1);
    }
  }
  return out;
}

// --- revenue -------------------------------------------------------------
//
// Itaka nights are entered at 0 in the PMS, so they are valued at the contract
// rate when quoting "revenue incl. Itaka". Every other booking is valued at its
// Totali, spread evenly across its nights.

const IT = DEFS.itaka_contract_2026;

/** Which Itaka leg a booking belongs to. Defaults to PL. */
function itakaLeg(b) {
  return /HUNGAR/i.test(`${b.channel} ${b.kompania}`) ? 'HU' : 'PL';
}

/** Contract rate for one Itaka night: main-period rate inside the main window. */
function itakaNightRate(date, leg) {
  const [mFrom, mTo] = IT[leg].main.split('/');
  const iso = date.toISOString().slice(0, 10);
  return (iso >= mFrom && iso <= mTo)
    ? IT.rate_per_room_night_main
    : IT.rate_per_room_night_shoulder;
}

/**
 * Revenue attributable to one stay-night of a booking.
 * Returns { pms, itaka } so callers can quote either basis.
 */
function nightRevenue(b, date) {
  if (b.segment === 'Itaka' && !(b.revenuePms > 0)) {
    return { pms: 0, itaka: itakaNightRate(date, itakaLeg(b)) };
  }
  return { pms: b.nights ? b.revenuePms / b.nights : 0, itaka: 0 };
}

/**
 * Nights AND revenue per calendar month per segment, night by night.
 * Returns Map<'YYYY-MM', Map<segment, {nights, pms, itaka, revenue}>>.
 */
function byMonthSegment(bookings, opts = {}) {
  const cut = opts.createdOnOrBefore || null;
  const out = new Map();
  for (const b of bookings) {
    if (cut && b.created > cut) continue;
    for (let i = 0; i < b.nights; i += 1) {
      const night = new Date(b.arrival.getTime() + i * DAY_MS);
      const k = monthKey(night);
      let seg = out.get(k);
      if (!seg) { seg = new Map(); out.set(k, seg); }
      let s = seg.get(b.segment);
      if (!s) { s = { nights: 0, pms: 0, itaka: 0, revenue: 0 }; seg.set(b.segment, s); }
      const r = nightRevenue(b, night);
      s.nights += 1; s.pms += r.pms; s.itaka += r.itaka; s.revenue += r.pms + r.itaka;
    }
  }
  return out;
}

/** Collapse a Map<segment,{...}> into one total. */
function totalOf(segMap) {
  const t = { nights: 0, pms: 0, itaka: 0, revenue: 0 };
  if (!segMap) return t;
  for (const s of segMap.values()) {
    t.nights += s.nights; t.pms += s.pms; t.itaka += s.itaka; t.revenue += s.revenue;
  }
  return t;
}

/** Every segment name seen, in a stable reporting order. */
const SEGMENT_ORDER = ['Direct', 'Wholesaler', 'Itaka', 'OTA', 'Agency', 'Tour Op', 'Other'];

function orderSegments(names) {
  const known = SEGMENT_ORDER.filter((s) => names.has(s));
  const rest = [...names].filter((s) => !SEGMENT_ORDER.includes(s)).sort();
  return [...known, ...rest];
}

/** Rooms available on a given date, from the capacity calendar. */
function roomsOn(date) {
  const iso = date.toISOString().slice(0, 10);
  for (const p of DEFS.capacity_calendar) {
    if (iso >= p.from && iso <= p.to) return p.rooms;
  }
  return null;
}

/**
 * Sellable room-nights in a month, summed day by day off the capacity calendar.
 *
 * June 2026 is overridden to 4240 nights: the Premium Wing opened gradually
 * through the month, and definitions.json records that the FLOW dashboard's
 * AVAIL_2026 uses that ramp rather than a flat 160 x 30. Using 4800 here would
 * understate June occupancy against every other Flower report.
 */
const CAPACITY_OVERRIDES = { '2026-06': 4240 };

function capacityNights(year, month) {
  const key = `${year}-${String(month).padStart(2, '0')}`;
  if (CAPACITY_OVERRIDES[key] !== undefined) {
    return { nights: CAPACITY_OVERRIDES[key], overridden: true };
  }
  const days = new Date(Date.UTC(year, month, 0)).getUTCDate();
  let total = 0;
  let missing = false;
  for (let d = 1; d <= days; d += 1) {
    const rooms = roomsOn(new Date(Date.UTC(year, month - 1, d)));
    if (rooms === null) { missing = true; } else { total += rooms; }
  }
  return { nights: total, overridden: false, missing };
}

module.exports = {
  DEFS,
  load,
  nightsByMonth,
  nightsByMonthAndSegment,
  byMonthSegment,
  totalOf,
  nightRevenue,
  itakaLeg,
  itakaNightRate,
  orderSegments,
  SEGMENT_ORDER,
  capacityNights,
  roomsOn,
  monthKey,
  parseDate,
  DAY_MS,
};
