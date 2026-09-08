'use strict';
/**
 * The whole commercial forecast, computed from this year's cleaned bookings
 * plus the compact baseline for the closed prior year.
 *
 * Shared by the local tools and by the server endpoint, so both produce
 * identical numbers. Nothing here reads a file or the network.
 *
 * Basis for every euro figure: hotel accommodation revenue only — the room and
 * the board priced into the rate, Itaka valued at the contract rate. F&B
 * outlets and Spa are not in the source data at all.
 */

const B = require('./bookings');

const DAY = 86400000;
const MONTHS_EN = ['', 'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];
const MONTHS_SQ = ['', 'janar', 'shkurt', 'mars', 'prill', 'maj', 'qershor',
  'korrik', 'gusht', 'shtator', 'tetor', 'nëntor', 'dhjetor'];

const iso = (d) => d.toISOString().slice(0, 10);

/**
 * Aggregate the compact baseline.
 * rows are [createdDate, arrivalMonth, segment, nights, revenue].
 * Returns Map<month, Map<segment, {nights, revenue}>>.
 */
function baselineAgg(baseline, createdOnOrBefore) {
  const out = new Map();
  for (const [created, month, segment, nights, revenue] of baseline.rows) {
    if (createdOnOrBefore && created > createdOnOrBefore) continue;
    let seg = out.get(month);
    if (!seg) { seg = new Map(); out.set(month, seg); }
    const e = seg.get(segment) || { nights: 0, revenue: 0 };
    e.nights += nights; e.revenue += revenue;
    seg.set(segment, e);
  }
  return out;
}

/** Baseline bookings created inside a window — used for last year's week pick-up. */
function baselineCreatedBetween(baseline, from, to) {
  const bySeg = new Map();
  let nights = 0; let revenue = 0;
  for (const [created, , segment, n, r] of baseline.rows) {
    if (created < from || created > to) continue;
    nights += n; revenue += r;
    const e = bySeg.get(segment) || { nights: 0, revenue: 0 };
    e.nights += n; e.revenue += r;
    bySeg.set(segment, e);
  }
  return { nights, revenue, bySeg };
}

function totalOf(segMap) {
  const t = { nights: 0, revenue: 0 };
  if (!segMap) return t;
  for (const s of segMap.values()) { t.nights += s.nights; t.revenue += s.revenue; }
  return t;
}

function forecastFrom({ bookings, stats, baseline, asOf, threshold = 0.60, sourceFile, sourceCreatedTime }) {
  const asOfD = B.parseDate(asOf);
  if (!asOfD) throw new Error(`Bad as-of date: ${asOf}`);
  const year = asOfD.getUTCFullYear();
  const baseYear = baseline.baseYear;
  const lyAsOf = new Date(Date.UTC(baseYear, asOfD.getUTCMonth(), asOfD.getUTCDate()));
  const lyAsOfISO = iso(lyAsOf);
  const currentMonth = asOfD.getUTCMonth() + 1;

  // --- this year, cut at as-of and in full
  const bookedNow = B.byMonthSegment(bookings, { createdOnOrBefore: asOfD });
  const allNow = B.byMonthSegment(bookings);
  // --- baseline, same two cuts
  const lyAtDate = baselineAgg(baseline, lyAsOfISO);
  const lyFinal = baselineAgg(baseline, null);

  const key = (m) => `${year}-${String(m).padStart(2, '0')}`;

  // --- year to date, both years, cut at the same DAY of the year.
  // Month-level totals would run to the end of the current month and overstate
  // it; year to date has to stop at the as-of date itself.
  const ytdNow = (() => {
    const bySeg = new Map();
    let nights = 0; let revenue = 0;
    const from = B.parseDate(`${year}-01-01`);
    for (const b of bookings) {
      for (let i = 0; i < b.nights; i += 1) {
        const night = new Date(b.arrival.getTime() + i * DAY);
        if (night < from || night > asOfD) continue;
        const r = B.nightRevenue(b, night);
        const v = r.pms + r.itaka;
        nights += 1; revenue += v;
        const e = bySeg.get(b.segment) || { nights: 0, revenue: 0 };
        e.nights += 1; e.revenue += v;
        bySeg.set(b.segment, e);
      }
    }
    return { bySeg, nights, revenue, adr: nights ? revenue / nights : 0 };
  })();

  // Last year the same window comes from the baseline's daily stay-date table.
  const ytdLastYearAgg = (() => {
    const bySeg = new Map();
    let nights = 0; let revenue = 0;
    const from = `${baseYear}-01-01`;
    for (const [stayDate, segment, n, r] of (baseline.daily || [])) {
      if (stayDate < from || stayDate > lyAsOfISO) continue;
      nights += n; revenue += r;
      const e = bySeg.get(segment) || { nights: 0, revenue: 0 };
      e.nights += n; e.revenue += r;
      bySeg.set(segment, e);
    }
    return { bySeg, nights, revenue, adr: nights ? revenue / nights : 0 };
  })();

  const ytd = ytdNow;
  const ytdLy = ytdLastYearAgg;

  // Year-on-year rate movement, clamped so a thin comparison cannot skew it.
  const uplift = Math.min(1.5, Math.max(0.8, ytdLy.adr > 0 ? ytd.adr / ytdLy.adr : 1));

  /**
   * Rate to value one segment's pick-up nights.
   * Last year's rate for THAT month beats the year-to-date figure, because
   * year-to-date is dominated by July and August and low-season rooms do not
   * sell at peak prices.
   */
  function chooseAdr(now, lyF, name) {
    if (now.nights >= 20) return { adr: now.revenue / now.nights, rule: 'booked-this-month' };
    if (lyF.nights >= 10) return { adr: (lyF.revenue / lyF.nights) * uplift, rule: 'ly-month-uplifted' };
    if (now.nights > 0) return { adr: now.revenue / now.nights, rule: 'booked-thin-sample' };
    const y = ytd.bySeg.get(name);
    if (y && y.nights) return { adr: y.revenue / y.nights, rule: 'ytd-segment' };
    if (lyF.nights > 0) return { adr: lyF.revenue / lyF.nights, rule: 'ly-month-thin' };
    return { adr: 0, rule: 'none' };
  }

  const months = [];
  for (let m = 1; m <= 12; m += 1) {
    const segNow = bookedNow.get(key(m)) || new Map();
    const segAll = allNow.get(key(m)) || new Map();
    const segLyAt = lyAtDate.get(m) || new Map();
    const segLyFin = lyFinal.get(m) || new Map();

    const elapsed = m < currentMonth;
    const names = B.orderSegments(new Set([...segNow.keys(), ...segLyFin.keys()]));

    let impliedRev = 0; let floorRev = 0; let pickupNights = 0;
    const rules = new Set();
    const bySegment = names.map((name) => {
      const now = segNow.get(name) || { nights: 0, revenue: 0 };
      const lyA = segLyAt.get(name) || { nights: 0, revenue: 0 };
      const lyF = segLyFin.get(name) || { nights: 0, revenue: 0 };
      const puN = Math.max(0, lyF.nights - lyA.nights);
      const puR = Math.max(0, lyF.revenue - lyA.revenue);
      const { adr, rule } = chooseAdr(now, lyF, name);
      if (puN > 0) rules.add(rule);
      const impl = now.revenue + (elapsed ? 0 : puN * adr);
      const floor = now.revenue + (elapsed ? 0 : puR);
      impliedRev += impl; floorRev += floor;
      if (!elapsed) pickupNights += puN;
      return {
        segment: name,
        bookedNights: now.nights,
        bookedRevenue: Math.round(now.revenue),
        lyOtbNights: lyA.nights,
        lyFinalNights: lyF.nights,
        lyFinalRevenue: Math.round(lyF.revenue),
        pickupNights: elapsed ? 0 : puN,
        adrApplied: Math.round(adr),
        adrRule: rule,
        impliedRevenue: Math.round(impl),
        impliedNights: now.nights + (elapsed ? 0 : puN),
        gapVsLyOtb: now.nights - lyA.nights,
      };
    }).filter((s) => s.bookedNights || s.pickupNights || s.lyFinalNights);

    const nowT = totalOf(segNow);
    const allT = totalOf(segAll);
    const lyAtT = totalOf(segLyAt);
    const lyFinT = totalOf(segLyFin);

    const cap = B.capacityNights(year, m);
    const capLy = B.capacityNights(baseYear, m);
    const lyPickupN = Math.max(0, lyFinT.nights - lyAtT.nights);
    const impliedNights = elapsed ? nowT.nights : nowT.nights + lyPickupN;
    const capRatio = capLy.nights ? cap.nights / capLy.nights : 1;

    const fcNights = nowT.nights + (elapsed ? 0 : pickupNights);
    months.push({
      month: m,
      name: MONTHS_EN[m],
      nameSq: MONTHS_SQ[m],
      status: elapsed ? 'actual' : (m === currentMonth ? 'in-progress' : 'forecast'),

      // occupancy
      otb: nowT.nights,
      allNights: allT.nights,
      lyOtb: lyAtT.nights,
      lyFinal: lyFinT.nights,
      lyPickup: lyPickupN,
      impliedNights,
      capacity: cap.nights,
      capacityLastYear: capLy.nights,
      capacityOverridden: !!cap.overridden,
      otbOcc: cap.nights ? nowT.nights / cap.nights : null,
      impliedOcc: cap.nights ? impliedNights / cap.nights : null,
      impliedOccScaled: cap.nights ? (nowT.nights + lyPickupN * capRatio) / cap.nights : null,
      lyFinalOcc: capLy.nights ? lyFinT.nights / capLy.nights : null,
      gapVsLyFinal: (cap.nights && capLy.nights)
        ? impliedNights / cap.nights - lyFinT.nights / capLy.nights : null,
      paceBehind: nowT.nights < lyAtT.nights,
      lowSeason: capLy.nights ? (lyFinT.nights / capLy.nights) < threshold : false,
      belowThreshold: cap.nights ? (impliedNights / cap.nights) < threshold : false,

      // revenue
      bookedNights: nowT.nights,
      bookedRevenue: Math.round(nowT.revenue),
      pickupNights: elapsed ? 0 : pickupNights,
      impliedRevenue: Math.round(impliedRev),
      floorRevenue: Math.round(floorRev),
      lyFinalRevenue: Math.round(lyFinT.revenue),
      lyFinalNights: lyFinT.nights,
      adrForecast: fcNights ? impliedRev / fcNights : 0,
      adrLastYear: lyFinT.nights ? lyFinT.revenue / lyFinT.nights : 0,
      vsLy: lyFinT.revenue ? impliedRev / lyFinT.revenue - 1 : null,
      adrRules: [...rules],
      bySegment,
    });
  }

  const sum = (arr, k) => arr.reduce((a, x) => a + x[k], 0);
  const q4rows = months.filter((m) => m.month >= 10);
  const q4Nights = sum(q4rows, 'bookedNights') + sum(q4rows, 'pickupNights');
  const q4 = {
    bookedRevenue: sum(q4rows, 'bookedRevenue'),
    impliedRevenue: sum(q4rows, 'impliedRevenue'),
    floorRevenue: sum(q4rows, 'floorRevenue'),
    lyFinalRevenue: sum(q4rows, 'lyFinalRevenue'),
    adrForecast: q4Nights ? sum(q4rows, 'impliedRevenue') / q4Nights : 0,
    adrLastYear: sum(q4rows, 'lyFinalNights') ? sum(q4rows, 'lyFinalRevenue') / sum(q4rows, 'lyFinalNights') : 0,
  };
  q4.vsLy = q4.lyFinalRevenue ? q4.impliedRevenue / q4.lyFinalRevenue - 1 : null;

  const fullYear = {
    impliedRevenue: sum(months, 'impliedRevenue'),
    floorRevenue: sum(months, 'floorRevenue'),
    lyFinalRevenue: sum(months, 'lyFinalRevenue'),
    actualToDate: sum(months.filter((m) => m.status === 'actual'), 'bookedRevenue'),
    stillToCome: sum(months.filter((m) => m.status !== 'actual'), 'impliedRevenue'),
  };
  fullYear.vsLy = fullYear.lyFinalRevenue ? fullYear.impliedRevenue / fullYear.lyFinalRevenue - 1 : null;

  // --- week's pick-up, both years
  const wkFrom = new Date(asOfD.getTime() - 6 * DAY);
  const lyWkFrom = new Date(lyAsOf.getTime() - 6 * DAY);
  const wk = bookings.filter((b) => b.created >= wkFrom && b.created <= asOfD);
  const wkBySeg = new Map();
  let wkNights = 0; let wkRev = 0;
  for (const b of wk) {
    let r = 0;
    for (let i = 0; i < b.nights; i += 1) {
      const nr = B.nightRevenue(b, new Date(b.arrival.getTime() + i * DAY));
      r += nr.pms + nr.itaka;
    }
    wkNights += b.nights; wkRev += r;
    const e = wkBySeg.get(b.segment) || { nights: 0, revenue: 0, bookings: 0 };
    e.nights += b.nights; e.revenue += r; e.bookings += 1;
    wkBySeg.set(b.segment, e);
  }
  const wkLy = baselineCreatedBetween(baseline, iso(lyWkFrom), lyAsOfISO);

  // --- Itaka against the room guarantees
  const IT = B.DEFS.itaka_contract_2026;
  const itaka = {};
  for (const leg of ['PL', 'HU']) {
    const [pf, pt] = IT[leg].period.split('/');
    const from = B.parseDate(pf); const to = B.parseDate(pt);
    let nights = 0;
    for (const b of bookings) {
      if (b.segment !== 'Itaka') continue;
      const isHu = /HUNGAR/i.test(`${b.channel} ${b.kompania}`);
      if ((leg === 'HU') !== isHu) continue;
      for (let i = 0; i < b.nights; i += 1) {
        const d = new Date(b.arrival.getTime() + i * DAY);
        if (d >= from && d <= to) nights += 1;
      }
    }
    const days = Math.round((to - from) / DAY) + 1;
    const guaranteed = IT[leg].rooms * days;
    itaka[leg] = {
      period: IT[leg].period,
      rooms: IT[leg].rooms,
      guaranteedNights: guaranteed,
      actualNights: nights,
      aboveGuarantee: Math.max(0, nights - guaranteed),
      unused: Math.max(0, guaranteed - nights),
      materialisation: guaranteed ? nights / guaranteed : 0,
    };
  }

  const lfmMonth = currentMonth === 1 ? 12 : currentMonth - 1;
  const lfmYear = currentMonth === 1 ? year - 1 : year;
  const lfm = months.find((m) => m.month === lfmMonth) || null;

  return {
    ok: true,
    generatedAt: new Date().toISOString(),
    asOf,
    asOfLastYear: lyAsOfISO,
    year,
    baseYear,
    threshold,
    basis: 'hotel accommodation revenue only (room and the board priced into the rate, Itaka at contract rate); EXCLUDES F&B outlets and Spa',
    ytd: {
      nights: ytd.nights,
      revenue: Math.round(ytd.revenue),
      adr: ytd.adr,
      bySegment: B.orderSegments(new Set(ytd.bySeg.keys())).map((n) => ({
        segment: n,
        nights: ytd.bySeg.get(n).nights,
        revenue: Math.round(ytd.bySeg.get(n).revenue),
        adr: ytd.bySeg.get(n).nights ? ytd.bySeg.get(n).revenue / ytd.bySeg.get(n).nights : 0,
        shareNights: ytd.nights ? ytd.bySeg.get(n).nights / ytd.nights : 0,
      })),
    },
    ytdLastYear: { nights: ytdLy.nights, revenue: Math.round(ytdLy.revenue), adr: ytdLy.adr },
    rateUplift: uplift,
    lastFullMonth: lfm ? {
      year: lfmYear,
      month: lfmMonth,
      nights: lfm.allNights,
      capacity: lfm.capacity,
      occ: lfm.capacity ? lfm.allNights / lfm.capacity : null,
      revenue: lfm.bookedRevenue,
      revpar: lfm.capacity ? lfm.bookedRevenue / lfm.capacity : null,
      lyOcc: lfm.lyFinalOcc,
      lyRevpar: lfm.capacityLastYear ? lfm.lyFinalRevenue / lfm.capacityLastYear : null,
    } : null,
    week: {
      from: iso(wkFrom),
      to: asOf,
      bookings: wk.length,
      nights: wkNights,
      revenue: Math.round(wkRev),
      bySegment: B.orderSegments(new Set(wkBySeg.keys())).map((n) => ({
        segment: n, ...wkBySeg.get(n), revenue: Math.round(wkBySeg.get(n).revenue),
      })),
    },
    weekLastYear: {
      from: iso(lyWkFrom),
      to: lyAsOfISO,
      nights: wkLy.nights,
      revenue: Math.round(wkLy.revenue),
      bySegment: B.orderSegments(new Set(wkLy.bySeg.keys())).map((n) => ({
        segment: n, nights: wkLy.bySeg.get(n).nights, revenue: Math.round(wkLy.bySeg.get(n).revenue),
      })),
    },
    months,
    q4,
    fullYear,
    itaka,
    dataNotes: {
      sourceFile: sourceFile || null,
      sourceCreatedTime: sourceCreatedTime || null,
      baselineBuiltAt: baseline.builtAt,
      baselineSource: baseline.source,
      rowsRead: stats ? stats.rowsRead : null,
      droppedOutOfRange: stats ? stats.droppedNightsOutOfRange : null,
      unmappedChannels: stats ? [...stats.unmappedChannels] : [],
      workbookReconciled: false,
      capacityNote: 'June 2026 is 4240 nights (Premium Wing ramp), not 160 x 30',
      agencyNote: 'Agency vs Wholesaler is not like-for-like across years: 2025 "Agjensi" with UK nationality is reclassified as Wholesaler, 2026 is not',
    },
  };
}

module.exports = { forecastFrom, baselineAgg, MONTHS_EN, MONTHS_SQ };
