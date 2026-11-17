// Downhill liters from column L ("คิวรถล่าสุด") of the คิวขาล่อง sheet in the Transport file.
// User decision (17 Nov 2026, option ก): keep these liters in their own node downhillFuel/{YYYY-MM}/{tripKey} and
// never modify the downhill trips already in the system. The audit merges them in at read time.
// Pure functions (no DOM/Firebase); the SHA-1 function is passed in (Node crypto / browser SubtleCrypto).

import { clean } from './parseUtils.js';
import { parsePumpText } from './importerTrips.js';
import { plateKey, PAISAN, CHECKED, markExtraDownhill } from './fuelAudit.js';
export { createAllIfAbsent } from './dbUtils.js';

// "(คิว17.11.69นว.)180ไพศาล" — the bracket names the uphill queue (date dd.mm.yy BE + place); the rest is the fill.
// Tolerates the board's variants: no "คิว", "คืว", a missing "(" or ")" and a mistyped date such as "28.098.68".
// Place = everything up to ")" when there is one; without ")" only Thai letters, spaces and dots, so a fill text
// such as "XX" / "NGV" / "120ไพศาล" right after the place is not swallowed into it.
// A 3-digit year ("26.09.669", "14.07.698ชม.") is a typo: the date is left blank rather than guessed.
const PREFIX = /^\(?\s*(?:ค[ิื]ว)?\s*(\d{1,2})\.(\d{1,3})\.(\d{2,3})\s*(?:([^)]*)\)|([ก-๙ .]*))\s*(.*)$/s;

export const COLL_STATUS = {
    filled: 'มีลิตร',
    none: 'ไม่ได้เติม (XX)',
    blank: 'อ้างคิวขาขึ้นแต่ไม่ได้เขียนน้ำมัน',
    ngv: 'รถ NGV',
    partner: 'รถร่วม',
    pending: 'รอเติมขาล่อง',
    unparsed: 'อ่านไม่ได้ — ตรวจเอง',
};

const sum = (o) => Object.values(o).reduce((a, b) => a + b, 0);

/** Column L text → { status, uphillRef, gasStationsLiters, gasStationsBaht, totalLiters, totalBaht, raw } ('' text → null). */
export function parseColumnL(text) {
    const raw = clean(text);
    if (!raw) return null;
    let uphillRef = null, fuelText = raw;
    const m = raw.match(PREFIX);
    if (m) {
        fuelText = m[6];
        const place = clean(m[4] ?? m[5]);
        const d = +m[1], mo = +m[2], y = 2500 + +m[3] - 543;
        uphillRef = m[3].length === 2 && mo >= 1 && mo <= 12 && d >= 1 && d <= 31
            ? { date: `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`, place }
            : { date: '', place };
    }
    const { liters, baht } = parsePumpText(fuelText);
    const totalLiters = sum(liters), totalBaht = sum(baht);
    const f = fuelText.trim();
    const status = totalLiters || totalBaht ? 'filled'
        : /^x+$/i.test(f) ? 'none'
        : /ngv/i.test(f) ? 'ngv'
        : /รอเติม/.test(raw) ? 'pending'
        : /^รถร่วม/.test(f) ? 'partner'
        : m && !f ? 'blank'
        : 'unparsed';
    return { status, uphillRef, gasStationsLiters: liters, gasStationsBaht: baht, totalLiters, totalBaht, raw };
}

/**
 * The key the Transport downhill import used for each trip (trips/{ym}/tdh_<date>_<sha1(...)[:10]>, 17 Nov 2026):
 * sha1 of date|driver|plate|origin|destination|queueType|revenue|netTransfer|n, n = occurrence of an identical row.
 * trips: parseDownhillQueue(rows).trips in sheet order. sha1hex: async (string) => hex string.
 */
export async function downhillTripKeys(trips, sha1hex) {
    const seen = {};
    const inputs = trips.map(t => {
        const base = [t.date, t.driver, t.plate, t.origin, t.destination, t.queueType, t.revenue, t.netTransfer].map(String).join('|');
        seen[base] = (seen[base] || 0) + 1;
        return `${base}|${seen[base]}`;
    });
    const hashes = await Promise.all(inputs.map(sha1hex));
    return trips.map((t, i) => `tdh_${t.date}_${hashes[i].slice(0, 10)}`);
}

/**
 * Column L entries ready to save, only for downhill trips that exist in the system.
 * trips/lastQueue: parallel arrays from parseDownhillQueue; keys: from downhillTripKeys;
 * existingTrips: { ym: { key: trip } }; existingFuel: { ym: { key: record } } (already saved column L data).
 * byStatus counts the rows that will be saved (toSave), by column L status.
 */
export function planDownhillFuel(trips, lastQueue, keys, existingTrips, existingFuel, meta = {}) {
    const out = { toSave: [], already: 0, changed: 0, noTrip: 0, byStatus: {} };
    trips.forEach((t, i) => {
        const rec = parseColumnL(lastQueue[i]);
        if (!rec) return;
        const ym = t.yearMonth, key = keys[i];
        if (!existingTrips[ym]?.[key]) { out.noTrip++; return; }
        const prev = existingFuel[ym]?.[key];
        if (prev) { out.already++; if (clean(prev.raw) !== rec.raw) out.changed++; return; }
        out.byStatus[rec.status] = (out.byStatus[rec.status] || 0) + 1;
        out.toSave.push({ ym, key, record: { ...rec, date: t.date, plate: t.plate, driver: t.driver, ...meta } });
    });
    return out;
}

/**
 * Audit input: a downhill trip without its own liters gets the column L fill. Returns the trip unchanged otherwise.
 * The copy keeps the trip's id so duplicate checks and carry-over still see the same trip.
 */
export function withDownhillFuel(trip, rec) {
    if (!rec || rec.status !== 'filled' || clean(trip.direction) !== 'ขาล่อง' || trip.totalLiters > 0 || trip.totalBaht > 0) return trip;
    return {
        ...trip,
        gasStationsLiters: { ...(rec.gasStationsLiters || {}) }, gasStationsBaht: { ...(rec.gasStationsBaht || {}) },
        totalLiters: Number(rec.totalLiters) || 0, totalBaht: Number(rec.totalBaht) || 0,
        fuelFrom: 'colL', fuelText: rec.raw, fuelRef: rec.uphillRef || null,
    };
}

/**
 * Column L downhill fills next to the uphill trip they name ("(คิว<date><place>)").
 * When that uphill trip already holds the downhill ไพศาล fill and column L lists a fill at ANOTHER pump (นิสสัน,
 * เอกวิทัศน์), the column L liters are an extra downhill fill. The rate sheet has no rule for an extra downhill fill,
 * so such trips become 'manual' (checked by a person; carry-over passes them through) instead of a misleading
 * "under / wrong pump" against the full downhill rate. A ไพศาล fill on both is left to flagDoubleDownhill.
 * results: audit results in trip order (uphill and downhill). loadedMonths: months whose uphill trips are in results.
 * The reference is exact (date + truck written by the team), unlike flagDoubleDownhill's guess, which only adds a note
 * and leaves column-L trips with a reference to this function — so the two can run in either order.
 */
export function reconcileColumnL(results, loadedMonths) {
    const ups = new Map();
    for (const r of results) if (r.leg !== 'down' && r.downhillRecorded) ups.set(`${r.trip.date}|${plateKey(r.trip.plate)}`, r);
    for (const r of results) {
        const ref = r.trip.fuelRef?.date;
        if (r.leg !== 'down' || r.trip.fuelFrom !== 'colL' || !ref) continue;
        // Only a checked figure is replaced: a no-rate result keeps its own reason (unknown plate, no rate…).
        if (Number(r.actual?.[PAISAN]) > 0 || !CHECKED.includes(r.status)) continue;
        if (!loadedMonths.has(ref.slice(0, 7))) {
            r.flags = [...r.flags, `ช่องคิวรถล่าสุดอ้างเที่ยวขาขึ้น ${ref} ซึ่งอยู่เดือนที่ไม่ได้โหลด — ระบบยังไม่ได้เทียบว่าไพศาลขาล่องบันทึกไว้ที่ขาขึ้นหรือไม่`];
            continue;
        }
        const up = ups.get(`${ref}|${plateKey(r.trip.plate)}`);
        if (up) markExtraDownhill(r, up);
    }
    return results;
}
