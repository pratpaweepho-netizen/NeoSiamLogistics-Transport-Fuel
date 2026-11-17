// One-time fix (user decision A, 17 Nov 2026): trips imported from the Transport workbook were stored ONE DAY EARLY
// (SheetJS cellDates in Node returned "1 Nov 23:59:56" for the Excel date 2 Nov; importerTrips.ymd is fixed now).
// Plan: compare every stored trip with the Transport file read with correct dates, and move only the trips whose
// date is exactly one day before the file's date. Pure functions (no DOM/Firebase) so they can be unit-tested.

import { clean } from './parseUtils.js';

export const nextDay = (d) => { const t = new Date(`${d}T12:00:00Z`); t.setUTCDate(t.getUTCDate() + 1); return t.toISOString().slice(0, 10); };
export const prevDay = (d) => { const t = new Date(`${d}T12:00:00Z`); t.setUTCDate(t.getUTCDate() - 1); return t.toISOString().slice(0, 10); };

// What identifies "the same trip" between the file and the database (date left out on purpose).
const sig = (t) => [clean(t.direction), clean(t.driver), clean(t.plate), clean(t.destination), Number(t.netTransfer) || 0, Number(t.revenue) || 0].join('|');

// Trips that never came from the Transport file keep their dates.
const isFixable = (key, t) => !key.startsWith('daily_') && !key.startsWith('tpl_') && t && !['daily', 'template'].includes(t.source)
    && !t.dateFixedFrom && /^\d{4}-\d{2}-\d{2}$/.test(t.date || '') && ['ขาขึ้น', 'ขาล่อง'].includes(clean(t.direction));

// Take one copy of `value` out of a multiset (Map value → count); true when there was one.
const take = (bag, value) => { const n = bag.get(value) || 0; if (!n) return false; bag.set(value, n - 1); return true; };
const bagOf = (values) => { const m = new Map(); for (const v of values) m.set(v, (m.get(v) || 0) + 1); return m; };

/**
 * fileTrips: trips parsed from the Transport file with correct dates (uphill + downhill).
 * dbByMonth: { 'YYYY-MM': { key: trip } } — the whole trips node.
 * tdhKeyMap: { oldKey: newKey } for tdh_ keys (the key contains the date; see downhillFuel.downhillTripKeys).
 * Returns { moves, counts, ambiguous, conflicts }.
 *   moves: [{ fromYm, fromKey, toYm, toKey, trip }] — trip = stored trip with the corrected date (+ dateFixedFrom).
 */
export function planDateFix(fileTrips, dbByMonth, tdhKeyMap = {}) {
    const fileBySig = {};
    for (const t of fileTrips) (fileBySig[sig(t)] ||= []).push(t.date);
    const groups = {};
    for (const [ym, month] of Object.entries(dbByMonth || {})) {
        for (const [key, t] of Object.entries(month || {})) {
            if (!isFixable(key, t)) continue;
            (groups[sig(t)] ||= []).push({ ym, key, t });
        }
    }
    const counts = { checked: 0, shift: 0, same: 0, notInFile: 0, ambiguous: 0, conflict: 0, monthChange: 0 };
    const shifts = [], ambiguous = [];
    for (const [s, list] of Object.entries(groups)) {
        counts.checked += list.length;
        const fileDates = fileBySig[s];
        if (!fileDates) { counts.notInFile += list.length; continue; }
        // Whole group fits "one day early" → move all. Whole group fits the file → nothing to do.
        // Fits both (identical trips on consecutive days) → cannot tell → leave the whole group and list it.
        const early = bagOf(fileDates.map(prevDay)), exact = bagOf(fileDates);
        const fitsEarly = list.every(x => take(early, x.t.date)), fitsExact = list.every(x => take(exact, x.t.date));
        if (fitsEarly && fitsExact) {
            counts.ambiguous += list.length;
            for (const x of list) ambiguous.push({ ym: x.ym, key: x.key, date: x.t.date, driver: x.t.driver, plate: x.t.plate, direction: x.t.direction, reason: 'ในไฟล์มีเที่ยวเหมือนกันวันติดกัน บอกไม่ได้ว่าวันไหนถูก' });
            continue;
        }
        if (fitsEarly) { shifts.push(...list); continue; }
        if (fitsExact) { counts.same += list.length; continue; }
        // Mixed group (e.g. a trip added by hand): decide trip by trip, only when the answer is clear.
        const e2 = bagOf(fileDates.map(prevDay)), x2 = bagOf(fileDates);
        for (const x of list) {
            const isEarly = (e2.get(x.t.date) || 0) > 0, isExact = (x2.get(x.t.date) || 0) > 0;
            if (isEarly && !isExact) { take(e2, x.t.date); shifts.push(x); }
            else if (isExact && !isEarly) { take(x2, x.t.date); counts.same++; }
            else if (!isEarly && !isExact) counts.notInFile++;
            else { counts.ambiguous++; ambiguous.push({ ym: x.ym, key: x.key, date: x.t.date, driver: x.t.driver, plate: x.t.plate, direction: x.t.direction, reason: 'วันเดียวกันและวันถัดไปในไฟล์มีเที่ยวเหมือนกัน' }); }
        }
    }
    let cand = [];
    for (const x of shifts) {
        const date = nextDay(x.t.date), toYm = date.slice(0, 7);
        let toKey = x.key;
        if (x.key.startsWith('tdh_')) {
            toKey = tdhKeyMap[x.key];
            if (!toKey) { counts.ambiguous++; ambiguous.push({ ym: x.ym, key: x.key, date: x.t.date, driver: x.t.driver, plate: x.t.plate, direction: x.t.direction, reason: 'หาแถวในไฟล์ของรหัสเที่ยวนี้ไม่พบ' }); continue; }
        }
        cand.push({ fromYm: x.ym, fromKey: x.key, toYm, toKey, trip: { ...x.t, date, yearMonth: toYm, dateFixedFrom: x.t.date } });
    }
    // Never overwrite a trip that is not itself moving. Dropping one move un-vacates its path, which can block
    // another move onto that path — repeat until nothing changes.
    const conflicts = [];
    for (;;) {
        const vacated = new Set(cand.map(m => `${m.fromYm}/${m.fromKey}`));
        const blocked = cand.filter(m => dbByMonth[m.toYm]?.[m.toKey] && !vacated.has(`${m.toYm}/${m.toKey}`));
        if (!blocked.length) break;
        for (const m of blocked) conflicts.push({ from: `${m.fromYm}/${m.fromKey}`, to: `${m.toYm}/${m.toKey}` });
        const drop = new Set(blocked);
        cand = cand.filter(m => !drop.has(m));
    }
    counts.conflict = conflicts.length;
    counts.monthChange = cand.filter(m => m.toYm !== m.fromYm).length;
    counts.shift = cand.length;
    return { moves: cand, counts, ambiguous, conflicts };
}

/**
 * Group moves into one atomic multi-path update per source month, in DESCENDING month order: a trip only moves
 * forward (to the same or the next month), so by the time month M writes into M+1, every trip that leaves M+1 has
 * already left, and no later batch deletes anything in M+1. Within a batch, deletes are written first so a new path
 * that equals a path being vacated keeps the new trip. Paths are relative to trips/.
 */
export function buildDateFixBatches(moves) {
    const byMonth = {};
    for (const m of moves) (byMonth[m.fromYm] ||= []).push(m);
    return Object.keys(byMonth).sort().reverse().map(ym => {
        const updates = {};
        for (const m of byMonth[ym]) updates[`${m.fromYm}/${m.fromKey}`] = null;
        for (const m of byMonth[ym]) updates[`${m.toYm}/${m.toKey}`] = m.trip;
        return { month: ym, moves: byMonth[ym], updates };
    });
}

// JSON with sorted keys at every level — Firebase and local objects can list keys in different orders.
const stable = (v) => Array.isArray(v) ? `[${v.map(stable).join(',')}]`
    : v && typeof v === 'object' ? `{${Object.keys(v).sort().map(k => `${JSON.stringify(k)}:${stable(v[k])}`).join(',')}}`
    : JSON.stringify(v);

/**
 * Right before a batch is written: drop moves whose source changed or vanished since the preview, or whose target
 * holds a trip that has not actually left. freshByMonth: current month objects ({ key: trip }) for every month the
 * batch touches. committedFrom: source paths of moves ALREADY WRITTEN in earlier batches (not the plan) — the
 * caller adds this batch's moves after a successful write. Paths are 'YYYY-MM/key'.
 */
export function recheckBatch(batch, freshByMonth, committedFrom) {
    const dropped = [];
    let ok = [];
    for (const m of batch.moves) {
        const cur = freshByMonth[m.fromYm]?.[m.fromKey];
        const { dateFixedFrom, ...expect } = m.trip;
        // The stored trip must be exactly what the preview saw (only date / yearMonth differ).
        const same = !!cur && cur.date === dateFixedFrom && stable({ ...cur, date: m.trip.date, yearMonth: m.trip.yearMonth }) === stable(expect);
        (same ? ok : dropped).push(m);
    }
    // A target is free when empty, or vacated by an earlier written batch, or by a move kept in this batch.
    for (;;) {
        const vacated = new Set([...committedFrom, ...ok.map(m => `${m.fromYm}/${m.fromKey}`)]);
        const blocked = ok.filter(m => freshByMonth[m.toYm]?.[m.toKey] && !vacated.has(`${m.toYm}/${m.toKey}`));
        if (!blocked.length) break;
        const drop = new Set(blocked);
        dropped.push(...blocked);
        ok = ok.filter(m => !drop.has(m));
    }
    const b = buildDateFixBatches(ok)[0];
    return { month: batch.month, moves: ok, updates: b ? b.updates : {}, dropped };
}
