// Attach uphill trip pay to the daily-board trips already in the system (user decision 17 Nov 2026, option ก).
// Source = the reviewed proposal file "ค่าเที่ยวขาขึ้น_จับคู่กับกระดานคิว_…_เสนอ.xlsx" (Transport 01 "คิวขาขึ้น" pay rows
// matched to board trips by date + truck head + driver [+ destination]):
//   sheet "A ตรงชัด"        → attached as a whole (approved by the user) unless the team writes "ไม่ใช่"
//   sheet "B ต้องตัดสิน"    → attached only when the team writes "ใช่" or the right trip id (daily_…)
//   sheet "C ไม่พบในกระดาน" → never here (those would be NEW trips — a separate decision)
// The confirm cell is read strictly: anything that is not blank / ใช่ / ไม่ใช่ / a trip id is an error to fix in the
// file, never a guess. Safety: a trip claimed by two rows is refused for BOTH; only trips with NO money at all are
// written; the trip's date, truck head and driver must equal the file row; the write itself is per trip and
// conditional (applyPayToTrip), so a trip deleted or paid meanwhile is left alone. Pure functions (no DOM / Firebase).

import { clean } from './parseUtils.js';
import { ymd } from './importerTrips.js';
import { plateKey } from './fuelAudit.js';

export const PAY_FIELDS = ['revenue', 'driverAllowance', 'fuelCost', 'deduction', 'netTransfer'];
export const PAY_FIELD_LABELS = { revenue: 'รายรับ', driverAllowance: 'เบี้ยขับ', fuelCost: 'ค่าน้ำมัน', deduction: 'หัก', netTransfer: 'สุทธิโอน' };
export const PROPOSAL_SHEETS = { A: 'A ตรงชัด', B: 'B ต้องตัดสิน' };
export const PAY_SOURCE_LABEL = 'ค่าเที่ยวขาขึ้น Transport 01 (คิวขาขึ้น) — จับคู่กับกระดานคิว 17 พ.ย. 69';
const COL = { ym: 'เดือน', date: 'วันที่', plate: 'ทะเบียน (ไฟล์)', driver: 'คนขับ (ไฟล์)', dest: 'ปลายทาง (ไฟล์)', id: 'รหัสเที่ยวในระบบ',
    revenue: 'รายรับ', driverAllowance: 'เบี้ยขับ', fuelCost: 'ค่าน้ำมัน', deduction: 'หัก', netTransfer: 'สุทธิโอน' };
const TRIP_ID_ONLY = /^(?:(?:ใช่|ยืนยัน)\s*)?(daily_[A-Za-z0-9_-]+)$/;
const YES = /^(ใช่|yes|ok|y|✓|✔)$/i;
const NO = /^(ไม่ใช่|no|n|x|✗|✘)$/i;

export const isUphillPayProposal = (sheetNames) => (sheetNames || []).includes(PROPOSAL_SHEETS.A);

// money cell: blank → 0; a number (or numeric text with commas) ≥ 0; anything else → NaN (row refused)
const moneyOf = (v) => {
    if (v == null || clean(v) === '') return 0;
    const n = typeof v === 'number' ? v : Number(clean(v).replace(/,/g, ''));
    return Number.isFinite(n) && n >= 0 ? n : NaN;
};

/** Board driver vs file driver: drop titles and spaces; `alias` (optional) maps a board short name to a full name. */
export const driverKey = (s) => clean(s).replace(/^(นางสาว|นาง|นาย|น\.ส\.)\s*/, '').replace(/\s+/g, '');
export const makeSameDriver = (alias = (x) => x) => (fileDriver, trip) => {
    const f = driverKey(fileDriver), b = driverKey(alias(clean(trip?.driver)));
    return !!f && !!b && f === b;
};

/**
 * sheets: { [sheetName]: rows[][] } — SheetJS header:1 WITH blankrows:true, so index + 1 = the Excel row.
 * Returns { items, errors }. item: { sheet, row, date, plate, driver, dest, money, tripId, decision }
 * decision: 'attach' | 'exclude' | 'pending' | 'error'.
 */
export function parseUphillPayProposal(sheets) {
    const items = [], errors = [];
    for (const [key, name] of Object.entries(PROPOSAL_SHEETS)) {
        const rows = sheets?.[name];
        if (!rows) { if (key === 'A') errors.push(`ไม่พบชีต "${name}"`); continue; }
        const head = (rows[0] || []).map(clean);
        const at = Object.fromEntries(Object.entries(COL).map(([k, h]) => [k, head.indexOf(h)]));
        const iOk = head.findIndex(h => h.startsWith('ยืนยัน'));
        const missing = Object.entries(at).filter(([, i]) => i < 0).map(([k]) => COL[k]);
        if (missing.length || iOk < 0) { errors.push(`ชีต "${name}": ไม่พบหัวคอลัมน์ ${[...missing, ...(iOk < 0 ? ['ยืนยัน'] : [])].join(', ')}`); continue; }
        rows.forEach((r, i) => {
            if (i === 0 || !r || r.every(c => clean(c) === '')) return;
            const excelRow = i + 1;
            const where = `ชีต "${name}" แถว ${excelRow}`;
            const ok = clean(r[iOk]).replace(/\s+/g, ' ');
            const money = Object.fromEntries(PAY_FIELDS.map(f => [f, moneyOf(r[at[f]])]));
            const item = { sheet: name, row: excelRow, date: ymd(r[at.date]), plate: clean(r[at.plate]), driver: clean(r[at.driver]), dest: clean(r[at.dest]), money, tripId: clean(r[at.id]), decision: '' };
            const idCell = ok.match(TRIP_ID_ONLY);
            const fail = (msg) => { item.decision = 'error'; errors.push(`${where}: ${msg}`); };
            if (NO.test(ok)) item.decision = 'exclude';
            else if (idCell) { item.tripId = idCell[1]; item.decision = 'attach'; }
            else if (YES.test(ok)) item.decision = 'attach';
            else if (ok === '') item.decision = key === 'A' ? 'attach' : 'pending';
            else fail(`ช่องยืนยันอ่านไม่ออก "${ok}" — เขียนแค่ ใช่ / ไม่ใช่ / รหัสเที่ยว (daily_…) หรือเว้นว่าง`);
            if (item.decision === 'attach') {
                const bad = PAY_FIELDS.filter(f => Number.isNaN(money[f]));
                if (bad.length) fail(`ยอดเงินไม่ใช่ตัวเลขที่ไม่ติดลบ (${bad.map(f => PAY_FIELD_LABELS[f]).join(', ')})`);
                else if (clean(r[at.netTransfer]) === '' || !(money.netTransfer > 0)) fail('สุทธิโอนว่างหรือเป็น 0');
                else if (!item.date) fail(`วันที่อ่านไม่ได้ "${clean(r[at.date])}"`);
                else if (clean(r[at.ym]) && clean(r[at.ym]) !== item.date.slice(0, 7)) fail(`เดือน "${clean(r[at.ym])}" ไม่ตรงกับวันที่ ${item.date}`);
                else if (!item.tripId) fail('ยืนยันแล้วแต่ไม่มีรหัสเที่ยวในระบบ — ใส่รหัสเที่ยว (daily_…) ในช่องยืนยัน');
            }
            items.push(item);
        });
    }
    return { items, errors };
}

const hasMoney = (t) => PAY_FIELDS.some(f => Number(t?.[f]) || 0);

/** Why this file row may not be written onto this stored trip ('' = ok). */
export function payWriteProblem(item, trip, sameDriver = makeSameDriver()) {
    if (!trip) return 'ไม่พบเที่ยวนี้ในระบบ';
    if (clean(trip.direction) !== 'ขาขึ้น') return 'เที่ยวในระบบไม่ใช่ขาขึ้น';
    if (clean(trip.date) !== item.date) return `วันที่ไม่ตรง (ระบบ ${clean(trip.date)} / ไฟล์ ${item.date})`;
    if (plateKey(trip.plate) !== plateKey(item.plate)) return `ทะเบียนไม่ตรง (ระบบ ${clean(trip.plate)} / ไฟล์ ${item.plate})`;
    if (!sameDriver(item.driver, trip)) return `คนขับไม่ตรง (ระบบ ${clean(trip.driver)} / ไฟล์ ${item.driver})`;
    if (hasMoney(trip)) return 'เที่ยวนี้มียอดเงินอยู่แล้ว — ไม่เขียนทับ';
    return '';
}

/**
 * items: from parseUphillPayProposal; tripsByMonth: { ym: { id: trip } } (the months the items point at).
 * Returns { writes:[{ ym, id, item }], skipped:[{ item, reason }], pending, excluded, totals:{ trips, money, byMonth } }.
 */
export function planUphillPay(items, tripsByMonth, sameDriver = makeSameDriver(), closedMonths = {}) {
    const writes = [], skipped = [];
    let pending = 0, excluded = 0;
    const attach = (items || []).filter(i => i.decision === 'attach');
    for (const i of items || []) { if (i.decision === 'pending') pending++; else if (i.decision === 'exclude') excluded++; }
    const claims = {};
    for (const i of attach) (claims[i.tripId] ||= []).push(i);
    for (const item of attach) {
        const ym = item.date.slice(0, 7);
        if (claims[item.tripId].length > 1) {
            skipped.push({ item, reason: `รหัสเที่ยวนี้ถูกอ้างจาก ${claims[item.tripId].length} แถว (${claims[item.tripId].map(x => `${x.sheet} แถว ${x.row}`).join(', ')}) — ไม่แนบทั้งหมด แก้ในไฟล์ให้เหลือแถวเดียว` });
            continue;
        }
        if (closedMonths?.[ym]) { skipped.push({ item, reason: 'เดือนนี้ปิดงวดแล้ว — ไม่แนบ (เปิดงวดก่อนถ้าต้องการแนบ)' }); continue; }
        const reason = payWriteProblem(item, tripsByMonth?.[ym]?.[item.tripId], sameDriver);
        if (reason) { skipped.push({ item, reason }); continue; }
        writes.push({ ym, id: item.tripId, item });
    }
    const money = Object.fromEntries(PAY_FIELDS.map(f => [f, writes.reduce((s, w) => s + w.item.money[f], 0)]));
    const byMonth = {};
    for (const w of writes) { const b = byMonth[w.ym] ||= { trips: 0, net: 0 }; b.trips++; b.net += w.item.money.netTransfer; }
    return { writes, skipped, pending, excluded, totals: { trips: writes.length, net: money.netTransfer, money, byMonth } };
}

/**
 * The write for ONE trip, as a transaction update function: given the trip as stored right now, returns the trip with
 * the money (and where it came from) added — or undefined (= abort, nothing written) when the trip is gone, has money
 * by now, or no longer matches the file row.
 */
export function applyPayToTrip(current, write, at, sameDriver = makeSameDriver()) {
    if (!current || payWriteProblem(write.item, current, sameDriver)) return undefined;
    const next = { ...current };
    for (const f of PAY_FIELDS) next[f] = write.item.money[f];
    next.payFrom = `${PAY_SOURCE_LABEL} · ชีต ${write.item.sheet} แถว ${write.item.row}`;
    next.payAttachedAt = at;
    return next;
}
