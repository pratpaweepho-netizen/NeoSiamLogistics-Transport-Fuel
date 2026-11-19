// Run: node tests/uphillPay.test.mjs   (no network)
import { parseUphillPayProposal, planUphillPay, applyPayToTrip, isUphillPayProposal, payWriteProblem, makeSameDriver } from '../src/uphillPay.js';

let fail = 0;
const eq = (name, a, b) => {
    const ok = JSON.stringify(a) === JSON.stringify(b);
    if (!ok) fail++;
    console.log(ok ? 'PASS' : 'FAIL', name, ok ? '' : `${JSON.stringify(a)} != ${JSON.stringify(b)}`);
};
const H = ['เดือน', 'วันที่', 'ทะเบียน (ไฟล์)', 'คนขับ (ไฟล์)', 'ปลายทาง (ไฟล์)', 'รายรับ', 'เบี้ยขับ', 'ค่าน้ำมัน', 'หัก', 'สุทธิโอน',
    'ทะเบียน (กระดาน)', 'คนขับ (กระดาน)', 'ปลายทาง (กระดาน)', 'รหัสเที่ยวในระบบ', 'เหตุผล'];
const row = (date, plate, net, id, ok = '', driver = 'สมชาย ใจดี') => ['2026-03', date, plate, driver, 'เชียงใหม่', 1500, 0, 400, 0, net, plate, 'สมชาย', 'เชียงใหม่', id, 'x', ok, ''];
const sheets = {
    'A ตรงชัด': [[...H, 'ยืนยัน (ใช่/ไม่ใช่)', 'หมายเหตุทีม'],
        row('2026-03-02', '70-1111-2222', 1900, 'daily_02-03-69_1'),           // r2 ok
        row('2026-03-02', '70-3333', 800, 'daily_02-03-69_2', 'ไม่ใช่'),      // r3 team excluded
        row('2026-03-03', '70-4444', 900, 'daily_03-03-69_1'),                 // r4 already has money
        row('2026-03-04', '70-5555', 700, 'daily_04-03-69_9'),                 // r5 not in the system
        row('2026-03-05', '70-6666', 600, 'daily_05-03-69_1'),                 // r6 date differs
        ['2026-03', '2026-03-06', '70-7777', 'x', 'y', 'abc', 0, 0, 0, 100, '', '', '', 'daily_06-03-69_1', 'x', '', ''], // r7 bad money
        row('2026-03-07', '70-1212', 500, 'daily_07-03-69_1', 'ผิดคัน'),      // r8 a note, not a decision → error
        [],                                                                    // r9 blank row kept (blankrows:true)
        ['2026-03', '2026-03-08', '70-1313', 'สมชาย ใจดี', 'y', 1, 0, 0, 0, '', '', '', '', 'daily_08-03-69_1', 'x', '', ''], // r10 net blank
        row('2026-03-09', '70-1414', 500, 'daily_09-03-69_1', '', 'สมหญิง'), // r11 driver differs from the board
        row('2569-03-10', '70-1515', 500, 'daily_10-03-69_1'),                // r12 Buddhist-era text date
    ],
    'B ต้องตัดสิน': [[...H, 'ยืนยัน / รหัสเที่ยวที่ถูก', 'หมายเหตุทีม'],
        row('2026-03-02', '70-8888', 500, 'daily_02-03-69_5', 'ใช่'),        // r2 ok
        row('2026-03-02', '70-8888', 400, '', 'daily_02-03-69_6'),           // r3 trip id typed
        row('2026-03-02', '70-9999', 300, 'daily_02-03-69_7'),               // r4 not decided → pending
        row('2026-03-02', '70-9999', 300, '', 'ใช่'),                        // r5 confirmed, no trip id → error
        row('2026-03-02', '70-1111', 1900, 'daily_02-03-69_1', 'ใช่'),      // r6 same trip as A r2 → both refused
        row('2026-03-02', '70-9999', 300, 'daily_02-03-69_7', 'ยืนยันไม่ได้'), // r7 not a decision → error
        row('2026-03-02', '70-9999', 300, 'daily_02-03-69_7', 'ไม่ใช่ ที่ถูกคือ daily_02-03-69_8'), // r8 mixed → error
    ],
};
eq('detects the proposal file', [isUphillPayProposal(['สรุป', 'A ตรงชัด']), isUphillPayProposal(['คิวขาขึ้น'])], [true, false]);
const p = parseUphillPayProposal(sheets);
eq('decisions per row (Excel row numbers incl. the blank row)', p.items.map(i => `${i.sheet[0]}${i.row}:${i.decision}`),
    ['A2:attach', 'A3:exclude', 'A4:attach', 'A5:attach', 'A6:attach', 'A7:error', 'A8:error', 'A10:error', 'A11:attach', 'A12:attach',
        'B2:attach', 'B3:attach', 'B4:pending', 'B5:error', 'B6:attach', 'B7:error', 'B8:error']);
const errAt = (sheetRow) => p.errors.find(e => e.includes(sheetRow)) || '';
eq('notes / mixed cells are errors, never a guess', [errAt('"A ตรงชัด" แถว 8').includes('อ่านไม่ออก'), errAt('"B ต้องตัดสิน" แถว 7').includes('อ่านไม่ออก'), errAt('"B ต้องตัดสิน" แถว 8').includes('อ่านไม่ออก')], [true, true, true]);
eq('blank net is an error', errAt('"A ตรงชัด" แถว 10').includes('สุทธิโอนว่าง'), true);
eq('BE text date normalised', p.items.find(i => i.row === 12).date, '2026-03-10');
eq('trip id typed in the confirm cell wins', p.items.find(i => i.row === 3 && i.sheet.startsWith('B')).tripId, 'daily_02-03-69_6');
const trips = { '2026-03': {
    'daily_02-03-69_1': { date: '2026-03-02', direction: 'ขาขึ้น', plate: '70-1111-9999', driver: 'สมชาย', revenue: 0, netTransfer: 0 },
    'daily_02-03-69_2': { date: '2026-03-02', direction: 'ขาขึ้น', plate: '70-3333', driver: 'สมชาย' },
    'daily_03-03-69_1': { date: '2026-03-03', direction: 'ขาขึ้น', plate: '70-4444', driver: 'สมชาย', netTransfer: 1300 },
    'daily_05-03-69_1': { date: '2026-03-04', direction: 'ขาขึ้น', plate: '70-6666', driver: 'สมชาย' },
    'daily_09-03-69_1': { date: '2026-03-09', direction: 'ขาขึ้น', plate: '70-1414', driver: 'สมชาย' },
    'daily_10-03-69_1': { date: '2026-03-10', direction: 'ขาขึ้น', plate: '70-1515', driver: 'นายสมชาย ใจดี' },
    'daily_02-03-69_5': { date: '2026-03-02', direction: 'ขาขึ้น', plate: '70-8888', driver: 'สมชาย' },
    'daily_02-03-69_6': { date: '2026-03-02', direction: 'ขาขึ้น', plate: '70-8888', driver: 'สมชาย' },
} };
const sameDriver = makeSameDriver((n) => (n === 'สมชาย' ? 'สมชาย ใจดี' : n)); // board short name → full name (alias)
const plan = planUphillPay(p.items, trips, sameDriver);
eq('writes only safe rows (truck head match, title "นาย" ignored)', plan.writes.map(w => w.id), ['daily_10-03-69_1', 'daily_02-03-69_5', 'daily_02-03-69_6']);
const reasonOf = (id) => (plan.skipped.find(s => s.item.tripId === id) || {}).reason || '';
eq('a trip claimed twice is refused for BOTH rows', plan.skipped.filter(s => s.item.tripId === 'daily_02-03-69_1').length, 2);
eq('skip reasons', [reasonOf('daily_03-03-69_1').startsWith('เที่ยวนี้มียอดเงินอยู่แล้ว'), reasonOf('daily_04-03-69_9').startsWith('ไม่พบเที่ยวนี้'), reasonOf('daily_05-03-69_1').startsWith('วันที่ไม่ตรง'), reasonOf('daily_09-03-69_1').startsWith('คนขับไม่ตรง')], [true, true, true, true]);
eq('pending / excluded counted, totals', [plan.pending, plan.excluded, plan.totals.trips, plan.totals.net, plan.totals.money.fuelCost], [1, 1, 3, 1400, 1200]);
// the per-trip conditional write
const w = plan.writes.find(x => x.id === 'daily_02-03-69_6');
const next = applyPayToTrip(trips['2026-03']['daily_02-03-69_6'], w, 'AT', sameDriver);
eq('write adds money + source, keeps the trip', [next.netTransfer, next.revenue, next.fuelCost, next.plate, next.payFrom.includes('แถว 3'), next.payAttachedAt], [400, 1500, 400, '70-8888', true, 'AT']);
eq('write aborts when the trip is gone / paid meanwhile / changed', [applyPayToTrip(null, w, 'AT', sameDriver), applyPayToTrip({ ...trips['2026-03']['daily_02-03-69_6'], netTransfer: 5 }, w, 'AT', sameDriver), applyPayToTrip({ ...trips['2026-03']['daily_02-03-69_6'], date: '2026-03-03' }, w, 'AT', sameDriver)], [undefined, undefined, undefined]);
{
    const closedPlan = planUphillPay(p.items, trips, sameDriver, { '2026-03': { closedAt: 'x' } });
    eq('closed month: nothing written, each row listed as skipped with the reason', [closedPlan.writes.length, closedPlan.skipped.filter(s => s.reason.startsWith('เดือนนี้ปิดงวดแล้ว')).length, closedPlan.totals.net], [0, plan.writes.length + plan.skipped.length - 2, 0]); // all but the 2 double-claimed rows
    eq('another month closed: same plan as before', planUphillPay(p.items, trips, sameDriver, { '2026-02': {} }).writes.map(w => w.id), plan.writes.map(w => w.id));
}
eq('downhill trip refused', payWriteProblem({ date: '2026-03-02', plate: '70-1', driver: 'a' }, { date: '2026-03-02', direction: 'ขาล่อง', plate: '70-1', driver: 'a' }), 'เที่ยวในระบบไม่ใช่ขาขึ้น');
eq('missing sheet A → error', parseUphillPayProposal({}).errors, ['ไม่พบชีต "A ตรงชัด"']);

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
