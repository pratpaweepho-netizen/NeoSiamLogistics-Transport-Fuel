// Trip money from the allowance table (ตารางเบี้ยเลี้ยง). Pure functions, no DOM.
//
// User decision 17 Nov 2026 (option ก), taken from the way accounting paid 214 real downhill trips with a gas/fuel cost
// (2025-01 → 2026-11, net = เบี้ยขับ − หัก on all 214; the new formula reproduces 158 of the 159 that have a table rate):
//   เบี้ยขับ   = เบี้ยตามตาราง + ค่าน้ำมัน/แก๊ส + สำรองค่าลง/ตั้งเบิกเพิ่ม
//   หัก       = ตามตาราง
//   สุทธิโอน  = เบี้ยขับ − หัก
// The extra advance (2,650–3,000 baht on 56 NGV "หนักล่อง" trips, notes "ตั้งเบิกที่ …") cannot be derived from
// anything stored, so it is typed — and for NGV หนักล่อง it is required (0 = none).
// The table is the DOWNHILL allowance table; an uphill trip with a fuel cost has no confirmed formula → typed by hand.

/** Same normalisation as logisticsService.findRate (trim, lower case, any whitespace incl. NBSP → one space). */
const norm = (s) => String(s ?? '').replace(/ /g, ' ').trim().toLowerCase().replace(/\s+/g, ' ');

/** Trip types that must state their extra advance (0 allowed). vehicleType: text the type contains; runType: equal. */
export const DEFAULT_ADVANCE_RULES = [{ vehicleType: 'ngv', runType: 'หนักล่อง' }];

export const needsAdvance = (queueType, vehicleType, rules = DEFAULT_ADVANCE_RULES) =>
    (rules || []).some(r => r && norm(r.runType) && norm(vehicleType).includes(norm(r.vehicleType)) && norm(queueType) === norm(r.runType));

const amount = (v) => { const t = String(v ?? '').trim(); return t === '' ? 0 : Number(t); };
const okAmount = (x) => Number.isFinite(x) && x >= 0;

/** What the calc depends on, as one comparable string (stored with the calc's own amounts). */
export const moneyInputsKey = ({ queueType, vehicleType, direction, fuelCost, advance }) =>
    JSON.stringify([norm(queueType), norm(vehicleType), norm(direction), amount(fuelCost), String(advance ?? '').trim()]);

/**
 * rate: { allowance, deduction } from the allowance table (or null). inputs: { fuelCost, advance ('' = blank),
 * queueType, vehicleType, direction }.
 * Returns { ok: true, a, d, n, table, f, v, key } or { ok: false, reason: 'no-rate' | 'need-advance' | 'bad-number' | 'uphill-fuel' }.
 */
export function calcTripMoney(rate, inputs = {}) {
    const { fuelCost = '', advance = '', queueType = '', vehicleType = '', direction = '' } = inputs;
    const f = amount(fuelCost), v = amount(advance);
    if (!okAmount(f) || !okAmount(v)) return { ok: false, reason: 'bad-number' };
    if (norm(direction) === 'ขาขึ้น' && f > 0) return { ok: false, reason: 'uphill-fuel' };
    if (!rate) return { ok: false, reason: 'no-rate' };
    if (String(advance ?? '').trim() === '' && needsAdvance(queueType, vehicleType)) return { ok: false, reason: 'need-advance' };
    const table = Number(rate.allowance) || 0, d = Number(rate.deduction) || 0;
    const a = table + f + v;
    return { ok: true, a, d, n: a - d, table, f, v, key: moneyInputsKey(inputs) };
}

export const MONEY_MSG = {
    'need-advance': 'รถ NGV คิวหนักล่อง: กรอกช่อง "สำรองค่าลง / ตั้งเบิกเพิ่ม" ก่อน (ไม่มีใส่ 0) แล้วระบบจะคำนวณเบี้ยขับและสุทธิโอนให้',
    'bad-number': 'ค่าน้ำมัน/แก๊ส และ สำรองค่าลง ต้องเป็นตัวเลขไม่ติดลบ',
    'uphill-fuel': 'เที่ยวขาขึ้นที่มีค่าน้ำมัน: ตารางเบี้ยเลี้ยงเป็นอัตราขาล่อง และยังไม่มีสูตรที่ยืนยันสำหรับขาขึ้น — กรอกยอดเงินเอง',
    stale: 'ประเภทคิว / ประเภทรถ / ทิศทาง / ค่าน้ำมัน หรือ สำรองค่าลง เปลี่ยนหลังจากคำนวณยอดเงิน — กดปุ่มคำนวณใหม่ หรือกรอกยอดเองก่อนบันทึก',
};

/**
 * Checked before saving a trip. auto: what the calc stored ({ a, d, n, key }) or null when the money is typed.
 * Returns '' (ok to save) or the reason it is blocked (rule: unsure money blocks saving).
 */
export function moneySaveProblem({ auto, queueType, vehicleType, direction, fuelCost, advance, allowance, net }) {
    if (!okAmount(amount(fuelCost)) || !okAmount(amount(advance))) return MONEY_MSG['bad-number'];
    if (auto && auto.key !== moneyInputsKey({ queueType, vehicleType, direction, fuelCost, advance })) return MONEY_MSG.stale;
    // NGV หนักล่อง with the advance not stated and no money at all = the calc refused and nobody typed the amounts.
    if (String(advance ?? '').trim() === '' && needsAdvance(queueType, vehicleType) && !amount(allowance) && !amount(net)) return MONEY_MSG['need-advance'];
    return '';
}

/** One-line explanation of the calculated amounts, for the note under the net field. */
export function describeTripMoney(m) {
    if (!m?.ok) return '';
    const parts = [`ตาราง ${m.table.toLocaleString()}`];
    if (m.f) parts.push(`ค่าน้ำมัน/แก๊ส ${m.f.toLocaleString()}`);
    if (m.v) parts.push(`สำรองค่าลง/ตั้งเบิกเพิ่ม ${m.v.toLocaleString()}`);
    return `เบี้ยขับ = ${parts.join(' + ')} = ${m.a.toLocaleString()} · สุทธิโอน = ${m.a.toLocaleString()} − หัก ${m.d.toLocaleString()} = ${m.n.toLocaleString()}`;
}

/** Value to store for the advance field: '' (not stated) → null so it is never mistaken for "0 = none". */
export const storedAdvance = (v) => { const t = String(v ?? '').trim(); return t === '' ? null : Number(t); };
