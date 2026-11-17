// Run: node tests/tripMoney.test.mjs   (no network)
import { calcTripMoney, needsAdvance, describeTripMoney, moneySaveProblem, storedAdvance, MONEY_MSG } from '../src/tripMoney.js';

let fail = 0;
const eq = (name, a, b) => {
    const ok = JSON.stringify(a) === JSON.stringify(b);
    if (!ok) fail++;
    console.log(ok ? 'PASS' : 'FAIL', name, ok ? '' : `${JSON.stringify(a)} != ${JSON.stringify(b)}`);
};
const NGV = { allowance: 1500, deduction: 200 };
const pick = (m) => m.ok ? [m.a, m.d, m.n] : m.reason;

// real trips (Transport คิวขาล่อง): allowance 6,200 / ded 200 / net 6,000 with gas 4,700
eq('NGV ตีเปล่า + gas 4,700 → 6,200 / 200 / 6,000', pick(calcTripMoney(NGV, { fuelCost: 4700, queueType: 'ตีเปล่า', vehicleType: 'รถNGV' })), [6200, 200, 6000]);
// real: 2026-07-17 พรชัย หนักล่อง allow 9,400 fuel 4,900 ded 200 net 9,200 → advance 3,000
eq('NGV หนักล่อง + gas 4,900 + advance 3,000 → 9,400 / 200 / 9,200', pick(calcTripMoney(NGV, { fuelCost: 4900, advance: '3000', queueType: 'หนักล่อง', vehicleType: 'รถNGV' })), [9400, 200, 9200]);
eq('NGV หนักล่อง with advance 0 (none) → 6,200 / 6,000', pick(calcTripMoney(NGV, { fuelCost: 4700, advance: '0', queueType: 'หนักล่อง', vehicleType: 'รถNGV' })), [6200, 200, 6000]);
eq('NGV หนักล่อง with the advance left blank → refused', pick(calcTripMoney(NGV, { fuelCost: 4700, advance: '', queueType: 'หนักล่อง', vehicleType: 'รถNGV' })), 'need-advance');
eq('no fuel, no advance → same as before (table − deduction)', pick(calcTripMoney({ allowance: 1500, deduction: 200 }, { queueType: 'หนักล่อง', vehicleType: '18W' })), [1500, 200, 1300]);
eq('no rate → no-rate', pick(calcTripMoney(null, {})), 'no-rate');
eq('negative / text numbers refused', [pick(calcTripMoney(NGV, { fuelCost: -1 })), pick(calcTripMoney(NGV, { advance: 'abc' }))], ['bad-number', 'bad-number']);
eq('needsAdvance only for NGV หนักล่อง', [needsAdvance('หนักล่อง', 'รถNGV'), needsAdvance(' หนักล่อง ', 'รถ ngv'), needsAdvance('ตีเปล่า', 'รถNGV'), needsAdvance('หนักล่อง', '18W')], [true, true, false, false]);
eq('note explains the sum', describeTripMoney(calcTripMoney(NGV, { fuelCost: 4900, advance: 3000, queueType: 'หนักล่อง', vehicleType: 'รถNGV' })),
    'เบี้ยขับ = ตาราง 1,500 + ค่าน้ำมัน/แก๊ส 4,900 + สำรองค่าลง/ตั้งเบิกเพิ่ม 3,000 = 9,400 · สุทธิโอน = 9,400 − หัก 200 = 9,200');

// review round (17 Nov 2026)
eq('uphill trip with a fuel cost → refused (no confirmed formula)', pick(calcTripMoney(NGV, { fuelCost: 3000, queueType: 'ตีเปล่า', vehicleType: 'รถNGV', direction: 'ขาขึ้น' })), 'uphill-fuel');
eq('uphill without fuel → table − deduction as before', pick(calcTripMoney(NGV, { queueType: 'ตีเปล่า', vehicleType: 'รถNGV', direction: 'ขาขึ้น' })), [1500, 200, 1300]);
eq('needsAdvance uses findRate normalisation (double / NBSP spaces)', [needsAdvance('หนักล่อง  ', 'รถ NGV'), needsAdvance(' หนักล่อง', 'รถNGV')], [true, true]);
const auto = (inputs) => { const m = calcTripMoney(NGV, inputs); return { a: m.a, d: m.d, n: m.n, key: m.key }; };
const base = { fuelCost: '4700', advance: '0', queueType: 'หนักล่อง', vehicleType: 'รถNGV', direction: 'ขาล่อง' };
const au = auto(base);
eq('save ok when nothing changed since the calc', moneySaveProblem({ auto: au, ...base, allowance: 6200, net: 6000 }), '');
eq('save blocked when the vehicle type changed after the calc (e.g. plate picked)', moneySaveProblem({ auto: au, ...base, vehicleType: '18W', allowance: 6200, net: 6000 }), MONEY_MSG.stale);
eq('save blocked when the queue type changed after the calc', moneySaveProblem({ auto: au, ...base, queueType: 'คิวพิเศษ', allowance: 6200, net: 6000 }), MONEY_MSG.stale);
eq('save blocked when gas or advance changed after the calc', [moneySaveProblem({ auto: au, ...base, fuelCost: '5000' }), moneySaveProblem({ auto: au, ...base, advance: '3000' })], [MONEY_MSG.stale, MONEY_MSG.stale]);
eq('NGV หนักล่อง, advance blank, no money → blocked', moneySaveProblem({ auto: null, ...base, advance: '', allowance: '', net: '' }), MONEY_MSG['need-advance']);
eq('old trip with typed money and blank advance → allowed (typed money is the user\'s)', moneySaveProblem({ auto: null, ...base, advance: '', allowance: 9400, net: 9200 }), '');
eq('negative advance / fuel blocked', [moneySaveProblem({ auto: null, ...base, advance: '-3000' }), moneySaveProblem({ auto: null, ...base, fuelCost: '-1' })], [MONEY_MSG['bad-number'], MONEY_MSG['bad-number']]);
eq('blank advance stored as null, 0 kept as 0', [storedAdvance(''), storedAdvance('  '), storedAdvance('0'), storedAdvance('2900')], [null, null, 0, 2900]);

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
