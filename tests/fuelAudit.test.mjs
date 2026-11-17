// Run: node tests/fuelAudit.test.mjs   (no network, no Excel file needed)
import { parseRateSheet, buildAuditContext, auditTrip, canonDestination, applyCarry, flagDoubleDownhill } from '../src/fuelAudit.js';

let fail = 0;
const eq = (name, a, b) => {
    const ok = JSON.stringify(a) === JSON.stringify(b);
    if (!ok) fail++;
    console.log(ok ? 'PASS' : 'FAIL', name, ok ? '' : `${JSON.stringify(a)} != ${JSON.stringify(b)}`);
};

// A small rate sheet in the same shape as "อัตราเติมน้ำมันประจำ ปั๊ม ต่างๆ"
const HDR = ['ต้นทาง', 'ปลายทาง', 'ปั๊มจิตตมาศ', 'ปั๊มศรีนครคลองขลุง', 'ปั๊มสยามนิสสัน สิงห์บุรี', 'รวมยอดขาขึ้น/ลิตร', 'ปั๊มไพศาลน.ว', 'เบี้ยขับ'];
const sheet = [
    ['สถานที่', '', 'รายการเติมน้ำมัน (พ่วง) ขาขึ้น', '', '', '', 'ขาล่อง'],
    HDR,
    ['จาก กรุงเทพ', 'พิษณุโลก', 80, '***', '***', 80, '180 (victor190)', 'เบี้ยขับ 600'],
    ['จาก กรุงเทพ', 'เชียงใหม่', '80+175', 50, '***', 305, '180 (victor190)', 'เบี้ยขับ 2000'],
    ['จาก กรุงเทพ', 'เชียงใหม่', '80+185', 50, '***', 315, '180 (victor190)', 'เบี้ยขับ 2000 (รถsino)'],
    ['จาก กรุงเทพ', 'แพร่', '80+120', '***', '***', 200, '179 (victor190)', 'เบี้ยขับ 1300'],
    ['จาก นครสวรรค์', 'กำแพงเพชร', '***', '***', 80, 80, '180 (victor190)', 'เบี้ยขับ 600'],
    [],
    ['สถานที่', '', 'รายการเติมน้ำมัน(หกล้อใหม่ isuzu) ขาขึ้น (ยกเลิก)'],
    HDR,
    ['จาก กรุงเทพ', 'ลำพูน', '50+50', 50, '***', 150, 100, 'เบี้ยขับ 1200'],
];
const { standards } = parseRateSheet({ 'เหนือ': sheet });
const std = (d, v = 'standard') => standards.find(s => s.destination === d && s.variant === v);
eq('parsed rows (cancelled section skipped)', standards.length, 5);
eq('plus-sum cell', std('เชียงใหม่').up, { jittamas_bangpain: 255, srinakorn_klongklung: 50 });
eq('sino variant', std('เชียงใหม่', 'sino').up, { jittamas_bangpain: 265, srinakorn_klongklung: 50 });
eq('downhill + victor', [std('พิษณุโลก').down, std('พิษณุโลก').downVictor], [180, 190]);
eq('excel row number', std('พิษณุโลก').row, 3);
eq('odd downhill flagged', std('แพร่').flags.some(f => f.includes('179')), true);
eq('NSW origin at Singburi pump flagged', std('กำแพงเพชร').flags.some(f => f.includes('นิสสัน')), true);
eq('canon destination typo', canonDestination('นครสรรค์'), 'นครสวรรค์');
eq('canon destination district', canonDestination('บึงสามพัน เพชรบูรณ์'), 'บึงสามพัน');

const vehicles = [
    { registrationNumber: '70-1111', vehicleType: '18W', driveType: 'รถบริษัท' },
    { registrationNumber: '70-2222', vehicleType: '18W', driveType: 'รถบริษัท', model: 'victor' },
    { registrationNumber: '70-3333', vehicleType: '18W', driveType: 'รถบริษัท', model: 'sino' },
    { registrationNumber: '75-9999', vehicleType: 'รถร่วม', driveType: 'รถร่วม' },
];
const ctx = buildAuditContext(standards, vehicles);
const trip = (o) => ({ direction: 'ขาขึ้น', origin: 'HiQ', plate: '70-1111-6436', destination: 'พิษณุโลก', gasStationsLiters: {}, ...o });
const run = (o, tol = 0) => auditTrip(trip(o), ctx, tol);
eq('ok incl. downhill', run({ gasStationsLiters: { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 } }).status, 'ok');
eq('ok without downhill on this trip', run({ gasStationsLiters: { jittamas_bangpain: 80 } }).status, 'ok');
eq('over', [run({ gasStationsLiters: { jittamas_bangpain: 100 } }).status, run({ gasStationsLiters: { jittamas_bangpain: 100 } }).diff], ['over', 20]);
eq('under', run({ gasStationsLiters: { jittamas_bangpain: 60 } }).status, 'under');
eq('wrong pump', run({ gasStationsLiters: { srinakorn_klongklung: 80 } }).status, 'wrong-pump');
eq('tolerance', run({ gasStationsLiters: { jittamas_bangpain: 100 } }, 20).status, 'ok');
eq('victor downhill 190', run({ plate: '70-2222', gasStationsLiters: { jittamas_bangpain: 80, daopaisan_nakhonsawan: 190 } }).status, 'ok');
eq('sino rate used', run({ plate: '70-3333', destination: 'เชียงใหม่', gasStationsLiters: { jittamas_bangpain: 265, srinakorn_klongklung: 50 } }).status, 'ok');
eq('partner truck skipped', run({ plate: '75-9999', gasStationsLiters: { jittamas_bangpain: 80 } }), null);
eq('downhill trip with liters is audited as a downhill leg', run({ direction: 'ขาล่อง', gasStationsLiters: { jittamas_bangpain: 80 } }).leg, 'down');
eq('no fuel skipped', run({}), null);
eq('unknown plate → no-rate', run({ plate: '99-0000', gasStationsLiters: { jittamas_bangpain: 80 } }).status, 'no-rate');
eq('destination outside sheet → no-rate', run({ destination: 'สาย 3', gasStationsLiters: { jittamas_bangpain: 80 } }).status, 'no-rate');
eq('NSW branch origin', run({ origin: 'ย่อย นว', destination: 'กำแพงเพชร', gasStationsLiters: { nissan_singburi: 80 } }).origin, 'นครสวรรค์');
eq('other branch origin → no-rate', run({ origin: 'สาขา ชม', gasStationsLiters: { jittamas_bangpain: 80 } }).status, 'no-rate');

// --- downhill (ขาล่อง ไพศาล written on the uphill trip) ---
eq('downhill recorded ok', run({ gasStationsLiters: { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 } }).downStatus, 'ok');
eq('downhill not recorded', run({ gasStationsLiters: { jittamas_bangpain: 80 } }).downStatus, 'missing');
const dOver = run({ gasStationsLiters: { jittamas_bangpain: 80, daopaisan_nakhonsawan: 200 } });
eq('downhill over', [dOver.downStatus, dOver.downActual, dOver.down], ['over', 200, 180]);

// --- trailer left at Kamphaeng Phet (sheet "ถอดหางกำแพงเพชร") ---
const dropSheet = [
    [], ['', '', '', '', 'แบบปัจจุบัน (เริ่ม 7 เม.ย. 69)'],
    ['', 'ต้นทาง', 'ปลายทาง', 'จำนวนน้ำมันที่เติม(ลิตร)', 'ประเภทรถ', 'ปลายทาง', 'จำนวน น้ำมัน (ลิตร)', 'เบี้ยขับ'],
    ['', 'กำแพงเพชร', 'แม่สอด', 120, 'แม่', 'แม่สอด', 130, 500],
    ['', '', '', '', 'ลูก', 'กำแพงเพชร', 80, 300],
    ['', '', '', '', '', '', 210, 800],
];
const dropStd = parseRateSheet({ 'เหนือ': sheet, 'ถอดหางกำแพงเพชร': dropSheet }).standards.filter(s => s.cls === 'KPP_DROP');
eq('trailer-drop rows parsed', dropStd.map(s => [s.destination, s.upTotal, s.allowance, s.row]), [['แม่สอด', 210, 800, 4]]);
eq('old table difference flagged', dropStd[0].flags.some(f => f.includes('120') && f.includes('130')), true);
const ctxDrop = buildAuditContext([...standards, ...dropStd], vehicles);
const drop = auditTrip(trip({ destination: 'แม่สอด', jobDetail: 'แม่ตรงร้าน วาไรตี้/ ตัดลูก กพ  ลนต.', gasStationsLiters: { jittamas_bangpain: 150, daopaisan_nakhonsawan: 80 } }), ctxDrop);
eq('trailer-drop trip → manual check with total', [drop.status, drop.expectedTotal, drop.diff], ['manual', 210, 20]);
eq('drop at Nakhon Sawan is not a KPP drop', auditTrip(trip({ jobDetail: 'แม่ กพ /ตัดลูก นว', gasStationsLiters: { jittamas_bangpain: 80 } }), ctxDrop).status, 'ok');

// --- liters carried to the driver's next trip (like TMS) ---
const seq = [
    run({ date: '2026-11-01', gasStationsLiters: { jittamas_bangpain: 100 } }),               // +20 over
    run({ date: '2026-11-02', destination: 'สาย 3', gasStationsLiters: { jittamas_bangpain: 50 } }), // no rate: balance unchanged
    run({ date: '2026-11-03', gasStationsLiters: { jittamas_bangpain: 70 } }),                // allot 80−20=60, filled 70 → carry out 20+70−80 = +10
    { ...run({ date: '2026-11-04', gasStationsLiters: { jittamas_bangpain: 50 } }), trip: { ...trip({ driver: 'B' }), date: '2026-11-04' } }, // other driver: 50−80 = −30
    run({ date: '2026-11-05', gasStationsLiters: { jittamas_bangpain: 60 } }),                // allot 80−10=70, filled 60 → 10+60−80 = −10
];
seq.forEach((r, i) => { if (i !== 3) r.trip.driver = 'A'; });
const bal = applyCarry(seq);
// TMS DriverPayService: carry_out = carry_in + filled − standard; allot = max(0, standard − carry_in)
eq('carry chain', seq.map(r => [r.carryIn, r.allot, r.carryOut]), [[0, 80, 20], [20, null, 20], [20, 60, 10], [0, 80, -30], [10, 70, -10]]);
eq('balances per driver', bal, { A: -10, B: -30 });

// --- review fixes ---
const dropTwice = [...dropSheet,
    ['', '', '', '', 'ประเภทรถ', 'ปลายทาง', 'จำนวน น้ำมัน (ลิตร)', 'เบี้ยขับ'],
    ['', '', '', '', 'แม่', 'แม่สอด', 150, 500],
    ['', '', '', '', 'ลูก', 'กำแพงเพชร', 80, 300]];
const dupDrop = parseRateSheet({ 'ถอดหางกำแพงเพชร': dropTwice }).standards;
eq('conflicting drop rows flagged as conflict', dupDrop.every(s => s.flags.some(f => f.startsWith('ขัดกับ'))), true);
eq('audit uses the first drop row', buildAuditContext(dupDrop, vehicles).drop['แม่สอด'].upTotal, 210);
const dropPaisan = auditTrip(trip({ destination: 'แม่สอด', jobDetail: 'ตัดลูก กพ', gasStationsLiters: { jittamas_bangpain: 130, daopaisan_nakhonsawan: 260 } }), ctxDrop);
eq('drop trip names a possible downhill fill', dropPaisan.reason.includes('ไพศาล 260'), true);
const unnamed = [run({ date: '2026-11-01', driver: '', gasStationsLiters: { jittamas_bangpain: 120 } }), run({ date: '2026-11-02', driver: '  ', gasStationsLiters: { jittamas_bangpain: 80 } })];
unnamed.forEach(r => { r.trip.driver = r.trip.driver ?? ''; });
const balU = applyCarry(unnamed);
eq('unnamed drivers are not pooled', [unnamed.map(r => [r.carryIn, r.carryOut]), balU], [[[null, null], [null, null]], {}]);

// --- downhill trips (คิวขาล่อง) once the team fills ปั๊มไพศาล ---
const down = (o) => auditTrip({ direction: 'ขาล่อง', origin: 'เกษตรไทย(ตาคลี)', destination: 'สุชาวดี(สมุทรสงคราม)', plate: '70-1111-6436', gasStationsLiters: {}, ...o }, ctx);
eq('downhill trip without liters skipped', down({}), null);
eq('downhill trip ok at class rate', [down({ gasStationsLiters: { daopaisan_nakhonsawan: 180 } }).status, down({ gasStationsLiters: { daopaisan_nakhonsawan: 180 } }).leg], ['ok', 'down']);
eq('downhill trip over', [down({ gasStationsLiters: { daopaisan_nakhonsawan: 200 } }).status, down({ gasStationsLiters: { daopaisan_nakhonsawan: 200 } }).diff], ['over', 20]);
eq('downhill victor 190', down({ plate: '70-2222', gasStationsLiters: { daopaisan_nakhonsawan: 190 } }).status, 'ok');
eq('downhill other pump → wrong pump', down({ gasStationsLiters: { nissan_singburi: 180 } }).status, 'wrong-pump');
eq('downhill partner skipped', down({ plate: '75-9999', gasStationsLiters: { daopaisan_nakhonsawan: 180 } }), null);
const pair = [
    { ...run({ gasStationsLiters: { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 } }) },
    down({ gasStationsLiters: { daopaisan_nakhonsawan: 180 } }),
];
pair[0].trip = { ...pair[0].trip, driver: 'C', date: '2026-11-10' };
pair[1].trip = { ...pair[1].trip, driver: 'C', date: '2026-11-12' };
flagDoubleDownhill(pair);
eq('same downhill fill on both records is flagged', pair.map(r => r.flags.some(f => f.startsWith('อาจบันทึกขาล่องซ้ำ'))), [true, true]);
eq('downhill trip moves the carry balance', applyCarry(pair).C, 0);

// --- 10-wheel head listed twice in the vehicle master: alone (10W) and with its trailer (18W) ---
const vehPair = [
    { registrationNumber: '70-8901', vehicleType: '10W', driveType: 'รถบริษัท' },
    { registrationNumber: '70-8901-8158', vehicleType: '18W', driveType: 'รถบริษัท' },
    { registrationNumber: '70-5555', vehicleType: '10W', driveType: 'รถบริษัท' },
];
const ctxPair = buildAuditContext(standards, vehPair);
const tp = (plate) => auditTrip(trip({ plate, destination: 'เชียงใหม่', gasStationsLiters: { jittamas_bangpain: 255, srinakorn_klongklung: 50 } }), ctxPair);
eq('head+trailer plate uses the 18W row', [tp('70-8901-8158').cls, tp('70-8901-8158').status], ['18W', 'ok']);
eq('head alone uses the 10W row, with a hint', [tp('70-8901').cls, tp('70-8901').flags.some(f => f.includes('70-8901-8158'))], ['10W', true]);
eq('other trailer number on a known head → that head\'s pair', tp('70-8901-9999').cls, '18W');
eq('10W head towing an unknown trailer → พ่วง rate', tp('70-5555-1234').cls, '18W');
// review fixes for downhill
const dd = buildAuditContext([{ cls: '18W', origin: 'กรุงเทพ', destination: 'ตาก', variant: 'standard', up: { jittamas_bangpain: 130 }, down: 180, downVictor: 190, flags: [] },
    { cls: '18W', origin: 'กรุงเทพ', destination: 'แพร่', variant: 'standard', up: { jittamas_bangpain: 200 }, down: 180, downVictor: 180, flags: [] },
    { cls: '18W', origin: 'กรุงเทพ', destination: 'ลำปาง', variant: 'standard', up: { jittamas_bangpain: 200 }, down: 180, downVictor: 180, flags: [] }], []).downDefault['18W'];
eq('victor downhill taken from rows that name it', dd, { down: 180, downVictor: 190 });
const seq2 = [
    run({ gasStationsLiters: { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 } }),
    down({ gasStationsLiters: { daopaisan_nakhonsawan: 180 } }),
    run({ gasStationsLiters: { jittamas_bangpain: 80 } }),
    down({ gasStationsLiters: { daopaisan_nakhonsawan: 180 } }),
];
['2026-11-01', '2026-11-02', '2026-11-03', '2026-11-04'].forEach((d, i) => { seq2[i].trip = { ...seq2[i].trip, driver: 'D', date: d }; seq2[i].flags = []; });
flagDoubleDownhill(seq2);
eq('only the matching pair is flagged', seq2.map(r => r.flags.length), [1, 1, 0, 0]);
const bahtOnlyDown = auditTrip({ direction: 'ขาล่อง', plate: '70-1111', gasStationsLiters: {}, gasStationsBaht: { daourai_sai3: 3000 } }, ctx);
eq('baht-only downhill trip is listed, not dropped', [bahtOnlyDown?.status, bahtOnlyDown?.bahtTotal], ['no-rate', 3000]);
eq('pump names in Thai', down({ gasStationsLiters: { nissan_singburi: 180 } }).flags.some(f => f.includes('nissan_singburi')), false);

// Codex round: a downhill trip with no fill in between breaks the pairing; malformed plates stay unmatched
const seq3 = [run({ gasStationsLiters: { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 } }), down({ gasStationsLiters: { daopaisan_nakhonsawan: 180 } })];
seq3[0].trip = { ...seq3[0].trip, driver: 'E', date: '2026-11-01', id: 'u1' }; seq3[0].flags = [];
seq3[1].trip = { ...seq3[1].trip, driver: 'E', date: '2026-11-04', id: 'd2' }; seq3[1].flags = [];
flagDoubleDownhill(seq3, [seq3[0].trip, { id: 'd1', direction: 'ขาล่อง', driver: 'E', date: '2026-11-02' }, seq3[1].trip]);
eq('downhill trip without fill in between → no duplicate flag', seq3.map(r => r.flags.length), [0, 0]);
eq('4-part plate does not inherit the trailer rate', tp('70-8901-8158-X').cls !== '18W' || tp('70-8901-8158-X').flags.some(f => f.startsWith('กระดานคิวเขียนแค่ทะเบียนหัว')), true);

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
