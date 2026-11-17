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
eq('drop at Nakhon Sawan is not a KPP drop', auditTrip(trip({ jobDetail: 'แม่ กพ /ตัดลูก นว', gasStationsLiters: { jittamas_bangpain: 80 } }), ctxDrop).status !== 'manual', true);
// User rule 17 Nov 2026: trailer dropped at นว → head (10-wheeler) rate uphill, towing rate downhill.
const nswStd = [
    { cls: '10W', origin: 'กรุงเทพ', destination: 'กำแพงเพชร', variant: 'standard', up: { jittamas_bangpain: 50 }, down: 120, downVictor: 120, flags: [] },
    { cls: '18W', origin: 'กรุงเทพ', destination: 'กำแพงเพชร', variant: 'standard', up: { jittamas_bangpain: 80 }, down: 180, downVictor: 190, flags: [] },
];
const ctxNsw = buildAuditContext(nswStd, [{ registrationNumber: '70-8901-8158', vehicleType: '18W', driveType: 'รถบริษัท' }]);
const nsw = (L, job = 'แม่ กพ /ตัดลูก นว ตัดลูกนว.') => auditTrip({ direction: 'ขาขึ้น', origin: 'สหพัฒน์', plate: '70-8901-8158', destination: 'กำแพงเพชร', jobDetail: job, gasStationsLiters: L }, ctxNsw);
eq('NSW drop: 10W uphill + towing downhill = ok', [nsw({ jittamas_bangpain: 50, daopaisan_nakhonsawan: 180 }).status, nsw({ jittamas_bangpain: 50, daopaisan_nakhonsawan: 180 }).expectedTotal], ['ok', 230]);
// User rule 17 Nov 2026: "ถ้าเป็นรถพ่วง ไปสาขากพ. จ่าย 80 ลิตร" — trailer left at กพ. and the trip ends at กพ.
// = an ordinary พ่วง trip to กำแพงเพชร (real trips: จิตตมาศ 80 + ไพศาล 180 downhill), not a manual check.
const kpp = (L, o = {}) => auditTrip({ direction: 'ขาขึ้น', origin: 'สหพัฒน์', plate: '70-8901-8158', destination: 'กำแพงเพชร', jobDetail: 'แม่ตรงร้าน วาไรตี้ กพ/ ตัดลูก กพ', gasStationsLiters: L, ...o }, ctxNsw);
eq('KPP drop ending at กพ. → พ่วง rate 80 + downhill 180 = ok', [kpp({ jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 }).status, kpp({ jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 }).expectedTotal], ['ok', 260]);
eq('KPP drop ending at กพ. with an extra เต็กย้ง 53 → over +53', [kpp({ jittamas_bangpain: 80, daopaisan_nakhonsawan: 180, tekyong_chonburi: 53 }).status, kpp({ jittamas_bangpain: 80, daopaisan_nakhonsawan: 180, tekyong_chonburi: 53 }).diff], ['over', 53]);
eq('template trailerDrop กำแพงเพชร ending at กพ. → normal rate too', kpp({ jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 }, { jobDetail: '', trailerDrop: 'กำแพงเพชร' }).status, 'ok');
eq('NSW drop typo "ตักลูกนว." recognised', nsw({ jittamas_bangpain: 50, daopaisan_nakhonsawan: 180 }, 'ตักลูกนว.').status, 'ok');
// review fixes for the NSW rule
const ctxNsw2 = buildAuditContext([...nswStd,
    { cls: '18W', origin: 'กรุงเทพ', destination: 'นครสวรรค์', variant: 'standard', up: {}, down: 180, downVictor: 190, flags: [] },
    { cls: '18W', origin: 'กรุงเทพ', destination: 'พิษณุโลก', variant: 'standard', up: { jittamas_bangpain: 80 }, down: 180, downVictor: 190, flags: [] }],
    [{ registrationNumber: '70-8901', vehicleType: '10W', driveType: 'รถบริษัท' }, { registrationNumber: '70-8901-8158', vehicleType: '18W', driveType: 'รถบริษัท', model: 'victor' }]);
const n2 = (o) => auditTrip({ direction: 'ขาขึ้น', origin: 'HiQ', plate: '70-8901-8158', jobDetail: 'ตัดลูก นว', gasStationsLiters: {}, ...o }, ctxNsw2);
eq('NSW drop, no 10W rate for the route → towing rate with a note', [n2({ destination: 'พิษณุโลก', gasStationsLiters: { jittamas_bangpain: 80 } }).status, n2({ destination: 'พิษณุโลก', gasStationsLiters: { jittamas_bangpain: 80 } }).flags.some(f => f.includes('ใช้อัตรา'))], ['ok', true]);
eq('destination นครสวรรค์ itself → no drop rule', n2({ destination: 'นครสวรรค์', gasStationsLiters: { daopaisan_nakhonsawan: 190 } }).status, 'ok');
eq('head-only plate on NSW drop uses the pair (victor 190 downhill)', auditTrip({ direction: 'ขาขึ้น', origin: 'HiQ', plate: '70-8901', jobDetail: 'ตัดลูกนว.', destination: 'กำแพงเพชร', gasStationsLiters: { jittamas_bangpain: 50, daopaisan_nakhonsawan: 190 } }, ctxNsw2).status, 'ok');
eq('"ลูกนวล" is not a drop', auditTrip({ direction: 'ขาขึ้น', origin: 'HiQ', plate: '70-8901-8158', jobDetail: 'ส่งร้านลูกนวล', destination: 'กำแพงเพชร', gasStationsLiters: { jittamas_bangpain: 80, daopaisan_nakhonsawan: 190 } }, ctxNsw2).rateCls, undefined);
eq('fallback to towing rate shows the class used', n2({ destination: 'พิษณุโลก', gasStationsLiters: { jittamas_bangpain: 80 } }).rateCls, '18W');
// Codex P2: route has a 10W row but no 18W row → downhill must not borrow the 18W class-wide default.
const ctxNsw3 = buildAuditContext([...nswStd,
    { cls: '10W', origin: 'กรุงเทพ', destination: 'ตาก', variant: 'standard', up: { jittamas_bangpain: 60 }, down: 120, downVictor: 120, flags: [] }],
    [{ registrationNumber: '70-8901-8158', vehicleType: '18W', driveType: 'รถบริษัท' }]);
const n3 = auditTrip({ direction: 'ขาขึ้น', origin: 'HiQ', plate: '70-8901-8158', jobDetail: 'ตัดลูกนว.', destination: 'ตาก', gasStationsLiters: { jittamas_bangpain: 60, daopaisan_nakhonsawan: 170 } }, ctxNsw3);
eq('NSW drop, no towing row for the route → downhill not checked, no class default', [n3.expectedTotal, n3.flags.some(f => f.includes('ไม่ได้ตรวจขาล่อง'))], [60, true]);
eq('typo นครสรรค์ is a drop', n2({ jobDetail: 'ตัดลูก นครสรรค์', destination: 'กำแพงเพชร', gasStationsLiters: { jittamas_bangpain: 50, daopaisan_nakhonsawan: 190 } }).status, 'ok');
eq('no drop: full trailer rate', auditTrip({ direction: 'ขาขึ้น', origin: 'สหพัฒน์', plate: '70-8901-8158', destination: 'กำแพงเพชร', gasStationsLiters: { jittamas_bangpain: 50, daopaisan_nakhonsawan: 180 } }, ctxNsw).status, 'under');

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

// User decision 17 Nov 2026: pickups in ชลบุรี (สหพัฒน์ / น้ำพริกฉั่วฮะเส็ง) may fill เต็กย้ง up to 53 L (old fuel base).
const { DEFAULT_ORIGIN_EXTRAS, normalizeOriginExtras, matchOriginExtras } = await import('../src/fuelAudit.js');
const ctxOx = buildAuditContext(nswStd, [{ registrationNumber: '70-8901-8158', vehicleType: '18W', driveType: 'รถบริษัท' }], { originExtras: DEFAULT_ORIGIN_EXTRAS });
const ox = (origin, L) => auditTrip({ direction: 'ขาขึ้น', origin, plate: '70-8901-8158', destination: 'กำแพงเพชร', gasStationsLiters: L }, ctxOx);
eq('สหพัฒน์ + เต็กย้ง 53 → ok', [ox('สหพัฒน์', { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180, tekyong_chonburi: 53 }).status, ox('สหพัฒน์', { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180, tekyong_chonburi: 53 }).expectedTotal], ['ok', 313]);
eq('สหพัฒน์ without เต็กย้ง → ok (allowance, not required)', ox('สหพัฒน์', { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 }).status, 'ok');
eq('เต็กย้ง 553 → over +500', [ox('สหพัฒน์', { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180, tekyong_chonburi: 553 }).status, ox('สหพัฒน์', { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180, tekyong_chonburi: 553 }).diff], ['over', 500]);
eq('น้ำพริกฉั่วฮะเส็ง + 53 → ok', ox('น้ำพริกฉั่วฮะเส็ง', { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180, tekyong_chonburi: 53 }).status, 'ok');
eq('คอลเกต alone + เต็กย้ง → over (not in the rule)', ox('คอลเกต', { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180, tekyong_chonburi: 53 }).status, 'over');
eq('คอลเกต+สหพัฒน์ matches', ox('คอลเกต+สหพัฒน์', { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180, tekyong_chonburi: 53 }).status, 'ok');
eq('"สหพัฒน์ ลพ" excluded', matchOriginExtras('สหพัฒน์  ลพ', normalizeOriginExtras(DEFAULT_ORIGIN_EXTRAS)).length, 0);
eq('Bangkok pickup with เต็กย้ง → still over', ox('HiQ', { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180, tekyong_chonburi: 53 }).status, 'over');
eq('no rules passed → no extras (pure module default is empty)', auditTrip({ direction: 'ขาขึ้น', origin: 'สหพัฒน์', plate: '70-8901-8158', destination: 'กำแพงเพชร', gasStationsLiters: { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180, tekyong_chonburi: 53 } }, ctxNsw).status, 'over');
// review: a rule on a pump that is already in the route rate allows only liters BEYOND the rate
const ctxNis = buildAuditContext([{ cls: '18W', origin: 'นครสวรรค์', destination: 'กำแพงเพชร', variant: 'standard', up: { nissan_singburi: 80 }, down: 180, downVictor: 190, flags: [] }],
    [{ registrationNumber: '70-8901-8158', vehicleType: '18W', driveType: 'รถบริษัท' }], { originExtras: [{ label: 'x', keywords: ['นว'], pump: 'nissan_singburi', maxLiters: 30 }] });
const nis = (L) => auditTrip({ direction: 'ขาขึ้น', origin: 'นว', plate: '70-8901-8158', destination: 'กำแพงเพชร', gasStationsLiters: L }, ctxNis);
eq('rule on a pump already in the rate: exact rate fill = ok (not under -30)', [nis({ nissan_singburi: 80, daopaisan_nakhonsawan: 180 }).status, nis({ nissan_singburi: 80, daopaisan_nakhonsawan: 180 }).expectedTotal], ['ok', 260]);
eq('… rate + 30 extra = ok, rate + 40 = over +10', [nis({ nissan_singburi: 110, daopaisan_nakhonsawan: 180 }).status, nis({ nissan_singburi: 120, daopaisan_nakhonsawan: 180 }).diff], ['ok', 10]);
// review: ไพศาล cannot be a rule pump (its liters are split uphill/downhill)
eq('ไพศาล rule is dropped', normalizeOriginExtras([{ keywords: ['ก'], pump: 'daopaisan_nakhonsawan', maxLiters: 50 }]).length, 0);
// review: spaces ignored in matching
eq('exclude matches without the space ("สหพัฒน์ลพ")', matchOriginExtras('สหพัฒน์ลพ', normalizeOriginExtras(DEFAULT_ORIGIN_EXTRAS)).length, 0);
eq('exclude typed with two spaces still works', matchOriginExtras('สหพัฒน์ ลพ', normalizeOriginExtras([{ keywords: ['สหพัฒน์'], exclude: ['สหพัฒน์  ลพ'], pump: 'tekyong_chonburi', maxLiters: 53 }])).length, 0);
// review: downhill trips never get the allowance
eq('downhill trip: no origin allowance', auditTrip({ direction: 'ขาล่อง', origin: 'สหพัฒน์', plate: '70-8901-8158', destination: 'กรุงเทพ', gasStationsLiters: { tekyong_chonburi: 53 } }, ctxOx).flags.some(f => f.includes('ได้เพิ่มไม่เกิน')), false);
// review: KPP drop beyond กพ. (manual) counts the allowance in its diff
const ctxDropOx = buildAuditContext([...standards, ...dropStd], vehicles, { originExtras: DEFAULT_ORIGIN_EXTRAS });
const kppOx = auditTrip(trip({ origin: 'สหพัฒน์', destination: 'แม่สอด', jobDetail: 'ตัดลูก กพ', gasStationsLiters: { jittamas_bangpain: 210, tekyong_chonburi: 53 } }), ctxDropOx);
const kppNo = auditTrip(trip({ origin: 'สหพัฒน์', destination: 'แม่สอด', jobDetail: 'ตัดลูก กพ', gasStationsLiters: { jittamas_bangpain: 210, tekyong_chonburi: 53 } }), ctxDrop);
eq('KPP drop beyond กพ.: เต็กย้ง allowance taken off the diff', [kppOx.status, kppOx.diff, kppNo.diff - kppOx.diff], ['manual', kppNo.diff - 53, 53]);
// User confirmed 17 Nov 2026: sheet rule "พ่วง ลากหางเปล่า ไปตัดไว้ที่ กำแพง น้ำมันจ่ายเป็นรถสิบล้อ+เพิ่ม 20 ลิตร".
const { EMPTY_KPP_CLASS, EMPTY_KPP_DROP } = await import('../src/fuelAudit.js');
const ruleSheet = [
    ['สถานที่', '', 'รายการเติมน้ำมัน (พ่วง) ขาขึ้น', '', '', '', 'ขาล่อง'],
    HDR,
    ['จาก กรุงเทพ', 'กำแพงเพชร', 80, '***', '***', 80, 180, 'เบี้ยขับ 900'],
    ['          พ่วง ลากหางเปล่า ไปตัดไว้ที่ กำแพง น้ำมันจ่ายเป็นรถสิบล้อ+เพิ่ม 20  ลิตร', '', '***', '***', '', '', '180 (victor190)'],
];
const pr = parseRateSheet({ เหนือบน: ruleSheet });
const er = pr.standards.find(s => s.cls === EMPTY_KPP_CLASS);
eq('parser reads the empty-trailer rule from the sheet (+20, row 4)', [er?.extraLiters, er?.row, pr.standards.filter(s => s.cls === '18W').length], [20, 4, 1]);
const ctxE = buildAuditContext([...nswStd, er], [{ registrationNumber: '70-8901-8158', vehicleType: '18W', driveType: 'รถบริษัท' }]);
const em = (L, o = {}) => auditTrip({ direction: 'ขาขึ้น', origin: 'HiQ', plate: '70-8901-8158', destination: 'กำแพงเพชร', jobDetail: 'ลากหางเปล่า ไปตัดไว้ที่ กพ', gasStationsLiters: L, ...o }, ctxE);
eq('empty trailer left at กพ.: 10W 50 + 20 + downhill 180 = ok', [em({ jittamas_bangpain: 70, daopaisan_nakhonsawan: 180 }).status, em({ jittamas_bangpain: 70, daopaisan_nakhonsawan: 180 }).rateCls], ['ok', '10W']);
eq('… without the extra 20 → ok too (allowance)', em({ jittamas_bangpain: 50, daopaisan_nakhonsawan: 180 }).status, 'ok');
eq('… 50 + 30 → over +10', [em({ jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 }).status, em({ jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 }).diff], ['over', 10]);
eq('… not the KPP-drop manual check', em({ jittamas_bangpain: 70, daopaisan_nakhonsawan: 180 }, { jobDetail: 'ลากหางเปล่า ตัดลูก กพ' }).status, 'ok');
eq('template option "กำแพงเพชร (หางเปล่า)" uses the same rule', em({ jittamas_bangpain: 70, daopaisan_nakhonsawan: 180 }, { jobDetail: '', trailerDrop: EMPTY_KPP_DROP }).status, 'ok');
eq('"ลากลูกเปล่าไปขึ้นน้ำตาล" (towed empty to LOAD) is not this rule', em({ jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 }, { jobDetail: 'กำแพงเพชร//ลากลูกเปล่า**ไปขึ้นน้ำตาลต่อ' }).rateCls, undefined);
// review: rule not in the imported rate data yet → manual check (never silently the full towing rate)
const noRule = (o) => auditTrip({ direction: 'ขาขึ้น', origin: 'HiQ', plate: '70-8901-8158', destination: 'กำแพงเพชร', gasStationsLiters: { jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 }, ...o }, ctxNsw);
eq('rule missing from the rate data → manual, says to import the rate file', [noRule({ jobDetail: 'ลากหางเปล่า ไปตัดไว้ที่ กพ' }).status, noRule({ trailerDrop: EMPTY_KPP_DROP }).status, /นำเข้าไฟล์อัตราใหม่/.test(noRule({ trailerDrop: EMPTY_KPP_DROP }).reason)], ['manual', 'manual', true]);
// review: rule is written for พ่วง only
const ctxE6 = buildAuditContext([...nswStd, er,
    { cls: '6W', origin: 'กรุงเทพ', destination: 'กำแพงเพชร', variant: 'standard', up: { jittamas_bangpain: 50 }, down: 100, downVictor: 100, flags: [] },
    { cls: '6W_BOX', origin: 'กรุงเทพ', destination: 'กำแพงเพชร', variant: 'standard', up: { jittamas_bangpain: 70 }, down: 150, downVictor: 150, flags: [] }],
    [{ registrationNumber: '70-1-2', vehicleType: '6W_BOX', driveType: 'รถบริษัท' }]);
eq('6W_BOX with the empty-trailer text → manual (rule is for พ่วง)', auditTrip({ direction: 'ขาขึ้น', origin: 'HiQ', plate: '70-1-2', destination: 'กำแพงเพชร', jobDetail: 'ลากหางเปล่า ไปตัดไว้ที่ กพ', gasStationsLiters: { jittamas_bangpain: 70 } }, ctxE6).status, 'manual');
// review: the +20 may be filled at ไพศาล (on the way to กพ.)
const emP = em({ jittamas_bangpain: 50, daopaisan_nakhonsawan: 200 });
eq('empty trailer: extra 20 at ไพศาล → ok, downhill reading stays 180', [emP.status, emP.downActual, emP.downStatus], ['ok', 180, 'ok']);
// review: loaded trailer dropped, came back empty → NOT the empty-trailer rule
const back = em({ jittamas_bangpain: 80, daopaisan_nakhonsawan: 180 }, { jobDetail: 'ตัดลูก กพ ส่งตาก ขากลับลากลูกเปล่ากลับ', destination: 'ตาก' });
eq('"ตัดลูก กพ … ขากลับลากลูกเปล่า" → KPP-drop table check, not the empty-trailer rule', [back.status, back.reason.startsWith('เที่ยวถอดหางที่กำแพงเพชร')], ['manual', true]);
// review: empty trailer beyond กพ. with no 10-wheel rate there → manual (not the full towing rate)
eq('empty trailer beyond กพ., no 10W rate for the destination → manual', em({ jittamas_bangpain: 80 }, { destination: 'นครสวรรค์' }).status, 'manual');
// review: parser — merged cell (rule text in column B), Thai digits, unreadable number warned
const pr2 = parseRateSheet({ s: [ruleSheet[0], ruleSheet[1], ruleSheet[2], ['', 'พ่วง ลากหางเปล่า ไปตัดไว้ที่ กำแพง น้ำมันจ่ายเป็นรถสิบล้อ+เพิ่ม ๒๕ ลิตร']] });
eq('rule in column B with Thai digits → +25', pr2.standards.find(s => s.cls === EMPTY_KPP_CLASS)?.extraLiters, 25);
const pr3 = parseRateSheet({ s: [ruleSheet[0], ruleSheet[1], ruleSheet[2], ['พ่วง ลากหางเปล่า ไปตัดไว้ที่ กำแพง น้ำมันจ่ายเป็นรถสิบล้อ']] });
eq('rule without a number → warning, no rule', [pr3.standards.some(s => s.cls === EMPTY_KPP_CLASS), pr3.warnings.some(w => w.includes('ลากหางเปล่า'))], [false, true]);
// review: two rule rows with different extras → reported as a conflict, not "same"
const pr4 = parseRateSheet({ a: ruleSheet, b: [ruleSheet[0], ruleSheet[1], ['พ่วง ลากหางเปล่า ไปตัดไว้ที่ กำแพง +30 ลิตร']] });
eq('two empty-trailer rules with different extras → conflict flag', pr4.standards.filter(s => s.cls === EMPTY_KPP_CLASS).some(s => s.flags.some(f => f.startsWith('ขัดกับแถว'))), true);
// Driver aliases (user-confirmed file, 17 Nov 2026)
const { parseDriverAliasSheet, normalizeDriverAliases, applyDriverAlias } = await import('../src/fuelAudit.js');
const AH = ['ชื่อในกระดานคิว', 'จำนวนเที่ยว', 'ชื่อเต็มที่เสนอ', 'หลักฐาน', 'ความมั่นใจ', 'ยืนยัน'];
const MASTER = ['สกล แปลกชัยภูมิ', 'สุชาติ เรืองศิริ', 'วรวุฒิ รุ่งน้อย', 'สมชาย ใจดี', 'สมชาย มีสุข', 'วิชัย', 'วิชัย ทองดี'];
const aliasRows = [AH,
    ['สกล', 253, 'สกล แปลกชัยภูมิ', '', 'สูง', 'ใช่'],
    ['สุชาตื', 2, '', '', 'ไม่พบ', 'สุชาติ เรืองศิริ'],   // user typed the right name
    ['XX', 10, '', '', 'ไม่พบ', ''],                        // not confirmed
    ['วรวุฒิ', 5, 'วรวุฒิ รุ่งน้อย', '', 'กลาง', '✓'],
    ['ใคร', 1, 'ไม่มีในระบบ', '', 'ต่ำ', 'ใช่'],              // full name not in the master
    ['ปิยะ', 1, 'สกล แปลกชัยภูมิ', '', 'ต่ำ', 'ไม่ใช่'],       // rejected
];
const ar = parseDriverAliasSheet(aliasRows, MASTER);
eq('alias file: confirmed / typed name / skipped (blank + rejected) / conflicts', [ar.aliases, ar.skipped, ar.conflicts.map(c => c.from)],
    [[{ from: 'สกล', to: 'สกล แปลกชัยภูมิ' }, { from: 'สุชาตื', to: 'สุชาติ เรืองศิริ' }, { from: 'วรวุฒิ', to: 'วรวุฒิ รุ่งน้อย' }], ['XX', 'ปิยะ'], ['ใคร']]);
eq('not an alias file → error', !!parseDriverAliasSheet([['a', 'b']], MASTER).error, true);
eq('drivers master not loaded → error, nothing saved', [!!parseDriverAliasSheet(aliasRows, []).error, parseDriverAliasSheet(aliasRows, []).aliases.length], [true, 0]);
const amb = parseDriverAliasSheet([AH, ['สมชาย', 3, 'สมชาย ใจดี', '', 'สูง', 'ใช่'], ['สมชาย', 2, 'สมชาย มีสุข', '', 'สูง', 'ใช่']], MASTER);
eq('one board name confirmed as two people → conflict, not saved', [amb.aliases.length, amb.conflicts.length], [0, 2]);
eq('board name that is another driver\'s full name → conflict', parseDriverAliasSheet([AH, ['วิชัย', 3, 'วิชัย ทองดี', '', 'สูง', 'ใช่']], MASTER).conflicts.length, 1);
const chain = parseDriverAliasSheet([AH, ['ก', 1, 'สกล แปลกชัยภูมิ', '', '', 'ใช่'], ['สกล แปลกชัยภูมิ', 1, 'สุชาติ เรืองศิริ', '', '', 'ใช่']], MASTER);
eq('chain (a target is also a board name) → conflict', chain.conflicts.some(c => c.reason.includes('อีกทอด')), true);
eq('header below a title row is found', parseDriverAliasSheet([['จับคู่ชื่อคนขับ 17/11'], ...aliasRows], MASTER).aliases.length, 3);
eq('aliases from settings (object form; false = none) → map', [normalizeDriverAliases({ 0: { from: ' สกล ', to: 'สกล แปลกชัยภูมิ' }, 1: { from: 'x', to: 'x' }, 2: { from: '', to: 'y' } }), normalizeDriverAliases(false)], [{ 'สกล': 'สกล แปลกชัยภูมิ' }, {}]);
// the real alias function + carry-over + double-downhill pairing
const amap = normalizeDriverAliases(ar.aliases);
const tA = applyDriverAlias({ id: 'a', driver: ' สกล ', plate: '70-1' }, amap);
eq('applyDriverAlias: full name, board name kept, untouched trip returned as is', [tA.driver, tA.driverBoard, applyDriverAlias({ driver: 'อื่น' }, amap).driver], ['สกล แปลกชัยภูมิ', ' สกล ', 'อื่น']);
const resA = [{ status: 'over', actualTotal: 90, expectedTotal: 80, trip: applyDriverAlias({ driver: 'สกล' }, amap) }, { status: 'ok', actualTotal: 80, expectedTotal: 80, trip: { driver: 'สกล แปลกชัยภูมิ' } }];
eq('driver named like an Object property is not aliased', [applyDriverAlias({ driver: 'constructor' }, amap).driver, applyDriverAlias({ driver: 'toString' }, {}).driver], ['constructor', 'toString']);
eq('carry-over: board name and full name become one balance', applyCarry(resA), { 'สกล แปลกชัยภูมิ': 10 });
eq('rules from settings (object form, text keywords) are normalised; bad rules dropped',
    normalizeOriginExtras(Object.values({ 0: { label: 'x', keywords: 'ก, ข', pump: 'nissan_singburi', maxLiters: '20' }, 1: { keywords: [], pump: 'x', maxLiters: 5 }, 2: { keywords: ['ค'], pump: '', maxLiters: 5 } })),
    [{ label: 'x', keywords: ['ก', 'ข'], exclude: [], pump: 'nissan_singburi', maxLiters: 20 }]);

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
