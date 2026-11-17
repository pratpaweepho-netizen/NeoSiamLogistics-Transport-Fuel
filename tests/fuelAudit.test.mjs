// Run: node tests/fuelAudit.test.mjs   (no network, no Excel file needed)
import { parseRateSheet, buildAuditContext, auditTrip, canonDestination } from '../src/fuelAudit.js';

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
eq('downhill trip skipped', run({ direction: 'ขาล่อง', gasStationsLiters: { jittamas_bangpain: 80 } }), null);
eq('no fuel skipped', run({}), null);
eq('unknown plate → no-rate', run({ plate: '99-0000', gasStationsLiters: { jittamas_bangpain: 80 } }).status, 'no-rate');
eq('destination outside sheet → no-rate', run({ destination: 'สาย 3', gasStationsLiters: { jittamas_bangpain: 80 } }).status, 'no-rate');
eq('NSW branch origin', run({ origin: 'ย่อย นว', destination: 'กำแพงเพชร', gasStationsLiters: { nissan_singburi: 80 } }).origin, 'นครสวรรค์');
eq('other branch origin → no-rate', run({ origin: 'สาขา ชม', gasStationsLiters: { jittamas_bangpain: 80 } }).status, 'no-rate');

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
