// Run: node tests/downhillFuel.test.mjs   (no network)
import crypto from 'crypto';
import { createRequire } from 'module';
import { parseColumnL, downhillTripKeys, planDownhillFuel, withDownhillFuel, reconcileColumnL, createAllIfAbsent } from '../src/downhillFuel.js';
import { parseDownhillQueue, detectSheetType } from '../src/importerTrips.js';
import { flagDoubleDownhill } from '../src/fuelAudit.js';

let fail = 0;
const eq = (name, a, b) => {
    const ok = JSON.stringify(a) === JSON.stringify(b);
    if (!ok) fail++;
    console.log(ok ? 'PASS' : 'FAIL', name, ok ? '' : `${JSON.stringify(a)} != ${JSON.stringify(b)}`);
};
const sha1 = async (s) => crypto.createHash('sha1').update(s).digest('hex');

// --- column L text
const f = parseColumnL('(คิว17.11.69นว.)180ไพศาล');
eq('filled: liters + uphill ref', [f.status, f.gasStationsLiters, f.totalLiters, f.uphillRef], ['filled', { daopaisan_nakhonsawan: 180 }, 180, { date: '2026-11-17', place: 'นว.' }]);
eq('XX → none', parseColumnL('(คิว15.07.69กพ.)XX').status, 'none');
eq('no "คิว" in the bracket', parseColumnL('(15.07.69กพ.)XX').uphillRef.date, '2026-07-15');
eq('missing ")" still reads the fill', parseColumnL('(คิว28.07.69ชม.120ไพศาล').totalLiters, 120);
eq('missing "(" still reads the fill', parseColumnL('คิว08.07.69พษ)40นิสสัน').gasStationsLiters, { nissan_singburi: 40 });
eq('two pumps', parseColumnL('(คิว01.11.69พษ.)100ไพศาล/80นิสสัน').totalLiters, 180);
eq('NGV', parseColumnL('(คิว01.11.69พษ.)NGV').status, 'ngv');
eq('waiting', parseColumnL('รอเติมขาล่อง').status, 'pending');
eq('partner truck', parseColumnL('รถร่วมจิระศักดิ์').status, 'partner');
eq('bracket only → blank', parseColumnL('(คิว07.07.69นว.)').status, 'blank');
eq('unknown pump typo is NOT guessed', parseColumnL('(คิว10.08.69นว.)60นิสนัส').status, 'unparsed');
eq('mistyped month → no uphill date, fill still read', [parseColumnL('(คิว28.098.68กพ.)50ไพศาล').uphillRef.date, parseColumnL('(คิว28.098.68กพ.)50ไพศาล').totalLiters], ['', 50]);
eq('empty → null', parseColumnL(''), null);
// review: missing ")" must not swallow a digit-less fill into the place name
eq('missing ")" + XX → none, place stays', [parseColumnL('(คิว07.07.69นว.XX').status, parseColumnL('(คิว07.07.69นว.XX').uphillRef.place], ['none', 'นว.']);
eq('missing ")" + NGV → ngv', parseColumnL('(คิว01.11.69พษ.NGV').status, 'ngv');
eq('place with a digit inside brackets', parseColumnL('(คิว27.10.69สาย 3)90ไพศาล').uphillRef.place, 'สาย 3');
// real rows (Transport 01): a digit in the place was read as the liters (180 ไพศาล → 3)
eq('"สาย3" place: liters are the fill, not the 3', parseColumnL('(คิว02.07.68สาย3)180ไพศาล/80นิสสัน').totalLiters, 260);
eq('"คลอง4" place', parseColumnL('(คิว26.08.69คลอง4)110ไพศาล/45นิสสัน').totalLiters, 155);
eq('3-digit year typo → blank uphill date, fill still read', [parseColumnL('(คิว26.09.669พช.)60นิสสัน').uphillRef.date, parseColumnL('(คิว26.09.669พช.)60นิสสัน').totalLiters], ['', 60]);
eq('"698ชม." typo → blank date, place ชม.', [parseColumnL('(คิว14.07.698ชม.)150ไพศาล').uphillRef.date, parseColumnL('(คิว14.07.698ชม.)150ไพศาล').totalLiters], ['', 150]);

// --- keys must equal the ones the Transport downhill import wrote (real key from production, 17 Nov 2026)
const real = { date: '2026-11-01', driver: 'พรชัย หุ่นประเสริฐ', plate: '70-8211-8212', origin: 'เกษตรไทย(เก้าเลี้ยว)', destination: 'นิวไทยที(นนทบุรี)', queueType: 'หนักล่อง', revenue: 1500, netTransfer: 6200 };
eq('key reproduces a real production key', (await downhillTripKeys([real], sha1))[0], 'tdh_2026-11-01_05b326e305');
const k2 = await downhillTripKeys([real, real], sha1);
eq('identical rows get different keys (occurrence)', k2[0] !== k2[1], true);

// --- parser returns column L in step with trips (blank/invalid rows skipped in both)
const sheet = [['h'], [new Date(2026, 10, 1), 'ก', 'หนักล่อง', '70-1111', '', 'A', 'B', '', '', '', '', '(คิว30.10.69นว.)180ไพศาล'], ['', '', '', '', '', '', '', '', '', '', '', 'orphan'], [new Date(2026, 10, 2), 'ข', '', '70-2222', '', 'A', 'B', '', '', '', '', 'XX']];
const pq = parseDownhillQueue(sheet);
eq('parser: lastQueue parallel to trips', [pq.trips.length, pq.lastQueue], [2, ['(คิว30.10.69นว.)180ไพศาล', 'XX']]);

// --- dates: SheetJS in Node gives "1 Nov 23:59:56" for the Excel date 2 Nov (serial 46328)
const { ymd } = await import('../src/importerTrips.js');
eq('date a few seconds before midnight = next day (what Excel shows)', ymd(new Date(2026, 10, 1, 23, 59, 56)), '2026-11-02');
eq('exact midnight unchanged', ymd(new Date(2026, 10, 2)), '2026-11-02');
eq('real file cell read like the app reads it', (() => {
    const X = createRequire(import.meta.url)('xlsx');
    const ws = X.utils.aoa_to_sheet([[46328]]); ws.A1.t = 'n'; ws.A1.z = 'd/m/yyyy';
    const wb = X.read(X.write({ SheetNames: ['s'], Sheets: { s: ws } }, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer', cellDates: true });
    return ymd(X.utils.sheet_to_json(wb.Sheets.s, { header: 1 })[0][0]);
})(), '2026-11-02');

// --- sheet detection: Transport 01.xlsx has a blank A1 header
const HDR1 = ['', 'คนขับ', 'ประเภทวิ่ง', 'ทะเบียน', 'ประเภทรถ', 'ขึ้นที่', 'ลงที่', 'จำนวน', 'ปั๊มไพศาล', 'ปั๊มนิสสัน', 'ปั๊มเอกวิทัศน์', 'คิวรถล่าสุด'];
eq('downhill sheet with blank A1 is detected', detectSheetType([HDR1]), 'downhill');
eq('old header with "วันที่" still detected', detectSheetType([['วันที่', ...HDR1.slice(1)]]), 'downhill');
eq('blank A1 without คิวรถล่าสุด stays unknown', detectSheetType([HDR1.slice(0, 11)]), 'unknown');

// --- plan: only trips in the system, never re-save
const trips = [{ ...real, yearMonth: '2026-11' }, { ...real, date: '2026-11-02', yearMonth: '2026-11' }, { ...real, date: '2026-11-03', yearMonth: '2026-11' }];
const keys = await downhillTripKeys(trips, sha1);
const plan = planDownhillFuel(trips, ['(คิว30.10.69นว.)180ไพศาล', 'XX', '(คิว01.11.69นว.)180ไพศาล'], keys,
    { '2026-11': { [keys[0]]: {}, [keys[1]]: {} } }, { '2026-11': { [keys[1]]: { raw: 'XX' } } }, { importedFrom: 'f.xlsx' });
eq('plan: save 1, already 1, not in system 1', [plan.toSave.length, plan.already, plan.noTrip, plan.changed], [1, 1, 1, 0]);
eq('plan record carries trip info + meta, no undefined', [plan.toSave[0].record.plate, plan.toSave[0].record.importedFrom, Object.values(plan.toSave[0].record).includes(undefined)], ['70-8211-8212', 'f.xlsx', false]);

// --- merge for the audit
const down = { id: 'tdh_x', direction: 'ขาล่อง', totalLiters: 0, totalBaht: 0, gasStationsLiters: {} };
const merged = withDownhillFuel(down, f);
eq('merge: downhill trip without liters takes column L', [merged.totalLiters, merged.fuelFrom, merged.id], [180, 'colL', 'tdh_x']);
eq('merge: original trip object untouched', down.totalLiters, 0);
eq('merge: trip with its own liters is not overridden', withDownhillFuel({ ...down, totalLiters: 50, gasStationsLiters: { nissan_singburi: 50 } }, f).totalLiters, 50);
eq('merge: XX record adds nothing', withDownhillFuel(down, parseColumnL('XX')), down);
eq('merge: uphill trip ignored', withDownhillFuel({ ...down, direction: 'ขาขึ้น' }, f).fuelFrom, undefined);

// --- create-only save: concurrent saves write once
const store = {}; let writes = 0;
const create = async (e) => { await null; if (store[e.key]) return false; store[e.key] = e; writes++; return true; };
const items = [{ key: 'a' }, { key: 'b' }];
const [s1, s2] = await Promise.all([createAllIfAbsent(items, create), createAllIfAbsent(items, create)]);
eq('concurrent saves: each key written once', [s1.added + s2.added, s1.skipped + s2.skipped, writes], [2, 2, 2]);
const sf = await createAllIfAbsent([{ key: 'x' }, { key: 'y' }], async (e) => { if (e.key === 'x') throw new Error('net'); return true; });
eq('a failed write is counted, the rest saved', [sf.added, sf.failed, sf.error.message], [1, 1, 'net']);

// --- double downhill: board writes a first name, Transport the full name → pair by head plate
const up = { leg: 'up', downhillRecorded: true, downActual: 180, flags: [], trip: { id: 'u', date: '2026-10-21', direction: 'ขาขึ้น', driver: 'สมชาย', plate: '71-1115-1116' } };
const dn = { leg: 'down', flags: [], actualTotal: 180, actual: { daopaisan_nakhonsawan: 180 }, trip: { id: 'd', date: '2026-10-23', direction: 'ขาล่อง', driver: 'สมชาย ใจดี', plate: '71-1115-1116' } };
flagDoubleDownhill([up, dn], [up.trip, dn.trip]);
eq('double downhill found by plate when names differ', [up.flags.length, dn.flags.length], [1, 1]);
const up2 = { ...up, flags: [], trip: { ...up.trip, plate: '70-9999' } }, dn2 = { ...dn, flags: [] };
flagDoubleDownhill([up2, dn2], [up2.trip, dn2.trip]);
eq('different truck and name → no flag', dn2.flags.length, 0);
const dnN = { leg: 'down', status: 'under', diff: -140, flags: [], actualTotal: 40, actual: { nissan_singburi: 40 }, trip: { id: 'd2', date: '2026-10-23', direction: 'ขาล่อง', driver: 'สมชาย ใจดี', plate: '71-1115-1116' } };
const upN = { ...up, flags: [] };
flagDoubleDownhill([upN, dnN], [upN.trip, dnN.trip]);
eq('own-liters downhill fill at another pump after the uphill ไพศาล → note only, status and diff unchanged (loose match = hint)', [dnN.flags.length, dnN.flags[0]?.startsWith('อาจเป็นการเติมเพิ่มขาล่อง'), dnN.status, dnN.diff], [1, true, 'under', -140]);
// review: a driver who changes trucks must not leave a stale pairing on the old truck
const upX = { leg: 'up', downhillRecorded: true, downActual: 180, flags: [], trip: { id: 'ux', date: '2026-10-20', direction: 'ขาขึ้น', driver: 'A', plate: '70-1111' } };
const upY = { leg: 'up', downhillRecorded: false, flags: [], trip: { id: 'uy', date: '2026-10-21', direction: 'ขาขึ้น', driver: 'A', plate: '70-2222' } };
const dnC = { leg: 'down', status: 'under', flags: [], actualTotal: 180, actual: { daopaisan_nakhonsawan: 180 }, trip: { id: 'dc', date: '2026-10-22', direction: 'ขาล่อง', driver: 'C', plate: '70-1111' } };
flagDoubleDownhill([upX, upY, dnC], [upX.trip, upY.trip, dnC.trip]);
eq('driver changed trucks → old truck no longer paired', dnC.flags.length, 0);
// review: a downhill trip in between under the OTHER key (other name, same truck) takes the uphill fill
const upB = { ...up, flags: [] };
const midB = { id: 'm', date: '2026-10-22', direction: 'ขาล่อง', driver: 'สมชาย ใจดี', plate: '71-1115-1116' };
const dnB = { leg: 'down', flags: [], actualTotal: 180, actual: { daopaisan_nakhonsawan: 180 }, trip: { id: 'd4', date: '2026-10-23', direction: 'ขาล่อง', driver: 'สมชาย', plate: '71-1115' } };
flagDoubleDownhill([upB, dnB], [upB.trip, midB, dnB.trip]);
eq('downhill trip in between under another name → no double flag', dnB.flags.length, 0);
// review: uphill-side message shows the ไพศาล liters, not the total
const upP = { ...up, flags: [] };
const dnP = { leg: 'down', flags: [], actualTotal: 180, actual: { daopaisan_nakhonsawan: 100, nissan_singburi: 80 }, trip: { ...dn.trip, id: 'd5' } };
flagDoubleDownhill([upP, dnP], [upP.trip, dnP.trip]);
eq('double flag on the uphill names ไพศาล 100, not the 180 total', upP.flags[0].includes('ไพศาล 100 ลิตร'), true);

// --- column L extra fill next to an uphill trip that already holds the downhill ไพศาล fill
const mk = (fill, ref) => ({ leg: 'down', status: 'wrong-pump', diff: -80, flags: [], actualTotal: 40, actual: fill,
    trip: { id: 'd3', date: '2026-10-23', direction: 'ขาล่อง', plate: '71-1115-1116', fuelFrom: 'colL', fuelText: '(คิว21.10.69นว.)40นิสสัน', fuelRef: { date: ref } } });
const upR = { leg: 'up', downhillRecorded: true, downActual: 180, flags: [], trip: { date: '2026-10-21', plate: '71-1115' } };
const x1 = mk({ nissan_singburi: 40 }, '2026-10-21');
reconcileColumnL([upR, x1], new Set(['2026-10']));
eq('extra นิสสัน after ไพศาล on the uphill row → manual (no rule), not wrong-pump', [x1.status, x1.diff, /เติมเพิ่ม 40/.test(x1.reason)], ['manual', 0, true]);
const x2 = mk({ nissan_singburi: 40 }, '2026-10-20');
reconcileColumnL([upR, x2], new Set(['2026-10']));
eq('uphill named in column L did not record downhill → audited as is', x2.status, 'wrong-pump');
const x3 = mk({ daopaisan_nakhonsawan: 180 }, '2026-10-21');
reconcileColumnL([upR, x3], new Set(['2026-10']));
eq('ไพศาล on both, column L names that uphill → "may be double" note on both, status unchanged', [x3.status, x3.flags.some(f => f.startsWith('อาจบันทึกขาล่องซ้ำ')), upR.flags.some(f => f.startsWith('อาจบันทึกขาล่องซ้ำ'))], ['wrong-pump', true, true]);
// full check 17 Nov 2026: column L names an EARLIER uphill; the truck's next uphill is the same day as this downhill
// → no double note (the old "most recent uphill" guess paired them: 55 false notes)
const upOld = { leg: 'up', downhillRecorded: true, downActual: 180, flags: [], trip: { id: 'u1', date: '2026-10-20', direction: 'ขาขึ้น', driver: 'K', plate: '70-5555' } };
const upNew = { leg: 'up', downhillRecorded: true, downActual: 180, flags: [], trip: { id: 'u2', date: '2026-10-24', direction: 'ขาขึ้น', driver: 'K', plate: '70-5555' } };
const dnRef = { leg: 'down', status: 'ok', flags: [], actualTotal: 180, actual: { daopaisan_nakhonsawan: 180 }, trip: { id: 'd9', date: '2026-10-24', direction: 'ขาล่อง', driver: 'K', plate: '70-5555', fuelFrom: 'colL', fuelRef: { date: '2026-10-18' } } };
// the uphill named by column L (18 Oct) exists but did not hold the downhill ไพศาล → two separate fills
const upRef18 = { leg: 'up', downhillRecorded: false, flags: [], trip: { id: 'u0', date: '2026-10-18', direction: 'ขาขึ้น', driver: 'K', plate: '70-5555' } };
const seqRef = [upRef18, upOld, upNew, dnRef];
reconcileColumnL(seqRef, new Set(['2026-10']));
flagDoubleDownhill(seqRef, seqRef.map(r => r.trip));
eq('column L names another uphill → no false "double" note on the same-day uphill', [dnRef.flags.length, upNew.flags.length], [0, 0]);
// review 17 Nov 2026: a ไพศาล downhill whose named uphill is missing (other plate on the board) is not silently passed
const dnLost = { leg: 'down', status: 'ok', flags: [], actualTotal: 180, actual: { daopaisan_nakhonsawan: 180 }, trip: { id: 'd10', date: '2026-10-24', direction: 'ขาล่อง', driver: 'K', plate: '70-6666', fuelFrom: 'colL', fuelRef: { date: '2026-10-20' } } };
reconcileColumnL([upOld, dnLost], new Set(['2026-10']));
eq('named uphill under another plate, same driver → "may be double" note on both', [dnLost.flags.some(f => f.startsWith('อาจบันทึกขาล่องซ้ำ') && f.includes('70-5555')), upOld.flags.some(f => f.includes('70-6666'))], [true, true]);
const dnNoFill = { ...dnLost, flags: [], status: 'ok', actual: { nissan_singburi: 60 }, trip: { ...dnLost.trip, id: 'd12' } };
reconcileColumnL([{ ...upOld, flags: [] }, dnNoFill], new Set(['2026-10']));
eq('same driver, other plate, no ไพศาล → no extra-fill verdict (needs the exact truck)', [dnNoFill.status, dnNoFill.flags.length], ['ok', 0]);
// the named uphill exists but did not hold the downhill ไพศาล → two separate fills, no note
const upPlain = { leg: 'up', downhillRecorded: false, flags: [], trip: { id: 'u3', date: '2026-10-20', direction: 'ขาขึ้น', driver: 'K', plate: '70-6666' } };
const dnOk = { ...dnLost, flags: [], trip: { ...dnLost.trip, id: 'd11' } };
reconcileColumnL([upPlain, dnOk], new Set(['2026-10']));
eq('named uphill found without downhill ไพศาล → no note', dnOk.flags.length, 0);
const x4 = mk({ nissan_singburi: 40 }, '2026-09-30');
reconcileColumnL([upR, x4], new Set(['2026-10']));
eq('uphill in a month not loaded → note, status unchanged', [x4.status, x4.flags.length], ['wrong-pump', 1]);
// review: a no-rate result keeps its own reason
const x5 = { ...mk({ nissan_singburi: 40 }, '2026-10-21'), status: 'no-rate', reason: 'ไม่พบทะเบียนในหน้าทะเบียนรถ' };
reconcileColumnL([upR, x5], new Set(['2026-10']));
eq('no-rate (unknown plate) is not turned into manual', [x5.status, x5.reason], ['no-rate', 'ไม่พบทะเบียนในหน้าทะเบียนรถ']);
// colL trip with an explicit uphill reference is left to reconcile, not re-paired by flagDoubleDownhill
const x6 = mk({ nissan_singburi: 40 }, '2026-10-20');
x6.trip.driver = 'สมชาย';
const upR2 = { ...upR, flags: [], trip: { ...upR.trip, id: 'u2', direction: 'ขาขึ้น', driver: 'สมชาย' } };
reconcileColumnL([upR2, x6], new Set(['2026-10']));
flagDoubleDownhill([upR2, x6], [upR2.trip, x6.trip]);
eq('colL naming another uphill → stays as audited', x6.status, 'wrong-pump');
// plan: byStatus counts only rows to save
eq('plan byStatus = rows to save', plan.byStatus, { filled: 1 });

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
