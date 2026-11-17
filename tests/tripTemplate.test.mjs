// Run: node tests/tripTemplate.test.mjs   (no network; builds a real .xlsx in memory with ExcelJS and reads it with SheetJS)
import { createRequire } from 'module';
import {
    TEMPLATE_SHEET, LIST_SHEET, TEMPLATE_COLUMNS, buildTemplateLists, generateTemplate, findTemplateHeader,
    parseTemplateDate, parseTemplateRows, planTemplateSave, saveTemplateTrips, headOf,
} from '../src/tripTemplate.js';
import { buildAuditContext, auditTrip } from '../src/fuelAudit.js';

const require = createRequire(import.meta.url);
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

let fail = 0;
const eq = (name, a, b) => {
    const ok = JSON.stringify(a) === JSON.stringify(b);
    if (!ok) fail++;
    console.log(ok ? 'PASS' : 'FAIL', name, ok ? '' : `${JSON.stringify(a)} != ${JSON.stringify(b)}`);
};

const vehicles = [
    { registrationNumber: '70-8901', vehicleType: '10W' },
    { registrationNumber: '70-8901-8158', vehicleType: '18W', model: 'victor' },
    { registrationNumber: '70-7893', vehicleType: '10W' },
];
const drivers = [{ name: 'พีรพล' }, { name: 'สมชาย ใจดี' }];
const H = TEMPLATE_COLUMNS.map(c => c.header);
const col = (header) => H.indexOf(header);
const row = (o) => { const r = H.map(() => ''); for (const [k, v] of Object.entries(o)) { const i = col(k); if (i < 0) throw new Error('no column ' + k); r[i] = v; } return r; };
const opts = { startMonth: '2026-11', vehicles, drivers, fileName: 'test.xlsx', today: '2026-11-17' };
const parse = (rs) => parseTemplateRows([H, ...rs], opts);

// --- dates
eq('date: Thai text 17/11/2569', parseTemplateDate('17/11/2569'), '2026-11-17');
eq('date: 2-digit Thai year 17/11/69', parseTemplateDate('17/11/69'), '2026-11-17');
eq('date: CE 17/11/2026', parseTemplateDate('17/11/2026'), '2026-11-17');
eq('date: Excel serial 46343', parseTemplateDate(46343), '2026-11-17');
eq('date: SheetJS Date a few seconds before midnight', parseTemplateDate(new Date(2026, 10, 16, 23, 59, 50)), '2026-11-17');
eq('date: 31/02/2569 is invalid', parseTemplateDate('31/02/2569'), '');
eq('date: Buddhist-year Date object', parseTemplateDate(new Date(2569, 10, 17)), '2026-11-17');

// --- plates
eq('head of a pair plate', headOf('70-8901-8158'), '70-8901');
const lists = buildTemplateLists({ drivers, vehicles, rateStandards: [{ destination: 'กำแพงเพชร' }], routes: [{ origin: 'HiQ', destination: 'ตาก' }] });
eq('lists: heads / trailers from master plates', [lists.heads, lists.trailers], [['70-7893', '70-8901'], ['8158']]);

// --- parsing
const ok1 = parse([row({ 'วันที่': '17/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาขึ้น', 'คนขับ': 'พีรพล', 'ทะเบียนหัว': '70-8901', 'ทะเบียนหาง': '8158', 'ปลายทาง': 'กำแพงเพชร', 'ตัดหางที่': 'นครสวรรค์', 'ลิตร จิตตมาศ (บางประอิน)': 50, 'ลิตร ดาวไพศาล (นครสวรรค์)': '180', 'ค่าเที่ยว (บาท)': '1,500' })]).rows[0];
eq('valid row: plate joined, liters, money, drop', [ok1.status, ok1.trip.plate, ok1.trip.totalLiters, ok1.trip.revenue, ok1.trip.trailerDrop, ok1.trip.source], ['ok', '70-8901-8158', 230, 1500, 'นครสวรรค์', 'template']);
eq('full plate typed in the trailer cell is not doubled', parse([row({ 'วันที่': '17/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาขึ้น', 'คนขับ': 'พีรพล', 'ทะเบียนหัว': '70-8901', 'ทะเบียนหาง': '70-8901-8158', 'ปลายทาง': 'กำแพงเพชร' })]).rows[0].trip.plate, '70-8901-8158');
const bad = parse([row({ 'วันที่': '', 'ขาขึ้น/ขาล่อง': 'ขึ้น', 'ทะเบียนหัว': '', 'ลิตร จิตตมาศ (บางประอิน)': '50ลิตร' })]).rows[0];
eq('missing date / bad direction / no head / non-numeric liters → error, no key', [bad.status, bad.errors.length, bad.key], ['error', 4, '']);
eq('row before the start month is blocked', parse([row({ 'วันที่': '31/10/2569', 'ขาขึ้น/ขาล่อง': 'ขาขึ้น', 'ทะเบียนหัว': '70-7893' })]).rows[0].status, 'error');
eq('negative liters blocked', parse([row({ 'วันที่': '17/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาขึ้น', 'ทะเบียนหัว': '70-7893', 'ลิตร จิตตมาศ (บางประอิน)': -5 })]).rows[0].status, 'error');
const w = parse([row({ 'วันที่': '17/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาขึ้น', 'ทะเบียนหัว': '99-9999', 'คนขับ': 'ไม่รู้จัก', 'ปลายทาง': 'ตาก', 'รายการงาน': 'ตัดลูกนว.' })]).rows[0];
eq('unknown plate / driver and drop written only in text → warnings, still saved', [w.status, w.warnings.length, !!w.key], ['warn', 3, true]);
eq('blank lines are skipped', parse([H.map(() => ''), row({ 'วันที่': '17/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาล่อง', 'ทะเบียนหัว': '70-7893', 'คนขับ': 'พีรพล' })]).rows.length, 1);
eq('excel row number kept', parse([H.map(() => ''), row({ 'วันที่': '17/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาล่อง', 'ทะเบียนหัว': '70-7893', 'คนขับ': 'พีรพล' })]).rows[0].excelRow, 3);
eq('not a template → error message', !!parseTemplateRows([['a', 'b']], opts).error, true);
// header renamed "ทิศทาง" → "ขาขึ้น/ขาล่อง" (user, 17 Nov 2026): the new name is in the template, old files still import
eq('template header says ขาขึ้น/ขาล่อง', H.includes('ขาขึ้น/ขาล่อง') && !H.includes('ทิศทาง'), true);
const oldH = H.map(h => (h === 'ขาขึ้น/ขาล่อง' ? 'ทิศทาง' : h));
const oldRow = row({ 'วันที่': '17/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาล่อง', 'ทะเบียนหัว': '70-7893', 'คนขับ': 'พีรพล' });
const oldRes = parseTemplateRows([oldH, oldRow], opts);
eq('old template with header "ทิศทาง" still imports', [oldRes.error, oldRes.rows[0]?.trip.direction, oldRes.rows[0]?.status], ['', 'ขาล่อง', 'ok']);
eq('both old and new header in one file → clear error, nothing imported', (() => { const r = parseTemplateRows([[...H, 'ทิศทาง'], [...oldRow, 'ขาขึ้น']], opts); return [r.rows.length, r.error.includes('"ขาขึ้น/ขาล่อง" และ "ทิศทาง"')]; })(), [0, true]);
eq('empty direction → message names the new column', parse([row({ 'วันที่': '17/11/2569', 'ทะเบียนหัว': '70-7893', 'คนขับ': 'พีรพล' })]).rows[0].errors, ['ไม่ได้เลือก ขาขึ้น/ขาล่อง']);
eq('Firebase-safe: no undefined fields', Object.values(ok1.trip).some(v => v === undefined), false);

// --- keys and save plan
const same = [row({ 'วันที่': '17/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาขึ้น', 'คนขับ': 'พีรพล', 'ทะเบียนหัว': '70-7893', 'ปลายทาง': 'ตาก' })];
const k1 = parse(same).rows[0].key, k2 = parse(same).rows[0].key;
eq('same row → same key on re-import', k1 === k2 && k1.startsWith('tpl_2026-11-17_'), true);
const twice = parse([same[0], same[0]]).rows;
eq('two identical rows in one file get different keys', twice[0].key !== twice[1].key, true);
// review: a copy-pasted row must not be saved twice silently → the 2nd is "possible-dup" (in this file)
planTemplateSave(twice, {});
eq('identical rows in one file → 2nd is possible-dup from the file', [twice[0].importState, twice[1].importState, twice[1].dupOf.source, twice[1].dupOf.row], ['new', 'possible-dup', 'file', 2]);
const fixedFuel = parse([row({ 'วันที่': '17/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาขึ้น', 'คนขับ': 'พีรพล', 'ทะเบียนหัว': '70-7893', 'ปลายทาง': 'ตาก', 'ลิตร จิตตมาศ (บางประอิน)': 99 })]).rows;
eq('corrected liters keep the key (no second trip)', fixedFuel[0].key, k1);

const parsed = parse([...same, row({ 'วันที่': '17/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาล่อง', 'คนขับ': 'พีรพล', 'ทะเบียนหัว': '70-8901', 'ทะเบียนหาง': '8158' }), row({ 'วันที่': '18/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาขึ้น', 'คนขับ': 'พีรพล', 'ทะเบียนหัว': '70-7893', 'ปลายทาง': 'ตาก' })]).rows;
const existing = { '2026-11': {
    [k1]: { ...parsed[0].trip, gasStationsLiters: { jittamas_bangpain: 1 } },                                     // same key, different figures
    'tdh_2026-11-17_abc': { date: '2026-11-17', direction: 'ขาล่อง', plate: '70-8901-8158', destination: 'x' }, // same truck/day/direction from Transport
} };
planTemplateSave(parsed, existing);
eq('plan: already (changed) / possible-dup / new', parsed.map(p => [p.importState, p.changed]), [['already', true], ['possible-dup', false], ['new', false]]);
eq('possible-dup names the source', parsed[1].dupOf.source, 'transport');

// --- save path with a fake database: writes only absent keys, never overwrites
const store = JSON.parse(JSON.stringify(existing));
// Fake of the Firebase create-only transaction: the check and the write happen together on the "server".
let writes = 0;
const db = {
    createIfAbsent: async (ym, k, v) => { await null; if (store[ym]?.[k]) return false; (store[ym] ||= {})[k] = v; writes++; return true; },
};
const r1 = await saveTemplateTrips(db, parsed);
eq('save: only the new row is written', [r1.added, Object.keys(store['2026-11']).length], [1, 3]);
const r2 = await saveTemplateTrips(db, parsed);
eq('save again: the row now exists → skipped, nothing overwritten', [r2.added, r2.skipped], [0, 1]);
const r3 = await saveTemplateTrips(db, parsed, { includeKeys: [parsed[1].key] });
eq('save with the ticked possible-dup row adds only that row', r3.added, 1);
eq('the changed row was never overwritten', store['2026-11'][k1].gasStationsLiters, { jittamas_bangpain: 1 });
// Codex P2: two people importing the same file at the same moment → each trip written once
const race = parse([row({ 'วันที่': '20/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาขึ้น', 'คนขับ': 'พีรพล', 'ทะเบียนหัว': '70-7893', 'ปลายทาง': 'ตาก' })]).rows;
planTemplateSave(race, {});
writes = 0;
const [ra, rb] = await Promise.all([saveTemplateTrips(db, race), saveTemplateTrips(db, race)]);
eq('concurrent imports: one writes, the other skips', [ra.added + rb.added, ra.skipped + rb.skipped, writes], [1, 1, 1]);
// a failed write is reported, the others are still saved
const failDb = { createIfAbsent: async (ym, k) => { if (k === twice[0].key) throw new Error('network'); return true; } };
const rf = await saveTemplateTrips(failDb, twice, { includeKeys: [twice[1].key] });
eq('one failed write is counted, the other row still saved', [rf.added, rf.failed, rf.error.message], [1, 1, 'network']);

// review: Firebase returns pump keys sorted → an unchanged re-import must not show "changed"
const two = parse([row({ 'วันที่': '17/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาขึ้น', 'คนขับ': 'พีรพล', 'ทะเบียนหัว': '70-7893', 'ปลายทาง': 'ตาก', 'ลิตร จิตตมาศ (บางประอิน)': 50, 'ลิตร ดาวไพศาล (นครสวรรค์)': 180 })]).rows;
const fromDb = { ...two[0].trip, gasStationsLiters: { daopaisan_nakhonsawan: 180, jittamas_bangpain: 50 } };
planTemplateSave(two, { '2026-11': { [two[0].key]: fromDb } });
eq('unchanged multi-pump row re-imported → already, not changed', [two[0].importState, two[0].changed], ['already', false]);

// review: an error row fixed later must not take the key of a trip that was saved from the same file
const A = { 'วันที่': '17/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาขึ้น', 'คนขับ': 'พีรพล', 'ทะเบียนหัว': '70-7893', 'ปลายทาง': 'ตาก' };
const orig = parse([row({ ...A, 'ลิตร จิตตมาศ (บางประอิน)': 'abc' }), row({ ...A, 'คนขับ': 'สมชาย ใจดี' })]).rows;
const savedB = { '2026-11': { [orig[1].key]: orig[1].trip } };
const fixedA = parse([row({ ...A, 'ลิตร จิตตมาศ (บางประอิน)': 50 })]).rows;
planTemplateSave(fixedA, savedB);
eq('fixed error row gets its own key → shown as possible-dup to tick, not "already"', [fixedA[0].key !== orig[1].key, fixedA[0].importState], [true, 'possible-dup']);

// review: the same truck's 2nd run of the day, added to the file later, can be ticked on its own
const run1 = parse([row({ ...A })]).rows;
const day = { '2026-11': { [run1[0].key]: { ...run1[0].trip } } };
const both = parse([row({ ...A }), row({ ...A, 'ปลายทาง': 'กำแพงเพชร' })]).rows;
planTemplateSave(both, day);
eq('morning run already saved, afternoon run → possible-dup (from template), tickable', [both[0].importState, both[1].importState, both[1].dupOf.source], ['already', 'possible-dup', 'template']);

// --- audit uses the explicit trailerDrop field
const std = [
    { cls: '10W', origin: 'กรุงเทพ', destination: 'กำแพงเพชร', variant: 'standard', up: { jittamas_bangpain: 50 }, down: 120, downVictor: 120, flags: [] },
    { cls: '18W', origin: 'กรุงเทพ', destination: 'กำแพงเพชร', variant: 'standard', up: { jittamas_bangpain: 80 }, down: 180, downVictor: 190, flags: [] },
];
const ctx = buildAuditContext(std, vehicles);
eq('template NSW drop, victor truck filled 180 at ไพศาล (rate 190) → under', auditTrip(ok1.trip, ctx).status, 'under');
const tDrop = { ...ok1.trip, gasStationsLiters: { jittamas_bangpain: 50, daopaisan_nakhonsawan: 190 } };
eq('template NSW drop, victor 190 downhill → ok', [auditTrip(tDrop, ctx).status, auditTrip(tDrop, ctx).rateCls], ['ok', '10W']);
eq('"ไม่ตัด" overrides drop text in the job detail', auditTrip({ ...tDrop, trailerDrop: 'ไม่ตัด', jobDetail: 'ตัดลูกนว.', gasStationsLiters: { jittamas_bangpain: 80, daopaisan_nakhonsawan: 190 } }, ctx).rateCls, undefined);
// trailer left at กพ., head goes on beyond กพ. → special drop table → manual check
eq('trailerDrop กำแพงเพชร, destination beyond กพ. → manual check', auditTrip({ ...tDrop, destination: 'แม่สอด', trailerDrop: 'กำแพงเพชร' }, ctx).status, 'manual');
// user rule 17 Nov 2026: พ่วง to กพ. = 80 L → trip ending at กพ. is audited at the normal พ่วง rate (victor 80 + 190)
eq('trailerDrop กำแพงเพชร, destination กพ. itself → normal พ่วง rate', auditTrip({ ...tDrop, trailerDrop: 'กำแพงเพชร', gasStationsLiters: { jittamas_bangpain: 80, daopaisan_nakhonsawan: 190 } }, ctx).status, 'ok');

// --- round trip: real .xlsx built by ExcelJS → SheetJS (cellDates like the app) → parser
const wb = generateTemplate(ExcelJS, lists);
const ws = wb.getWorksheet(TEMPLATE_SHEET);
const put = (excelRow, o) => { for (const [k, v] of Object.entries(o)) ws.getCell(excelRow, col(k) + 1).value = v; };
// A date typed in Excel is a whole-day serial; ExcelJS writes Date values as UTC, so give it UTC midnight.
put(2, { 'วันที่': new Date(Date.UTC(2026, 10, 17)), 'ขาขึ้น/ขาล่อง': 'ขาขึ้น', 'คนขับ': 'พีรพล', 'ทะเบียนหัว': '70-8901', 'ทะเบียนหาง': '8158', 'ปลายทาง': 'กำแพงเพชร', 'ตัดหางที่': 'นครสวรรค์', 'ลิตร จิตตมาศ (บางประอิน)': 50, 'ลิตร ดาวไพศาล (นครสวรรค์)': 190 });
put(4, { 'วันที่': '17/11/2569', 'ขาขึ้น/ขาล่อง': 'ขาล่อง', 'คนขับ': 'พีรพล', 'ทะเบียนหัว': '70-7893', 'ลิตร ดาวไพศาล (นครสวรรค์)': 120 });
const buf = await wb.xlsx.writeBuffer();
const read = XLSX.read(buf, { type: 'buffer', cellDates: true });
eq('round trip: sheets', read.SheetNames, [TEMPLATE_SHEET, 'วิธีกรอก', LIST_SHEET]);
const sh = read.Sheets[TEMPLATE_SHEET];
const rows = XLSX.utils.sheet_to_json(sh, { header: 1, blankrows: true, defval: '' });
eq('round trip: header found', findTemplateHeader(rows), 0);
const rt = parseTemplateRows(rows, opts, XLSX.utils.decode_range(sh['!ref']).s.r + 1).rows;
eq('round trip: 2 rows, dates, excel rows', rt.map(p => [p.trip.date, p.excelRow, p.trip.totalLiters]), [['2026-11-17', 2, 240], ['2026-11-17', 4, 120]]);
eq('round trip: audit of the read-back row', auditTrip(rt[0].trip, ctx).status, 'ok');

// dropdowns must really be in the file (the reason ExcelJS is used)
const zipXml = await (async () => { const JSZip = require('jszip'); const z = await JSZip.loadAsync(buf); return z.file('xl/worksheets/sheet1.xml').async('string'); })();
const dvCount = (zipXml.match(/<dataValidation /g) || []).length;
eq('round trip: data validations written (list + number columns)', dvCount >= 8 + TEMPLATE_COLUMNS.filter(c => c.type === 'num').length, true);
eq('round trip: direction list is strict (stop), destination list is a warning', [/type="list"[^>]*sqref="B2:B3001"|sqref="B2:B3001"[^>]*type="list"/.test(zipXml), /errorStyle="warning"/.test(zipXml)], [true, true]);
eq('round trip: data sheet has no sample rows', XLSX.utils.sheet_to_json(XLSX.read(await generateTemplate(ExcelJS, lists).xlsx.writeBuffer(), { type: 'buffer' }).Sheets[TEMPLATE_SHEET], { header: 1, blankrows: false }).length, 1);

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
