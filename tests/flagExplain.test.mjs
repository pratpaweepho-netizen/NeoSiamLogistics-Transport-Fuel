// Run: node tests/flagExplain.test.mjs
import { readFileSync } from 'fs';
import { STATUS_EXPLAIN, DOWN_EXPLAIN, FLAG_EXPLAIN, explainFlag, explainOneLine, todoForRow } from '../src/flagExplain.js';
import { AUDIT_STATUS, DOWN_STATUS } from '../src/fuelAudit.js';

let fail = 0;
const eq = (name, a, b) => {
    const ok = JSON.stringify(a) === JSON.stringify(b);
    if (!ok) fail++;
    console.log(ok ? 'PASS' : 'FAIL', name, ok ? '' : `${JSON.stringify(a)} != ${JSON.stringify(b)}`);
};
const full = (e) => !!(e && e.what && e.why && e.todo?.length);

eq('every audit status has what / why / steps', Object.keys(AUDIT_STATUS).filter(k => !full(STATUS_EXPLAIN[k])), []);
eq('every downhill result except "ตรง" has what / why / steps', Object.keys(DOWN_STATUS).filter(k => k !== 'ok' && !full(DOWN_EXPLAIN[k])), []);
eq('every flag type is complete', FLAG_EXPLAIN.filter(f => !full(f) || !f.title).map(f => f.key), []);
// a suspicion is never written as certain
const allText = JSON.stringify([STATUS_EXPLAIN, DOWN_EXPLAIN, FLAG_EXPLAIN.map(({ test, ...e }) => e)]);
eq('no "ผิดแน่" / "โกง" wording', [allText.includes('ผิดแน่'), allText.includes('โกง')], [false, false]);

// the texts fuelAudit.js really writes (copied from its templates) are recognised
const keyOf = (t) => explainFlag(t)?.key || null;
eq('real flag texts → explanation', [
    keyOf('อาจบันทึกขาล่องซ้ำ: เที่ยวขาขึ้น 2026-11-02 บันทึกไพศาลขาล่อง 150 ลิตรไว้แล้ว'),
    keyOf('อาจเป็นการเติมเพิ่มขาล่อง: เที่ยวขาขึ้น 2026-11-02 บันทึกไพศาลขาล่อง 150 ลิตรไว้แล้ว — ตรวจเอง'),
    keyOf('เติมปั๊มอื่นแทนไพศาล: ดาวอุไร 40 ลิตร'),
    keyOf('เติมปั๊มที่ไม่อยู่ในอัตรา: PTT 30 ลิตร'),
    keyOf('เติมเป็นเงิน 1,000 บาท (ไม่มีอัตราเป็นเงิน)'),
    keyOf('กระดานคิวเขียนแค่ทะเบียนหัว แต่ลิตรรวมตรงกับอัตราพ่วงของ 70-1234 — ตรวจแบบลากหาง (ควรเขียนเลขหางในกระดาน)'),
    keyOf('ขาล่องไพศาล 150 ลิตร ไม่ได้บันทึกในเที่ยวนี้'),
    keyOf('อัตราแถวนี้มีข้อสังเกต: แถวซ้ำ (ค่าเท่ากัน)'),
], ['double-down', 'down-topup', 'off-pump', 'off-pump', 'baht-only', 'head-only', 'down-missing', 'rate-note']);
eq('information-only notes have no explanation', [keyOf('ลากหางเปล่า: ได้เพิ่มไม่เกิน 20 ลิตร ตามตารางอัตรา (x แถว 3)'), keyOf('เติมเพิ่มตามต้นทาง: ดาวอุไร ได้เพิ่มไม่เกิน 30 ลิตร (นับ 20)'), keyOf(''), keyOf(undefined)], [null, null, null, null]);

// guard against the source texts drifting: each recognised prefix still exists in fuelAudit.js
const src = readFileSync(new URL('../src/fuelAudit.js', import.meta.url), 'utf8');
eq('flag texts still written by fuelAudit.js', ['อาจบันทึกขาล่องซ้ำ', 'อาจเป็นการเติมเพิ่มขาล่อง', 'เติมปั๊มอื่นแทนไพศาล', 'เติมปั๊มที่ไม่อยู่ในอัตรา', 'เติมเป็นเงิน', 'กระดานคิวเขียนแค่ทะเบียนหัว', 'ไม่ได้บันทึกในเที่ยวนี้', 'อัตราแถวนี้มีข้อสังเกต', 'ตารางอัตรายังไม่มีกฎการเติมเพิ่มขาล่อง'].filter(p => !src.includes(p)), []);

eq('one line for tooltip / Excel', explainOneLine({ what: 'ก', why: 'ข', todo: ['ค', 'ง'] }), 'ก — ข · ต้องทำ: 1) ค 2) ง');
eq('row steps: none for a clean trip', todoForRow({ status: 'ok', flags: [] }), []);
const steps = todoForRow({ status: 'over', downStatus: 'missing', reason: '', flags: ['ขาล่องไพศาล 150 ลิตร ไม่ได้บันทึกในเที่ยวนี้', 'อาจบันทึกขาล่องซ้ำ: x'] });
eq('row steps: status first, then flags, no repeats', [steps[0], steps.length, new Set(steps).size === steps.length],
    [STATUS_EXPLAIN.over.todo[0], STATUS_EXPLAIN.over.todo.length + DOWN_EXPLAIN.missing.todo.length + FLAG_EXPLAIN[0].todo.length, true]);

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
