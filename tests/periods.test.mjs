// Run: node tests/periods.test.mjs
import { closedOf, settingsForRestore, rateSaveClosedHit } from '../src/periods.js';

let fail = 0;
const eq = (name, a, b) => {
    const ok = JSON.stringify(a) === JSON.stringify(b);
    if (!ok) fail++;
    console.log(ok ? 'PASS' : 'FAIL', name, ok ? '' : `${JSON.stringify(a)} != ${JSON.stringify(b)}`);
};
const closed = { '2026-09': { closedAt: 'x' }, '2026-10': { closedAt: 'y' } };
eq('closed months among the written ones', [closedOf(['2026-11', '2026-10', '2026-10', ''], closed), closedOf(['2026-11'], closed), closedOf(['2026-10'], undefined)], [['2026-10'], [], []]);
eq('moving a trip from a closed month to an open one is caught (both months checked)', closedOf(['2026-10', '2026-11'], closed), ['2026-10']);
const fileSettings = { fuelAudit: { tolerance: 5 }, closedMonths: { '2026-08': {} }, periodLog: { a: 1 }, tripTemplate: { startMonth: '2026-12' } };
const current = { fuelAudit: { tolerance: 0 }, closedMonths: closed, periodLog: { b: 2 }, restoreLog: { r: 1 } };
eq('settings restore keeps the CURRENT lock state and logs', settingsForRestore(fileSettings, current),
    { fuelAudit: { tolerance: 5 }, closedMonths: closed, periodLog: { b: 2 }, tripTemplate: { startMonth: '2026-12' }, restoreLog: { r: 1 } });
eq('settings restore never re-closes months from an old file', settingsForRestore(fileSettings, { fuelAudit: {} }), { fuelAudit: { tolerance: 5 }, tripTemplate: { startMonth: '2026-12' } });
eq('rate start dates that would re-judge closed months', [rateSaveClosedHit('2026-10-15', closed), rateSaveClosedHit('2026-11-13', closed), rateSaveClosedHit('', closed)], [['2026-10'], [], ['2026-09', '2026-10']]);

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
