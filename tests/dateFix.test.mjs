// Run: node tests/dateFix.test.mjs   (no network)
import { planDateFix, buildDateFixBatches, recheckBatch, nextDay } from '../src/dateFix.js';

let fail = 0;
const eq = (name, a, b) => {
    const ok = JSON.stringify(a) === JSON.stringify(b);
    if (!ok) fail++;
    console.log(ok ? 'PASS' : 'FAIL', name, ok ? '' : `${JSON.stringify(a)} != ${JSON.stringify(b)}`);
};
const T = (date, o = {}) => ({ date, yearMonth: date.slice(0, 7), direction: 'ขาล่อง', driver: 'ก', plate: '70-1', destination: 'X', netTransfer: 1000, revenue: 1500, ...o });
// Same loop as handleDateFixApply: recheck against the current data, write, then mark the sources as vacated.
const apply = (db, batches) => {
    const committed = new Set();
    for (const b0 of batches) {
        const b = recheckBatch(b0, JSON.parse(JSON.stringify(db)), committed);
        for (const [p, v] of Object.entries(b.updates)) { const [ym, k] = p.split('/'); if (v === null) delete db[ym][k]; else (db[ym] ||= {})[k] = v; }
        for (const m of b.moves) committed.add(`${m.fromYm}/${m.fromKey}`);
    }
    return db;
};

eq('nextDay across month / year', [nextDay('2026-10-31'), nextDay('2026-12-31'), nextDay('2028-02-28')], ['2026-11-01', '2027-01-01', '2028-02-29']);

// basic: stored one day early → moved; same month keeps the key, month change moves it
const file = [T('2026-11-02'), T('2026-11-01', { driver: 'ข' })];
const db = { '2026-10': { '-a': T('2026-10-31', { driver: 'ข' }) }, '2026-11': { '-b': T('2026-11-01') } };
const p = planDateFix(file, db);
eq('plan: 2 shifts, 1 month change', [p.counts.shift, p.counts.monthChange], [2, 1]);
eq('moved trip keeps the old date in dateFixedFrom', p.moves.find(m => m.fromKey === '-a').trip, { ...T('2026-11-01', { driver: 'ข' }), dateFixedFrom: '2026-10-31' });

// daily-board / template / already fixed trips are never touched
const db2 = { '2026-11': { 'daily_01-11-69_r1': T('2026-11-01'), 'tpl_x': T('2026-11-01', { source: 'template' }), '-c': { ...T('2026-11-01'), dateFixedFrom: '2026-10-31' } } };
eq('daily / template / already-fixed trips skipped', planDateFix([T('2026-11-02')], db2).counts.checked, 0);

// correct date stays
eq('date already equal to the file → unchanged', planDateFix([T('2026-11-02')], { '2026-11': { '-d': T('2026-11-02') } }).counts, { checked: 1, shift: 0, same: 1, notInFile: 0, ambiguous: 0, conflict: 0, monthChange: 0 });
// not in file
eq('trip not in the file → unchanged', planDateFix([T('2026-11-02', { driver: 'zz' })], { '2026-11': { '-e': T('2026-11-01') } }).counts.notInFile, 1);

// consecutive identical trips: system {D-1, D} vs file {D, D+1} → both shift (group fits "early", not "exact")
const p3 = planDateFix([T('2026-11-02'), T('2026-11-03')], { '2026-11': { '-f': T('2026-11-01'), '-g': T('2026-11-02') } });
eq('consecutive identical trips both move', [p3.counts.shift, p3.counts.same], [2, 0]);
// fits both ways (system {D}, file {D, D+1}) → ambiguous, nothing moved
const p4 = planDateFix([T('2026-11-02'), T('2026-11-03')], { '2026-11': { '-h': T('2026-11-02') } });
eq('fits both readings → ambiguous, not moved', [p4.counts.shift, p4.counts.ambiguous], [0, 1]);

// tdh_ keys get the new key; missing map → ambiguous
const db5 = { '2026-11': { 'tdh_2026-11-01_old': T('2026-11-01') } };
eq('tdh key re-keyed', planDateFix([T('2026-11-02')], db5, { 'tdh_2026-11-01_old': 'tdh_2026-11-02_new' }).moves[0].toKey, 'tdh_2026-11-02_new');
eq('tdh key without map → ambiguous', planDateFix([T('2026-11-02')], db5, {}).counts.ambiguous, 1);

// conflict: target occupied by a trip that is not moving
const db6 = { '2026-11': { 'tdh_2026-11-01_old': T('2026-11-01'), 'tdh_2026-11-02_new': T('2026-11-02', { driver: 'other' }) } };
const p6 = planDateFix([T('2026-11-02'), T('2026-11-02', { driver: 'other' })], db6, { 'tdh_2026-11-01_old': 'tdh_2026-11-02_new' });
eq('target held by a trip that stays → conflict, not moved', [p6.counts.conflict, p6.moves.length], [1, 0]);

// chain: A's new key = B's old key, B moves too (same month and across a month end) → nothing lost
const chainDb = {
    '2026-10': { 'tdh_2026-10-31_A': T('2026-10-31', { driver: 'A' }) },
    '2026-11': { 'tdh_2026-11-01_B': T('2026-11-01', { driver: 'B' }), 'tdh_2026-11-02_C': T('2026-11-02', { driver: 'C' }) },
};
const chainFile = [T('2026-11-01', { driver: 'A' }), T('2026-11-02', { driver: 'B' }), T('2026-11-03', { driver: 'C' })];
const chainMap = { 'tdh_2026-10-31_A': 'tdh_2026-11-01_B', 'tdh_2026-11-01_B': 'tdh_2026-11-02_C', 'tdh_2026-11-02_C': 'tdh_2026-11-03_D' };
const pc = planDateFix(chainFile, chainDb, chainMap);
const movingFrom = new Set(pc.moves.map(m => `${m.fromYm}/${m.fromKey}`));
const after = apply(JSON.parse(JSON.stringify(chainDb)), buildDateFixBatches(pc.moves));
eq('chain of re-keys across a month end: all three trips present with true dates',
    Object.entries(after['2026-11']).map(([k, t]) => `${k}:${t.driver}:${t.date}`).sort(),
    ['tdh_2026-11-01_B:A:2026-11-01', 'tdh_2026-11-02_C:B:2026-11-02', 'tdh_2026-11-03_D:C:2026-11-03']);
eq('chain: old month emptied', Object.keys(after['2026-10']).length, 0);
eq('batches go from the latest month backwards', buildDateFixBatches(pc.moves).map(b => b.month), ['2026-11', '2026-10']);

// recheck: a trip edited after the preview is dropped, not overwritten
const pe = planDateFix([T('2026-11-02')], { '2026-11': { '-k': T('2026-11-01') } });
const edited = { '2026-11': { '-k': { ...T('2026-11-01'), netTransfer: 999 } } };
const rb = recheckBatch(buildDateFixBatches(pe.moves)[0], edited, new Set());
eq('trip changed since the preview → dropped', [rb.dropped.length, Object.keys(rb.updates).length], [1, 0]);
const rb2 = recheckBatch(buildDateFixBatches(pe.moves)[0], { '2026-11': { '-k': T('2026-11-01') } }, new Set());
eq('unchanged trip → same-key update keeps the trip (no delete)', rb2.updates, { '2026-11/-k': { ...T('2026-11-02'), dateFixedFrom: '2026-11-01' } });
const rb3 = recheckBatch(buildDateFixBatches(pe.moves)[0], { '2026-11': {} }, new Set());
eq('trip deleted since the preview → dropped', rb3.dropped.length, 1);

// review: a move dropped at write time must not leave its old path counted as vacated
const chainDb2 = { '2026-10': { 'tdh_2026-10-05_h1': T('2026-10-05', { driver: 'H' }), 'tdh_2026-10-06_h2': T('2026-10-06', { driver: 'H' }) } };
const map2 = { 'tdh_2026-10-05_h1': 'tdh_2026-10-06_h2', 'tdh_2026-10-06_h2': 'tdh_2026-10-07_h3' };
const pc2 = planDateFix([T('2026-10-06', { driver: 'H' }), T('2026-10-07', { driver: 'H' })], chainDb2, map2);
const editedH2 = JSON.parse(JSON.stringify(chainDb2)); editedH2['2026-10']['tdh_2026-10-06_h2'].netTransfer = 5555; // edited after the preview
const rbH = recheckBatch(buildDateFixBatches(pc2.moves)[0], editedH2, new Set());
eq('edited occupant stays → the move onto it is dropped too (nothing overwritten)', [rbH.moves.length, rbH.dropped.length, Object.keys(rbH.updates).length], [0, 2, 0]);
// review: a target held by a trip that ends up NOT moving (no key map) is a conflict at plan time
const pc3 = planDateFix([T('2026-10-06', { driver: 'H' }), T('2026-10-07', { driver: 'H' })], chainDb2, { 'tdh_2026-10-05_h1': 'tdh_2026-10-06_h2' });
eq('occupant without a key map → conflict shown in the preview, not a silent drop', [pc3.counts.shift, pc3.counts.conflict, pc3.counts.ambiguous], [0, 1, 1]);
// committed set: a target vacated by an earlier written batch is free
const crossDb = { '2026-10': { 'tdh_2026-10-31_A': T('2026-10-31', { driver: 'A' }) }, '2026-11': { 'tdh_2026-11-01_B': T('2026-11-01', { driver: 'B' }) } };
const crossMap = { 'tdh_2026-10-31_A': 'tdh_2026-11-01_B', 'tdh_2026-11-01_B': 'tdh_2026-11-02_C' };
const pcx = planDateFix([T('2026-11-01', { driver: 'A' }), T('2026-11-02', { driver: 'B' })], crossDb, crossMap);
const [bNov, bOct] = buildDateFixBatches(pcx.moves);
eq('later month not written yet → move into it is blocked', recheckBatch(bOct, crossDb, new Set()).dropped.length, 1);
eq('after the later month is written → move into it is allowed', recheckBatch(bOct, crossDb, new Set(['2026-11/tdh_2026-11-01_B'])).moves.length, 1);
eq('batch order Nov then Oct', [bNov.month, bOct.month], ['2026-11', '2026-10']);

// running the plan again after the fix moves nothing
eq('idempotent: second plan moves nothing', planDateFix(chainFile, after, chainMap).counts.shift, 0);

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
