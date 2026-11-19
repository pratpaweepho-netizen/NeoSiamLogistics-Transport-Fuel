// Run: node tests/backup.test.mjs
import { buildBackup, readBackup, planRestore, countNode, BACKUP_APP } from '../src/backup.js';

let fail = 0;
const eq = (name, a, b) => {
    const ok = JSON.stringify(a) === JSON.stringify(b);
    if (!ok) fail++;
    console.log(ok ? 'PASS' : 'FAIL', name, ok ? '' : `${JSON.stringify(a)} != ${JSON.stringify(b)}`);
};
const root = {
    trips: { '2026-10': { a: {}, b: {} }, '2026-11': { c: {} } },
    downhillFuel: { '2026-10': { x: {} } },
    fuelRateStandards: { meta: {}, rows: { r1: {}, r2: {}, r3: {} }, history: { h: {} } },
    drivers: { d1: {}, d2: {} },
    settings: { fuelAudit: {}, tripTemplate: {} },
};
eq('counts: trips per record, rates per row', [countNode('trips', root.trips), countNode('fuelRateStandards', root.fuelRateStandards), countNode('drivers', root.drivers), countNode('x', null)], [3, 3, 2, 0]);
const file = JSON.parse(JSON.stringify(buildBackup(root, '2026-11-19T10:00:00Z')));
eq('backup holds every part with counts', [file.app, file.version, Object.keys(file.data), file.counts], [BACKUP_APP, 2, ['trips', 'downhillFuel', 'fuelRateStandards', 'drivers', 'settings'], { trips: 3, downhillFuel: 1, fuelRateStandards: 3, drivers: 2, settings: 2 }]);
const back = readBackup(file);
eq('reads its own file back', [back.ok, back.version, back.counts.trips], [true, 2, 3]);
const cut = JSON.parse(JSON.stringify(file)); delete cut.data.trips['2026-11'];
eq('a file that lost records is refused', [readBackup(cut).ok, /trips/.test(readBackup(cut).error)], [false, true]);
eq('not this app / unknown version refused', [readBackup({ app: 'x' }).ok, readBackup({ app: BACKUP_APP, version: 9 }).ok, readBackup(null).ok], [false, false, false]);
const v1 = readBackup({ app: BACKUP_APP, version: 1, routes: [{ id: 'r1', a: 1 }], fuelRecords: [], vehicles: [{ id: 'v1' }, { noId: 1 }] });
eq('old version-1 file (3 parts) is read', [v1.ok, v1.version, v1.data.routes, v1.counts], [true, 1, { r1: { a: 1 } }, { routes: 1, fuelRecords: 0, vehicles: 1 }]);
const plan = planRestore(back, { trips: 5, drivers: 2 }, ['trips', 'nope']);
eq('restore plan: only parts in the file, chosen ones flagged', [plan.chosen, plan.rows.find(r => r.name === 'trips')], [['trips'], { name: 'trips', label: 'เที่ยวรถ (คิวรถ)', inFile: 3, now: 5, selected: true }]);

// review 19 Nov 2026: integer-keyed parts come back from the SDK as sparse arrays (holes → null in JSON)
const sparse = []; sparse[0] = { n: 'a' }; sparse[2] = { n: 'b' }; sparse[3] = { n: 'c' };
const sb = JSON.parse(JSON.stringify(buildBackup({ gasStations: sparse }, 'T')));
eq('sparse array: counted without holes, own file accepted', [sb.counts.gasStations, readBackup(sb).ok], [3, true]);

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
