// Run: node tests/backup.test.mjs
import { buildBackup, readBackup, planRestore, countNode, BACKUP_APP, isRawDatabase, stampFromFileName, fileToText, emptyPartsOverData } from '../src/backup.js';

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

// NAS nightly copy / Firebase export = the bare database root (19 Nov 2026)
{
    const raw = { ...JSON.parse(JSON.stringify(root)), settings: { fuelAudit: { tolerance: 0 } } };
    const r = readBackup(raw, 'fleet-fuel_2026-11-20_0200.json.gz');
    eq('bare database copy accepted, counted, dated from the file name', [r.ok, r.version, r.counts.trips, r.exportedAt, r.data === raw], [true, 'raw', 3, '2026-11-20T02:00', true]);
    eq('file names without a time / without a date', [stampFromFileName('fleet-fuel-export-2026-11-20.json'), stampFromFileName('backup.json'), readBackup(raw).exportedAt], ['2026-11-20', '', '']);
    eq('not a database copy → refused', [
        readBackup({ hello: 1 }).ok,                                      // no trips part
        readBackup({ trips: {}, 'bad key': 1 }).ok,                       // a key the database cannot have
        readBackup([{ trips: {} }]).ok,                                   // array
        readBackup({ trips: 5 }).ok,                                      // trips not an object
        readBackup({ app: 'other', trips: {} }).ok,                       // another app's wrapped file
        isRawDatabase(buildBackup(root, 'T')),                            // our own v2 file goes the checked way
        readBackup({ trips: {}, settings: {} }).ok,                       // empty trips (Codex P2: would wipe trips)
        readBackup({ trips: { orders: { a: {} } }, settings: {} }).ok,    // another app's "trips" (not by month)
        readBackup({ trips: { '2026-10': { a: {} } } }).ok,               // no settings part → not this database
        readBackup({ trips: { '2026-10': 3 }, settings: {} }).ok,         // a month that holds no records
        readBackup({ trips: { '2026-10': {} }, settings: {} }).ok,        // months but not one trip (Codex round 2)
        readBackup({ trips: { '2026-10': { a: {} } }, settings: {}, orders: {} }).ok, // a part this database has not
    ], [false, false, false, false, false, false, false, false, false, false, false, false]);
    eq('every part of the real database is known (else NAS copies would be refused)', ['fuelRecords', 'drivers', 'rateTable', 'routes', 'downhillFuel', 'vehicles', 'dataMigration', 'gasStations', 'trips', 'fuelRateStandards', 'settings', 'vehicleTypes'].filter(k => k !== 'trips' && k !== 'settings').filter(k => !isRawDatabase({ trips: { '2026-10': { a: {} } }, settings: {}, [k]: {} })), []);
    eq('restore refuses an empty part over existing records (any file version)', [
        emptyPartsOverData([{ label: 'คนขับ', selected: true, inFile: 0, now: 52 }, { label: 'ทะเบียนรถ', selected: true, inFile: 460, now: 461 }, { label: 'ปั๊ม', selected: false, inFile: 0, now: 7 }, { label: 'ประเภทรถ', selected: true, inFile: 0, now: 0 }]),
    ], [['คนขับ']]);
    // .gz from the NAS is unpacked in the browser; plain JSON passes through
    const { gzipSync } = await import('zlib');
    const text = JSON.stringify(raw);
    const gz = new Blob([gzipSync(Buffer.from(text, 'utf8'))]);
    eq('fileToText: .gz unpacked, plain JSON unchanged (Thai kept)', [await fileToText(gz) === text, await fileToText(new Blob(['{"ก":1}'])), JSON.parse(await fileToText(new Blob([gzipSync(Buffer.from('{"ก":1}'))]))).ก], [true, '{"ก":1}', 1]);
    let err = '';
    try { await fileToText(new Blob([new Uint8Array([0x1f, 0x8b, 1, 2, 3])])); } catch (e) { err = 'thrown'; }
    eq('a broken .gz is an error, not an empty restore', err, 'thrown');
}

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
