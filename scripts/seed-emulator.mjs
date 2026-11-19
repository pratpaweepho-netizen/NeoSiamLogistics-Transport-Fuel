// Copy data into the LOCAL Firebase emulator (never into production).
//   npm run emu            (terminal 1: starts the emulator on 127.0.0.1:9000)
//   npm run emu:seed       (copies the production data — READ only — into the emulator)
//   npm run emu:seed -- --from-backup "path/to/NeoSiamFuel_สำรองทั้งระบบ_….json"   (from a full backup file instead)
//   npm run dev:emu        (the app, using the emulator; a yellow bar says so)
import fs from 'fs';
import { countAll } from '../src/backup.js';

const PROD = 'https://fleet-fuel-management-system-default-rtdb.asia-southeast1.firebasedatabase.app';
const NS = 'fleet-fuel-management-system-default-rtdb';
const EMU = `http://127.0.0.1:9000`;

const target = new URL(`${EMU}/.json?ns=${NS}`);
if (!['127.0.0.1', 'localhost'].includes(target.hostname)) throw new Error('refusing: the target is not the local emulator');

const i = process.argv.indexOf('--from-backup');
let data;
if (i > 0) {
    const file = JSON.parse(fs.readFileSync(process.argv[i + 1], 'utf8'));
    if (!file || file.version !== 2 || !file.data) throw new Error('not a full (version 2) backup file');
    data = file.data;
    console.log('source: backup file', process.argv[i + 1]);
} else {
    const r = await fetch(`${PROD}/.json`); // GET only
    if (!r.ok) throw new Error(`reading production failed: ${r.status}`);
    data = await r.json();
    console.log('source: production (read only)');
}
try { await fetch(`${EMU}/`); } catch { throw new Error('the emulator is not running — start it with: npm run emu'); }
const res = await fetch(target, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) });
if (!res.ok) throw new Error(`writing to the emulator failed: ${res.status} ${await res.text()}`);
console.log('emulator seeded (records, counted like the backup):', JSON.stringify(countAll(data)));
