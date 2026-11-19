// Full backup / restore of the whole database (19 Nov 2026 — the old button saved only routes, fuel records and
// vehicles: no trips, drivers, rates, column-L liters or settings, and there was no way back).
// Pure functions (no DOM / Firebase): building the file, reading it back, and planning a restore.

export const BACKUP_APP = 'NeoSiamLogistics & Transport Fuel';
export const BACKUP_VERSION = 2;

/** Thai names of the database parts, for previews (unknown parts show their key). */
export const NODE_LABELS = {
    trips: 'เที่ยวรถ (คิวรถ)', fuelRateStandards: 'ตารางอัตราเติมน้ำมัน (ทุกฉบับ)', downhillFuel: 'ลิตรขาล่อง (คอลัมน์ L)',
    drivers: 'คนขับ', vehicles: 'ทะเบียนรถ', routes: 'เส้นทาง (ฐานเดิม)', fuelRecords: 'บันทึกน้ำมัน (เดิม)',
    rateTable: 'ตารางเบี้ยเลี้ยง', settings: 'การตั้งค่า', gasStations: 'ปั๊มน้ำมัน', vehicleTypes: 'ประเภทรถ',
    dataMigration: 'สถานะการย้ายข้อมูล',
};
// Nodes stored as { YYYY-MM: { key: record } } — counted record by record.
const BY_MONTH = new Set(['trips', 'downhillFuel']);

/** Number of records in one database part (month-partitioned parts count their records, not their months). */
export function countNode(name, value) {
    if (value == null) return 0;
    if (typeof value !== 'object') return 1;
    // null entries are holes of a sparse array (the SDK turns integer keys into arrays; JSON writes holes as null)
    const n = (o) => Object.values(o || {}).filter(x => x != null).length;
    if (BY_MONTH.has(name)) return Object.values(value).reduce((s, m) => s + (m && typeof m === 'object' ? n(m) : 0), 0);
    if (name === 'fuelRateStandards') return n(value.rows);
    return n(value);
}

export const countAll = (root) => Object.fromEntries(Object.entries(root || {}).map(([k, v]) => [k, countNode(k, v)]));

/** The backup file: the whole database plus who/when/how much, so a restore can show what it would bring back. */
export function buildBackup(root, exportedAt) {
    const data = JSON.parse(JSON.stringify(root || {})); // counted exactly as it will be read back from the file
    return { app: BACKUP_APP, version: BACKUP_VERSION, exportedAt, counts: countAll(data), data };
}

// Version 1 files (old button): { routes:[{id,…}], fuelRecords:[…], vehicles:[…] } → keyed objects.
const arrayToKeyed = (list) => Object.fromEntries((list || []).filter(x => x && x.id).map(({ id, ...rest }) => [id, rest]));

/**
 * A parsed JSON backup → { ok, version, exportedAt, data:{ node: value }, counts, error }.
 * Accepts this app's version 2 (whole database) and the old version 1 (3 parts only).
 */
export function readBackup(json) {
    if (!json || typeof json !== 'object' || json.app !== BACKUP_APP) return { ok: false, error: 'ไม่ใช่ไฟล์สำรองของระบบนี้' };
    if (json.version === BACKUP_VERSION) {
        if (!json.data || typeof json.data !== 'object') return { ok: false, error: 'ไฟล์สำรองไม่มีข้อมูล' };
        const counts = countAll(json.data);
        // the counts written at backup time must match what is in the file (a truncated/edited file is refused)
        const bad = Object.keys({ ...(json.counts || {}), ...counts }).filter(k => (json.counts || {})[k] !== counts[k]);
        if (bad.length) return { ok: false, error: `จำนวนข้อมูลในไฟล์ไม่ตรงกับที่บันทึกตอนสำรอง (${bad.join(', ')}) — ไฟล์อาจเสียหรือถูกแก้` };
        return { ok: true, version: 2, exportedAt: json.exportedAt || '', data: json.data, counts };
    }
    if (json.version === 1) {
        const data = {};
        for (const k of ['routes', 'fuelRecords', 'vehicles']) if (Array.isArray(json[k])) data[k] = arrayToKeyed(json[k]);
        return { ok: true, version: 1, exportedAt: json.exportedAt || '', data, counts: countAll(data) };
    }
    return { ok: false, error: `ไฟล์สำรองรุ่น ${json.version} — ระบบนี้อ่านไม่ได้` };
}

/**
 * Which parts a restore would replace: for each selected part, records in the file vs now.
 * Parts not in the file cannot be selected; a part is replaced WHOLE (records added since the backup are lost).
 */
export function planRestore(backup, currentCounts, selected) {
    const rows = Object.keys(backup.data).sort().map(name => ({
        name, label: NODE_LABELS[name] || name, inFile: backup.counts[name] || 0, now: (currentCounts || {})[name] || 0,
        selected: (selected || []).includes(name),
    }));
    return { rows, chosen: rows.filter(r => r.selected).map(r => r.name) };
}
