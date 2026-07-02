// Phase 3 — parse master data: คนขับ (drivers + bank account) and ฐานคำนวน (per-diem rate table).
// Pure functions (no DOM/Firebase) so they can be unit-tested in Node.
import { parseNum, clean } from './parseUtils.js';

// คนขับ sheet: ชื่อ | เลขบช
export function parseDrivers(rows) {
    if (!rows || rows.length < 2) return { drivers: [], stats: { total: 0, withBank: 0, skipped: 0 }, warnings: ['ไม่พบข้อมูล'] };
    const header = rows[0].map(clean);
    const iName = header.indexOf('ชื่อ');
    const iBank = header.findIndex(h => h.includes('เลขบ')); // เลขบช / เลขบัญชี
    if (iName < 0) return { drivers: [], stats: { total: 0, withBank: 0, skipped: 0 }, warnings: ['ไม่พบคอลัมน์ "ชื่อ" — อาจเลือกชีตผิด'] };
    const seen = new Set();
    const drivers = [];
    let skipped = 0, withBank = 0;
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        const name = clean(row[iName]);
        if (!name) { skipped++; continue; }
        const key = name.toLowerCase().replace(/\s+/g, ' ');
        if (seen.has(key)) continue;
        seen.add(key);
        const bankAccount = iBank >= 0 ? clean(row[iBank]) : '';
        if (bankAccount) withBank++;
        drivers.push({ name, bankAccount });
    }
    return { drivers, stats: { total: drivers.length, withBank, skipped }, warnings: [] };
}

// ทะเบียน sheet: ทะเบียน | ประเภทรถ | ประเภทรถขับ (รถบริษัท/รถร่วม) | บริษัทรถ | Remark
export function parseVehicles(rows) {
    if (!rows || rows.length < 2) return { vehicles: [], stats: { total: 0, own: 0, partner: 0, skipped: 0 }, warnings: ['ไม่พบข้อมูล'] };
    const header = rows[0].map(clean);
    const iPlate = header.indexOf('ทะเบียน');
    const iType = header.findIndex(h => h.startsWith('ประเภทร') && !h.includes('ขับ')); // 'ประเภทรถ' / 'ประเภทรภ'(typo)
    const iDrive = header.findIndex(h => h.includes('ขับ'));                            // ประเภทรถขับ
    const iCompany = header.indexOf('บริษัทรถ');
    if (iPlate < 0) return { vehicles: [], stats: { total: 0, own: 0, partner: 0, skipped: 0 }, warnings: ['ไม่พบคอลัมน์ "ทะเบียน" — อาจเลือกชีตผิด'] };
    const seen = new Set();
    const vehicles = [];
    let skipped = 0, own = 0, partner = 0;
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        const registrationNumber = clean(row[iPlate]);
        if (!registrationNumber || registrationNumber === 'ทะเบียน') { skipped++; continue; } // skip blanks & repeated headers
        const key = registrationNumber.toLowerCase().replace(/\s+/g, ' ');
        if (seen.has(key)) continue;
        seen.add(key);
        const driveType = iDrive >= 0 ? clean(row[iDrive]) : '';
        if (driveType.includes('ร่วม')) partner++; else own++;
        vehicles.push({
            registrationNumber,
            vehicleType: iType >= 0 ? clean(row[iType]) : '',
            driveType,
            company: iCompany >= 0 ? clean(row[iCompany]) : '',
        });
    }
    return { vehicles, stats: { total: vehicles.length, own, partner, skipped }, warnings: [] };
}

// ฐานคำนวน sheet: ประเภทวิ่ง | ประเภทรถ | เบี้ยเลี้ยง | หักเงินสะสม
export function parseRateTable(rows) {
    if (!rows || rows.length < 2) return { rates: [], stats: { total: 0, duplicates: 0, skipped: 0 }, warnings: ['ไม่พบข้อมูล'] };
    const header = rows[0].map(clean);
    const iRun = header.indexOf('ประเภทวิ่ง');
    const iType = header.indexOf('ประเภทรถ');
    const iAllow = header.indexOf('เบี้ยเลี้ยง');
    const iDeduct = header.findIndex(h => h.includes('หักเงิน'));
    if (iRun < 0 || iType < 0) return { rates: [], stats: { total: 0, duplicates: 0, skipped: 0 }, warnings: ['ไม่พบคอลัมน์ ประเภทวิ่ง/ประเภทรถ — อาจเลือกชีตผิด'] };
    const seen = new Set();
    const rates = [];
    let skipped = 0, duplicates = 0;
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        const runType = clean(row[iRun]);
        const vehicleType = clean(row[iType]);
        if (!runType || !vehicleType) { skipped++; continue; }
        const key = `${runType}|${vehicleType}`.toLowerCase().replace(/\s+/g, ' ');
        if (seen.has(key)) { duplicates++; continue; }
        seen.add(key);
        rates.push({
            runType, vehicleType,
            allowance: iAllow >= 0 ? parseNum(row[iAllow]) : 0,
            deduction: iDeduct >= 0 ? parseNum(row[iDeduct]) : 0,
        });
    }
    return { rates, stats: { total: rates.length, duplicates, skipped }, warnings: [] };
}
