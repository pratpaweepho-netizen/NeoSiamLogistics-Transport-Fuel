// Phase 2 — parse the daily trip queues (คิวขาขึ้น / คิวขาล่อง) into a unified "trip" model.
// Pure functions (no DOM/Firebase) so they can be unit-tested in Node.
// A trip is stored under trips/{yearMonth}/{id} to avoid loading all ~16k rows at once.

import { parseNum, clean } from './parseUtils.js';

// Local calendar date (Thai tz in the browser) → 'YYYY-MM-DD'. Avoids UTC off-by-one.
const ymd = (v) => {
    if (v instanceof Date && !isNaN(v)) {
        let y = v.getFullYear();
        if (y >= 2500) y -= 543; // normalize Buddhist-era years (e.g. 2569 → 2026)
        return `${y}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
    }
    const s = clean(v);
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return '';
    let y = parseInt(m[1], 10);
    if (y >= 2500) y -= 543;
    return `${y}-${m[2]}-${m[3]}`;
};

// Station columns by fixed index within each sheet (layout is stable in Transport 2026.xlsx).
const UP_STATIONS = [
    { c: 11, key: 'jittamas_bangpain', unit: 'L' },
    { c: 12, key: 'jittamas_bangpain', unit: 'B' },
    { c: 13, key: 'srinakorn_klongklung', unit: 'L' },
    { c: 14, key: 'daopaisan_nakhonsawan', unit: 'L' },
    { c: 15, key: 'nissan_singburi', unit: 'L' },
    { c: 16, key: 'daourai_sai3', unit: 'B' },
    { c: 17, key: 'caltex_sahapat', unit: 'L' },
    { c: 18, key: 'caltex_sahapat', unit: 'B' },
    { c: 20, key: 'ekawitat_nonthaburi', unit: 'L' },
    { c: 21, key: 'daopaisan_nakhonsawan', unit: 'L' },
];
const DOWN_STATIONS = [
    { c: 8, key: 'daopaisan_nakhonsawan', unit: 'L' },
    { c: 9, key: 'nissan_singburi', unit: 'L' },
    { c: 10, key: 'ekawitat_nonthaburi', unit: 'L' },
];

const collectFuel = (row, stationCols) => {
    const gasStationsLiters = {}, gasStationsBaht = {};
    for (const s of stationCols) {
        const v = parseNum(row[s.c]);
        if (v > 0) {
            const bucket = s.unit === 'B' ? gasStationsBaht : gasStationsLiters;
            bucket[s.key] = (bucket[s.key] || 0) + v;
        }
    }
    return {
        gasStationsLiters, gasStationsBaht,
        totalLiters: Object.values(gasStationsLiters).reduce((a, b) => a + b, 0),
        totalBaht: Object.values(gasStationsBaht).reduce((a, b) => a + b, 0),
    };
};

// Detect which kind of sheet this is, so the Import UI can pick the right parser.
// Returns: 'fuelbase' | 'uphill' | 'downhill' | 'daily' | 'unknown'.
export function detectSheetType(rows) {
    if (!rows || !rows.length) return 'unknown';
    // Scan the first few rows — some sheets have their header on row 2/3.
    const scan = rows.slice(0, 4).map(r => (r || []).map(clean));
    const flat = scan.flat();
    const has = (s) => flat.includes(s);
    const h0 = scan[0] || [];

    if (has('trucktype') && has('จุดจอด')) return 'fuelbase';
    // Master sheets (Phase 3)
    if (clean(h0[0]) === 'ชื่อ' && flat.some(c => c.startsWith('เลขบ'))) return 'drivers';
    if (has('ประเภทวิ่ง') && has('เบี้ยเลี้ยง')) return 'ratetable';
    if (clean(h0[0]) === 'ทะเบียน' && has('บริษัทรถ')) return 'vehicles';
    // Transport 2026.xlsx trip sheets (header on row 0, station columns like ปั๊มจิตตมาศ).
    if (clean(h0[0]) === 'Date' && has('พขร.') && has('ปลายทาง')) return 'uphill';
    if (clean(h0[0]) === 'วันที่' && has('คนขับ') && has('ขึ้นที่')) return 'downhill';
    // Partner queue board (คิวรถร่วม-ขาขึ้น 2026.xlsx)
    if (has('คิวรถร่วม') && (has('ว/ด/ป') || has('รายละเอียด'))) return 'partnerqueue';
    // Daily dispatch board (คิวขาขึ้น ปี 2026.xlsx): title "คิวรถ วันที่...", generic single 'ปั๊มน้ำมัน' column.
    if (flat.some(c => c.startsWith('คิวรถ วันที่')) || (has('ขึ้นที่') && has('ลงที่') && has('ปั๊มน้ำมัน'))) return 'daily';
    if (has('พขร.') && has('ปั๊มจิตตมาศ')) return 'uphill';
    return 'unknown';
}

// คิวขาขึ้น (uphill). Returns { trips, stats, warnings }.
export function parseUphillQueue(rows) {
    const warnings = [];
    if (!rows || rows.length < 2) return { trips: [], stats: emptyStats(), warnings: ['ไม่พบข้อมูล'] };
    // Layout-drift guard: fuel columns are read by fixed index (UP_STATIONS). If a column was
    // inserted/removed upstream, the anchor header "ปั๊มจิตตมาศ" won't sit at its expected index
    // (11) — warn so the user can catch silently-shifted data instead of trusting bad numbers.
    const header0 = (rows[0] || []).map(clean);
    if (header0.some(h => h.includes('จิตตมาศ')) && !clean(header0[11]).includes('จิตตมาศ')) {
        warnings.push('⚠️ ตำแหน่งคอลัมน์ปั๊มอาจเปลี่ยนไปจากรูปแบบเดิม — ตรวจสอบยอดน้ำมันหลังนำเข้า');
    }
    const trips = [];
    let skipped = 0;
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.every(c => c === '' || c == null)) continue;
        const date = ymd(row[0]);
        const driver = clean(row[2]);
        const plate = clean(row[3]);
        if (!date || (!driver && !plate)) { skipped++; continue; }
        const fuel = collectFuel(row, UP_STATIONS);
        trips.push({
            date, yearMonth: date.slice(0, 7), direction: 'ขาขึ้น',
            queueType: clean(row[1]), driver, plate, vehicleType: clean(row[4]),
            crossingFee: clean(row[5]), waitingFee: clean(row[6]),
            origin: clean(row[7]), destination: clean(row[10]),
            jobDetail: clean(row[8]), notes: clean(row[9]),
            ...fuel,
            driverAllowance: parseNum(row[30]),
            fuelCost: parseNum(row[31]),
            revenue: parseNum(row[32]),
            deduction: parseNum(row[33]),
            netTransfer: parseNum(row[34]),
            distance: parseNum(row[37]),
            extraNote: clean(row[35]),
        });
    }
    return { trips, stats: statsOf(trips, skipped), warnings };
}

// คิวขาล่อง (downhill). Returns { trips, stats, warnings }.
export function parseDownhillQueue(rows) {
    const warnings = [];
    if (!rows || rows.length < 2) return { trips: [], stats: emptyStats(), warnings: ['ไม่พบข้อมูล'] };
    const trips = [];
    let skipped = 0;
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.every(c => c === '' || c == null)) continue;
        const date = ymd(row[0]);
        const driver = clean(row[1]);
        const plate = clean(row[3]);
        if (!date || (!driver && !plate)) { skipped++; continue; }
        const fuel = collectFuel(row, DOWN_STATIONS);
        trips.push({
            date, yearMonth: date.slice(0, 7), direction: 'ขาล่อง',
            queueType: clean(row[2]), driver, plate, vehicleType: clean(row[4]),
            crossingFee: '', waitingFee: '',
            origin: clean(row[5]), destination: clean(row[6]),
            jobDetail: clean(row[7]), notes: clean(row[20]),
            ...fuel,
            driverAllowance: parseNum(row[16]), // รวมค่าเดินทาง
            fuelCost: parseNum(row[14]),        // ค่าGasขาล่อง
            revenue: parseNum(row[13]),         // เงินที่ต้องได้
            deduction: parseNum(row[17]),       // หักเงินสะสม
            netTransfer: parseNum(row[19]),     // รวมเงิน
            distance: parseNum(row[21]),
            extraNote: '',
        });
    }
    return { trips, stats: statsOf(trips, skipped), warnings };
}

// Partner queue board (คิวรถร่วม-ขาขึ้น 2026.xlsx): เดือน | ว/ด/ป | คิวรถร่วม | พขร. | ทะเบียน | ต้นทาง | รายละเอียด | ปลายทาง
// No fuel/money — just dispatch. Tagged fleet:'รถร่วม'.
export function parsePartnerQueue(rows) {
    const warnings = [];
    if (!rows || rows.length < 2) return { trips: [], stats: emptyStats(), warnings: ['ไม่พบข้อมูล'] };
    const header = rows[0].map(clean);
    const iDate = header.indexOf('ว/ด/ป');
    const iDir = header.indexOf('คิวรถร่วม');
    const iDriver = header.indexOf('พขร.');
    const iPlate = header.indexOf('ทะเบียน');
    const iOrigin = header.indexOf('ต้นทาง');
    const iDetail = header.findIndex(h => h.includes('รายละเอียด'));
    const iDest = header.indexOf('ปลายทาง');
    if (iDate < 0 || iPlate < 0) return { trips: [], stats: emptyStats(), warnings: ['ไม่พบคอลัมน์ ว/ด/ป หรือ ทะเบียน — อาจเลือกชีตผิด'] };
    const trips = [];
    let skipped = 0;
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.every(c => c === '' || c == null)) continue;
        const date = ymd(row[iDate]);
        const plate = clean(row[iPlate]);
        if (!date || !plate) { skipped++; continue; }
        trips.push({
            date, yearMonth: date.slice(0, 7),
            direction: iDir >= 0 ? (clean(row[iDir]) || 'ขาขึ้น') : 'ขาขึ้น',
            fleet: 'รถร่วม',
            queueType: 'รถร่วม',
            driver: iDriver >= 0 ? clean(row[iDriver]) : '',
            plate,
            vehicleType: '',
            crossingFee: '', waitingFee: '',
            origin: iOrigin >= 0 ? clean(row[iOrigin]) : '',
            destination: iDest >= 0 ? clean(row[iDest]) : '',
            jobDetail: iDetail >= 0 ? clean(row[iDetail]) : '',
            notes: '',
            gasStationsLiters: {}, gasStationsBaht: {}, totalLiters: 0, totalBaht: 0,
            driverAllowance: 0, fuelCost: 0, revenue: 0, deduction: 0, netTransfer: 0, distance: 0, extraNote: '',
        });
    }
    return { trips, stats: statsOf(trips, skipped), warnings };
}

// Map a pump-name fragment (from the free-text ปั๊มน้ำมัน column) to a station key.
const STATION_NAME_MAP = [
    ['จิตตมาศ', 'jittamas_bangpain'],
    ['ศรีนคร', 'srinakorn_klongklung'],
    ['ไพศาล', 'daopaisan_nakhonsawan'],
    ['นิสสัน', 'nissan_singburi'],
    ['ดาวอุไร', 'daourai_sai3'],
    ['เอกวิ', 'ekawitat_nonthaburi'],
    ['เต็กย', 'tekyong_chonburi'],
    ['เครือสห', 'caltex_sahapat'],
    ['สหพัฒน', 'caltex_sahapat'],
    ['รุ่งพัฒน', 'rungpattana_wangnoi'],
    // Typos seen on the daily board, mapping confirmed by the user (9 Nov 2026).
    ['จิตมาศ', 'jittamas_bangpain'],
    ['บางปะอิน', 'jittamas_bangpain'],
    ['ไพศา', 'daopaisan_nakhonsawan'],
    ['อุไร', 'daourai_sai3'],
    ['เต็ก', 'tekyong_chonburi'],
    ['เต้กย', 'tekyong_chonburi'],
    ['ต็กย', 'tekyong_chonburi'],
];
const mapPumpName = (name) => { for (const [frag, key] of STATION_NAME_MAP) if (name.includes(frag)) return key; return null; };

// Parse the free-text ปั๊มน้ำมัน value (e.g. "190จิตตมาศ/50ศรีนคร/500บาทดาวอุไร") → per-pump liters/baht.
export function parsePumpText(text) {
    const liters = {}, baht = {};
    const s = clean(text);
    if (!s || s === 'XX') return { liters, baht };
    for (let tok of s.split('/')) {
        tok = tok.replace(/\([^)]*\)/g, '').trim(); // drop parenthetical notes
        if (!tok || tok === 'XX') continue;
        const m = tok.match(/^([\d,]+)/);           // leading number
        if (!m) continue;
        const num = parseInt(m[1].replace(/,/g, ''), 10);
        if (!isFinite(num) || num <= 0) continue;
        const isBaht = tok.includes('บาท');
        const name = tok.slice(m[1].length).replace('บาท', '').trim();
        const key = mapPumpName(name);
        if (!key) continue;
        const bucket = isBaht ? baht : liters;
        bucket[key] = (bucket[key] || 0) + num;
    }
    return { liters, baht };
}

// Parse ONE daily-board sheet (คิวขาขึ้น ปี 2026.xlsx). sheetName like "02.01.69" (dd.mm.yy BE) → date.
export function parseDailyRows(sheetName, rows) {
    const m = String(sheetName).match(/^(\d{2})\.(\d{2})\.(\d{2})$/);
    if (!m || !rows) return { trips: [], skipped: 0 };
    const date = `${1957 + parseInt(m[3], 10)}-${m[2]}-${m[1]}`; // BE yy (2500+yy) − 543 = 1957+yy
    const ym = date.slice(0, 7);
    const hi = rows.findIndex(r => { const s = r.map(c => clean(c)); return s.includes('ทะเบียน') && s.includes('ปั๊มน้ำมัน'); });
    if (hi < 0) return { trips: [], skipped: 0 };
    const H = rows[hi].map(c => clean(c));
    const iQ = H.indexOf('คิวขาขึ้น'), iDr = H.indexOf('พขร.'), iPl = H.indexOf('ทะเบียน'),
        iOr = H.indexOf('ขึ้นที่'), iJob = H.indexOf('รายการ'), iNote = H.indexOf('หมายเหตุ'),
        iDest = H.indexOf('ลงที่'), iFuel = H.indexOf('ปั๊มน้ำมัน');
    const trips = []; let skipped = 0;
    for (let i = hi + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r || r.every(c => c === '' || c == null)) continue;
        const plate = clean(r[iPl]), driver = iDr >= 0 ? clean(r[iDr]) : '';
        const { liters, baht } = parsePumpText(r[iFuel]);
        const hasFuel = Object.keys(liters).length > 0 || Object.keys(baht).length > 0;
        // Not a trip: the "ทะเบียน" header repeated for a second block, or a spare driver line with
        // no plate and no fuel. A plate-less row that carries fuel is kept so no fuel is lost.
        if (plate === 'ทะเบียน' || (!plate && !hasFuel)) { skipped++; continue; }
        trips.push({
            // Row position used in import keys (daily_<sheet>_r<sourceRow>). SheetJS sets no __rowNum__ with
            // header:1, so this is the 1-based position among the sheet's NON-blank rows, not the Excel row number.
            // It does not depend on which rows the parser skips, but it shifts if blank rows are added/removed
            // above a trip — the Jul–Nov 2026 keys already in Firebase use this numbering, so keep it.
            sourceRow: (r.__rowNum__ ?? i) + 1,
            date, yearMonth: ym, direction: 'ขาขึ้น', source: 'daily',
            queueType: iQ >= 0 ? clean(r[iQ]) : '', driver, plate, vehicleType: '',
            crossingFee: '', waitingFee: '',
            origin: iOr >= 0 ? clean(r[iOr]) : '', destination: iDest >= 0 ? clean(r[iDest]) : '',
            jobDetail: iJob >= 0 ? clean(r[iJob]) : '', notes: iNote >= 0 ? clean(r[iNote]) : '',
            gasStationsLiters: liters, gasStationsBaht: baht,
            totalLiters: Object.values(liters).reduce((a, b) => a + b, 0),
            totalBaht: Object.values(baht).reduce((a, b) => a + b, 0),
            driverAllowance: 0, fuelCost: 0, revenue: 0, deduction: 0, netTransfer: 0, distance: 0, extraNote: '',
        });
    }
    return { trips, skipped };
}

function emptyStats() { return { total: 0, skipped: 0, withFuel: 0, months: 0, minDate: '', maxDate: '' }; }
function statsOf(trips, skipped) {
    const dates = trips.map(t => t.date).filter(Boolean).sort();
    const months = new Set(trips.map(t => t.yearMonth));
    return {
        total: trips.length,
        skipped,
        withFuel: trips.filter(t => t.totalLiters > 0 || t.totalBaht > 0).length,
        months: months.size,
        minDate: dates[0] || '',
        maxDate: dates[dates.length - 1] || '',
    };
}
