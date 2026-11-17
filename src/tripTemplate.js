// Excel template for importing trips (ขาขึ้น / ขาล่อง) in place of the "บันทึกน้ำมัน" form.
// User decisions 17 Nov 2026: one sheet, 1 row = 1 trip leg with a ขาขึ้น/ขาล่อง column; money columns optional;
// the dispatcher fills it daily; from the start month on, trips come from this template only.
// Pure functions (no DOM/Firebase): ExcelJS is passed in, so Node tests and the browser share this code.

import { clean } from './parseUtils.js';
import { GAS_STATION_FRIENDLY_NAMES, ROUTE_STATION_COLUMNS } from './constants.js';
import { textDropPlace, normPlate, plateKey } from './fuelAudit.js';
import { createAllIfAbsent } from './dbUtils.js';
import { ymd } from './importerTrips.js';

export const TEMPLATE_SHEET = 'นำเข้าเที่ยวรถ';
export const LIST_SHEET = 'รายการ';
export const HELP_SHEET = 'วิธีกรอก';
export const DIRECTIONS = ['ขาขึ้น', 'ขาล่อง'];
export const NO_DROP = 'ไม่ตัด';
export const DROP_OPTIONS = [NO_DROP, 'นครสวรรค์', 'กำแพงเพชร'];
export const DEFAULT_QUEUE_TYPES = ['ลนต.', 'ลงงานห้าง', 'ตีเปล่า', 'ค้าง', 'หนักล่อง'];
const TEMPLATE_ROWS = 3000;          // rows that get dropdowns / number checks (~2 months of trips; help says 1 file per month)
const MAX_LITERS_PER_PUMP = 1000;    // above this a fill is almost certainly a typo → warning

const pumpName = (k) => GAS_STATION_FRIENDLY_NAMES[k] || k;

// Column layout. `list` = dropdown from the lists sheet; `strict` = Excel refuses other values (stop style).
export const TEMPLATE_COLUMNS = [
    { key: 'date', header: 'วันที่', width: 12, type: 'date', required: true },
    // Header renamed from "ทิศทาง" (user, 17 Nov 2026); old downloaded templates still import via `aliases`.
    { key: 'direction', header: 'ขาขึ้น/ขาล่อง', aliases: ['ทิศทาง'], width: 13, list: 'directions', strict: true, required: true },
    { key: 'queueType', header: 'ประเภทคิว', width: 12, list: 'queueTypes' },
    { key: 'driver', header: 'คนขับ', width: 22, list: 'drivers' },
    { key: 'head', header: 'ทะเบียนหัว', width: 12, list: 'heads', required: true },
    { key: 'trailer', header: 'ทะเบียนหาง', width: 11, list: 'trailers' },
    { key: 'origin', header: 'ต้นทาง', width: 16, list: 'origins' },
    { key: 'destination', header: 'ปลายทาง', width: 20, list: 'destinations' },
    { key: 'trailerDrop', header: 'ตัดหางที่', width: 11, list: 'drops', strict: true },
    { key: 'jobDetail', header: 'รายการงาน', width: 28 },
    ...ROUTE_STATION_COLUMNS.flatMap(s => s.units.map(u => ({
        key: `${u === 'Baht' ? 'B' : 'L'}:${s.key}`, station: s.key, unit: u,
        header: `${u === 'Baht' ? 'บาท' : 'ลิตร'} ${pumpName(s.key)}`, width: 13, type: 'num', group: 'fuel',
    }))),
    { key: 'revenue', header: 'ค่าเที่ยว (บาท)', width: 12, type: 'num', group: 'money' },
    { key: 'netTransfer', header: 'โอนสุทธิ (บาท)', width: 12, type: 'num', group: 'money' },
    { key: 'fuelCost', header: 'ค่าน้ำมัน (บาท)', width: 12, type: 'num', group: 'money' },
    { key: 'driverAllowance', header: 'เบี้ยเลี้ยง (บาท)', width: 12, type: 'num', group: 'money' },
    { key: 'deduction', header: 'หัก (บาท)', width: 10, type: 'num', group: 'money' },
    { key: 'distance', header: 'ระยะทาง (กม.)', width: 11, type: 'num', group: 'money' },
    { key: 'notes', header: 'หมายเหตุ', width: 28 },
];

// First two dash parts = the head ("70-8901-8158" → "70-8901"); same rule as the fuel audit's vehicle lookup.
export const headOf = plateKey;
const uniqSorted = (arr) => [...new Set(arr.map(clean).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'th'));

// Dropdown lists, rebuilt from the masters every time the template is downloaded.
export function buildTemplateLists({ drivers = [], vehicles = [], rateStandards = [], routes = [], runTypes = [] } = {}) {
    const plates = vehicles.map(v => normPlate(v?.registrationNumber)).filter(Boolean);
    return {
        directions: DIRECTIONS,
        drops: DROP_OPTIONS,
        queueTypes: uniqSorted([...DEFAULT_QUEUE_TYPES, ...runTypes]),
        drivers: uniqSorted(drivers.map(d => d?.name)),
        heads: uniqSorted(plates.map(headOf)),
        trailers: uniqSorted(plates.filter(p => p.split('-').length > 2).map(p => p.split('-').slice(2).join('-'))),
        origins: uniqSorted(routes.map(r => r?.origin)),
        destinations: uniqSorted([...rateStandards.map(s => s?.destination), ...routes.map(r => r?.destination)]),
    };
}

const colLetter = (n) => { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

const HELP_LINES = [
    'วิธีกรอกแม่แบบนำเข้าเที่ยวรถ',
    '',
    '1. กรอกในชีต "นำเข้าเที่ยวรถ" — 1 แถว = 1 เที่ยว (ขาขึ้นหรือขาล่อง แยกแถวกัน)',
    '2. ช่องที่ต้องกรอกทุกแถว: วันที่, ขาขึ้น/ขาล่อง, ทะเบียนหัว',
    '   ช่อง "ขาขึ้น/ขาล่อง": เลือก ขาขึ้น (ออกไปส่งของต่างจังหวัด) หรือ ขาล่อง (เที่ยวกลับ) — ไปกลับ 1 รอบ = 2 แถว',
    '3. วันที่: พิมพ์แบบ วัน/เดือน/ปี เช่น 17/11/2569 หรือ 17/11/2026 · ตอนนำเข้าให้ดูช่องวันที่ในหน้าตัวอย่างว่าถูกต้อง',
    '   แนะนำ: ใช้ไฟล์ใหม่ทุกเดือน (ดาวน์โหลดแม่แบบใหม่ต้นเดือน จะได้รายชื่อคนขับ/ทะเบียนล่าสุด)',
    '4. ทะเบียนหัว / ทะเบียนหาง: เลือกจากรายการ — รถวิ่งเดี่ยวให้เว้นทะเบียนหางว่าง',
    '   รถพ่วง เช่น หัว 70-8901 หาง 8158 → ระบบรวมเป็น 70-8901-8158 ให้เอง',
    '5. ตัดหางที่: ว่าง = ไม่ตัด · "นครสวรรค์" = ตัดหางที่ นว. (ขาขึ้นคิดเรทตัวแม่) · "กำแพงเพชร" = ถอดหางที่ กพ.',
    '6. น้ำมัน: ใส่ "ลิตร" แยกตามปั๊มเป็นตัวเลขเท่านั้น (ไม่ใส่คำว่า ลิตร) · ปั๊มที่จ่ายเป็นเงินใช้ช่อง "บาท"',
    '   ขาล่องที่เติมไพศาล ให้กรอกในแถวขาล่อง ช่อง "ลิตร ดาวไพศาล (นครสวรรค์)"',
    '7. ช่องเงิน (ค่าเที่ยว / โอนสุทธิ / ค่าน้ำมัน / เบี้ยเลี้ยง / หัก / ระยะทาง) ไม่บังคับ',
    '8. ปลายทางที่ไม่มีในรายการพิมพ์เองได้ (Excel จะเตือนแต่บันทึกได้)',
    '',
    'ตอนนำเข้า (หน้า "นำเข้าข้อมูล" ในเว็บ):',
    '• ระบบแสดงตัวอย่างและผลตรวจน้ำมันก่อนบันทึกเสมอ',
    '• แถวที่ผิด (ไม่มีวันที่ / ขาขึ้น/ขาล่อง / ทะเบียนหัว, ตัวเลขผิด) จะไม่ถูกบันทึก — ดาวน์โหลดไฟล์ "แถวที่ผิด" ไปแก้แล้วนำเข้าใหม่',
    '• นำเข้าไฟล์เดิมซ้ำได้ ระบบไม่สร้างเที่ยวซ้ำ และไม่ทับข้อมูลที่บันทึกไปแล้ว (แก้ข้อมูลเก่าที่หน้า "คิวรถ")',
    '',
    'ตัวอย่างการกรอก (อย่ากรอกตัวอย่างนี้ลงชีตข้อมูล):',
    '  วันที่ 17/11/2569 · ขาขึ้น/ขาล่อง = ขาขึ้น · ทะเบียนหัว 70-8901 · ทะเบียนหาง 8158 · ปลายทาง กำแพงเพชร · ตัดหางที่ นครสวรรค์ · ลิตร จิตตมาศ 50',
];

// Build the workbook. Returns the ExcelJS workbook; the caller writes it (writeBuffer / writeFile).
export function generateTemplate(ExcelJS, lists) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'ระบบติดตามน้ำมันโลจิสติกส์';
    const ws = wb.addWorksheet(TEMPLATE_SHEET, { views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }] });
    const help = wb.addWorksheet(HELP_SHEET);
    const lst = wb.addWorksheet(LIST_SHEET, { state: 'hidden' });

    // Lists sheet: one column per list, header in row 1.
    const listRange = {};
    Object.entries(lists).forEach(([name, values], i) => {
        const c = colLetter(i + 1);
        lst.getCell(`${c}1`).value = name;
        values.forEach((v, j) => { lst.getCell(`${c}${j + 2}`).value = v; });
        listRange[name] = `'${LIST_SHEET}'!$${c}$2:$${c}$${Math.max(2, values.length + 1)}`;
    });

    ws.columns = TEMPLATE_COLUMNS.map(c => ({ header: c.header, key: c.key, width: c.width }));
    const head = ws.getRow(1);
    head.height = 36;
    TEMPLATE_COLUMNS.forEach((c, i) => {
        const cell = head.getCell(i + 1);
        const fill = c.required ? 'FFFDE68A' : c.group === 'fuel' ? 'FFBFDBFE' : c.group === 'money' ? 'FFE5E7EB' : 'FFD1FAE5';
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
        cell.font = { bold: true };
        cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
        const col = colLetter(i + 1);
        const range = `${col}2:${col}${TEMPLATE_ROWS + 1}`;
        if (c.type === 'date') {
            ws.getColumn(i + 1).numFmt = 'dd/mm/yyyy';
        } else if (c.type === 'num') {
            ws.getColumn(i + 1).numFmt = '#,##0.##';
            ws.dataValidations.add(range, { type: 'decimal', operator: 'greaterThanOrEqual', allowBlank: true, formulae: [0], showErrorMessage: true, errorStyle: 'stop', errorTitle: 'ตัวเลขเท่านั้น', error: 'ช่องนี้ใส่ได้เฉพาะตัวเลข 0 ขึ้นไป' });
        } else if (c.list && listRange[c.list]) {
            ws.dataValidations.add(range, {
                type: 'list', allowBlank: true, formulae: [listRange[c.list]], showErrorMessage: true,
                errorStyle: c.strict ? 'stop' : 'warning',
                errorTitle: c.strict ? 'เลือกจากรายการเท่านั้น' : 'ไม่อยู่ในรายการ',
                error: c.strict ? `เลือก: ${(lists[c.list] || []).join(' / ')}` : 'ค่านี้ไม่อยู่ในรายการ — กด "ใช่" ถ้าต้องการใช้ค่านี้',
            });
        }
    });

    help.getColumn(1).width = 120;
    HELP_LINES.forEach((t, i) => { const cell = help.getCell(`A${i + 1}`); cell.value = t; if (i === 0) cell.font = { bold: true, size: 14 }; });
    return wb;
}

// Row index of the template header (searching the first rows), or -1.
export function findTemplateHeader(rows) {
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
        const cells = (rows[i] || []).map(clean);
        if ((cells.includes('ขาขึ้น/ขาล่อง') || cells.includes('ทิศทาง')) && cells.includes('ทะเบียนหัว') && cells.includes('ตัดหางที่')) return i;
    }
    return -1;
}

// Excel date cell → 'YYYY-MM-DD' ('' when unreadable). Handles Date objects (SheetJS cellDates, which can land a
// few seconds before local midnight), Excel serial numbers, and typed text: d/m/yyyy, d/m/yy (Thai 2-digit year
// such as 69 = 2569), YYYY-MM-DD; Buddhist-era years are converted.
export function parseTemplateDate(v) {
    let y, m, d;
    if (v instanceof Date && !isNaN(v)) {
        // Same reading as every Excel import (absorbs SheetJS's few seconds before midnight).
        [y, m, d] = ymd(v).split('-').map(Number);
    } else if (typeof v === 'number' && isFinite(v) && v > 20000 && v < 80000) {
        const t = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5);
        y = t.getUTCFullYear(); m = t.getUTCMonth() + 1; d = t.getUTCDate();
    } else {
        const s = clean(v);
        let mm = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (mm) { y = +mm[1]; m = +mm[2]; d = +mm[3]; }
        else if ((mm = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2}|\d{4})$/))) {
            d = +mm[1]; m = +mm[2]; y = +mm[3];
            if (mm[3].length === 2) y = y >= 40 ? 2500 + y : 2000 + y; // "69" → 2569 (Thai habit), "26" → 2026
        } else return '';
    }
    if (y >= 2400) y -= 543;
    if (y < 2020 || y > 2100 || m < 1 || m > 12 || d < 1) return '';
    const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
    if (d > last) return '';
    return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

// Strict number: blank → 0; anything that is not a plain non-negative number → NaN (blocks the row).
const parseStrictNum = (v) => {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) && v >= 0 ? v : NaN;
    const s = clean(v).replace(/[,\s]/g, '');
    if (s === '' || s === '-') return 0;
    return /^\d+(\.\d+)?$/.test(s) ? parseFloat(s) : NaN;
};

const thMonth = (ym) => { const [y, m] = ym.split('-').map(Number); return `${m}/${y + 543}`; };

/**
 * Parse template rows into trips with per-row errors (block saving) and warnings (saved, shown).
 * rows: sheet_to_json(ws, { header: 1, blankrows: true, defval: '' }); firstExcelRow: Excel row number of rows[0].
 * opts: { startMonth:'YYYY-MM', vehicles, drivers, fileName, today:'YYYY-MM-DD' }
 */
export function parseTemplateRows(rows, opts = {}, firstExcelRow = 1) {
    const { startMonth = '', vehicles = [], drivers = [], fileName = '', today = '' } = opts;
    const h = findTemplateHeader(rows);
    if (h < 0) return { rows: [], error: 'ไม่พบหัวตารางของแม่แบบ (ต้องมีช่อง ขาขึ้น/ขาล่อง / ทะเบียนหัว / ตัดหางที่)' };
    const header = rows[h].map(clean);
    const idx = {};
    const missing = [];
    for (const c of TEMPLATE_COLUMNS) {
        const names = [c.header, ...(c.aliases || [])];
        const found = names.filter(name => header.includes(name));
        // Old and new name both present (e.g. a column added by hand) → do not guess which one holds the data.
        if (found.length > 1) return { rows: [], error: `แม่แบบมีทั้งช่อง ${found.map(n => `"${n}"`).join(' และ ')} — ลบช่องที่ไม่ใช้ออกให้เหลือช่องเดียว แล้วนำเข้าใหม่` };
        const i = found.length ? header.indexOf(found[0]) : -1;
        if (i >= 0) idx[c.key] = i; else if (c.required || c.type === 'num' && c.group === 'fuel') missing.push(c.header);
    }
    if (missing.length) return { rows: [], error: `แม่แบบไม่ครบ ขาดช่อง: ${missing.join(', ')} — ดาวน์โหลดแม่แบบใหม่` };

    const plateSet = new Set(vehicles.map(v => normPlate(v?.registrationNumber)).filter(Boolean));
    const driverSet = new Set(drivers.map(d => clean(d?.name)).filter(Boolean));
    const get = (r, key) => (idx[key] == null ? '' : r[idx[key]]);
    const out = [];
    for (let i = h + 1; i < rows.length; i++) {
        const r = rows[i] || [];
        if (!r.some(v => clean(v) !== '')) continue;
        const excelRow = firstExcelRow + i;
        const errors = [], warnings = [];

        const date = parseTemplateDate(get(r, 'date'));
        if (!date) errors.push(clean(get(r, 'date')) ? `วันที่อ่านไม่ได้ ("${clean(get(r, 'date'))}")` : 'ไม่มีวันที่');
        else {
            if (startMonth && date.slice(0, 7) < startMonth) errors.push(`วันที่ก่อนเดือนเริ่มใช้แม่แบบ (${thMonth(startMonth)}) — เที่ยวช่วงนี้มาจากกระดานคิว/ไฟล์ Transport แล้ว`);
            if (today && date > today) warnings.push('วันที่เป็นวันในอนาคต');
        }
        const direction = clean(get(r, 'direction'));
        if (!DIRECTIONS.includes(direction)) errors.push(direction ? `ช่อง ขาขึ้น/ขาล่อง ใส่ "${direction}" — ต้องเลือก ขาขึ้น หรือ ขาล่อง` : 'ไม่ได้เลือก ขาขึ้น/ขาล่อง');

        const head = normPlate(get(r, 'head'));
        let trailer = normPlate(get(r, 'trailer'));
        if (head && trailer.startsWith(`${head}-`)) trailer = trailer.slice(head.length + 1); // full plate typed in the trailer cell
        const plate = head ? (trailer ? `${head}-${trailer}` : head) : '';
        if (!head) errors.push('ไม่มีทะเบียนหัว');
        else if (plateSet.size && !plateSet.has(plate)) warnings.push(`ไม่พบทะเบียน ${plate} ในหน้าจัดการทะเบียนรถ — ตรวจน้ำมันไม่ได้ถ้าไม่รู้ประเภทรถ`);

        const driver = clean(get(r, 'driver'));
        if (!driver) warnings.push('ไม่มีชื่อคนขับ — ยกลิตรข้ามเที่ยวให้คนขับไม่ได้');
        else if (driverSet.size && !driverSet.has(driver)) warnings.push(`ไม่พบคนขับ "${driver}" ในรายชื่อ`);

        const destination = clean(get(r, 'destination'));
        if (!destination && direction === 'ขาขึ้น') warnings.push('ไม่มีปลายทาง — ตรวจอัตราน้ำมันไม่ได้');

        let trailerDrop = clean(get(r, 'trailerDrop'));
        if (trailerDrop && !DROP_OPTIONS.includes(trailerDrop)) errors.push(`ตัดหางที่ "${trailerDrop}" ต้องเป็น ${DROP_OPTIONS.join(' / ')} หรือเว้นว่าง`);
        if (!trailerDrop) trailerDrop = NO_DROP;
        if (trailerDrop !== NO_DROP) {
            if (direction === 'ขาล่อง') warnings.push('ช่องตัดหางใช้กับขาขึ้นเท่านั้น — ขาล่องไม่นำมาคิด');
            else if (plate.split('-').length < 3) warnings.push('ระบุตัดหาง แต่ไม่มีทะเบียนหาง');
        }

        const jobDetail = clean(get(r, 'jobDetail')), notes = clean(get(r, 'notes')), queueType = clean(get(r, 'queueType'));
        if (trailerDrop === NO_DROP && direction === 'ขาขึ้น') {
            const said = textDropPlace({ jobDetail, notes, queueType });
            if (said) warnings.push(`รายการงาน/หมายเหตุเขียนว่าตัดหางที่${said} แต่ช่อง "ตัดหางที่" ว่าง — ระบบถือว่าไม่ตัด`);
        }

        const gasStationsLiters = {}, gasStationsBaht = {}, money = {};
        for (const c of TEMPLATE_COLUMNS) {
            if (c.type !== 'num' || idx[c.key] == null) continue;
            const raw = get(r, c.key);
            const n = parseStrictNum(raw);
            if (Number.isNaN(n)) { errors.push(`ช่อง "${c.header}" ไม่ใช่ตัวเลข ("${clean(raw)}")`); continue; }
            if (c.group === 'fuel') {
                if (!n) continue;
                const bucket = c.unit === 'Baht' ? gasStationsBaht : gasStationsLiters;
                bucket[c.station] = (bucket[c.station] || 0) + n;
                if (c.unit !== 'Baht' && n > MAX_LITERS_PER_PUMP) warnings.push(`${c.header} ${n.toLocaleString()} ลิตร มากผิดปกติ`);
            } else money[c.key] = n;
        }
        const totalLiters = Object.values(gasStationsLiters).reduce((a, b) => a + b, 0);
        const totalBaht = Object.values(gasStationsBaht).reduce((a, b) => a + b, 0);

        const trip = {
            date, yearMonth: date ? date.slice(0, 7) : '', direction, driver, plate, trailerPlate: trailer,
            vehicleType: '', origin: clean(get(r, 'origin')), destination, queueType, jobDetail, notes, trailerDrop,
            gasStationsLiters, gasStationsBaht, totalLiters, totalBaht,
            revenue: money.revenue || 0, netTransfer: money.netTransfer || 0, fuelCost: money.fuelCost || 0,
            driverAllowance: money.driverAllowance || 0, deduction: money.deduction || 0, distance: money.distance || 0,
            crossingFee: '', waitingFee: '', extraNote: '',
            source: 'template', importedFrom: clean(fileName), sourceRow: excelRow,
        };
        out.push({ excelRow, status: errors.length ? 'error' : warnings.length ? 'warn' : 'ok', errors, warnings, trip, raw: r });
    }
    assignTemplateKeys(out);
    return { rows: out, header: rows[h], error: '' };
}

// Small sync string hash (two FNV-1a passes) — keys must be computable in the browser without async crypto.
const hashStr = (s) => {
    let a = 0x811c9dc5, b = 0x01000193 ^ 0x5bd1e995;
    for (let i = 0; i < s.length; i++) {
        const c = s.charCodeAt(i);
        a = Math.imul(a ^ c, 0x01000193) >>> 0;
        b = Math.imul(b ^ c, 0x5bd1e995) >>> 0;
    }
    return a.toString(36).padStart(7, '0') + b.toString(36).padStart(7, '0');
};

// Stable key per trip: same row imported again → same key → never saved twice. The identity leaves out fuel and
// money, so a corrected figure does not create a second trip (it is reported as "already imported" instead).
// occurrence numbers identical identities within one file (two runs, same truck/day/destination). Rows with
// errors still count when their identity is readable, so fixing one later does not take a saved row's number.
export function assignTemplateKeys(parsedRows) {
    const seen = {};
    for (const p of parsedRows) {
        const t = p.trip;
        const readable = t.date && DIRECTIONS.includes(t.direction) && t.plate;
        const id = `${t.date}|${t.direction}|${headOf(t.plate)}|${t.destination}`;
        if (readable) seen[id] = (seen[id] || 0) + 1;
        p.key = p.status === 'error' || !readable ? '' : `tpl_${t.date}_${hashStr(`${id}|${seen[id]}`)}`;
    }
    return parsedRows;
}

// Order-insensitive compare (Firebase returns object keys sorted, the parser keeps column order).
const canonObj = (o) => Object.keys(o || {}).sort().map(k => [k, Number(o[k]) || 0]);
const sameTripContent = (a, b) => {
    const pick = (t) => JSON.stringify([clean(t.driver), clean(t.plate), clean(t.origin), clean(t.destination), clean(t.queueType), clean(t.trailerDrop),
        canonObj(t.gasStationsLiters), canonObj(t.gasStationsBaht), Number(t.revenue) || 0, Number(t.netTransfer) || 0, Number(t.fuelCost) || 0,
        Number(t.driverAllowance) || 0, Number(t.deduction) || 0, Number(t.distance) || 0, clean(t.jobDetail), clean(t.notes)]);
    return pick(a) === pick(b);
};

const tripSource = (key, t) => t?.source || (key.startsWith('tdh_') ? 'transport' : key.startsWith('daily_') ? 'daily' : '');

/**
 * Decide what happens to each valid row, against the trips already in the database and earlier rows of the file.
 * existingByMonth: { 'YYYY-MM': { key: trip } }.
 * importState: 'new' | 'already' (same key exists; `changed` when the figures differ — never overwritten)
 *            | 'possible-dup' (the same truck, day and direction is already in the database or earlier in this
 *              file — skipped unless the user ticks that row). dupOf: { source, destination, row? }.
 */
export function planTemplateSave(parsedRows, existingByMonth) {
    // One index per month: date|direction|head → first matching [key, trip].
    const index = {};
    const indexFor = (ym) => {
        if (index[ym]) return index[ym];
        const m = new Map();
        for (const [k, e] of Object.entries(existingByMonth[ym] || {})) {
            if (!e) continue;
            const id = `${e.date}|${clean(e.direction)}|${headOf(e.plate)}`;
            if (!m.has(id)) m.set(id, [k, e]);
        }
        return (index[ym] = m);
    };
    const inFile = new Map(); // date|direction|head → first valid row of this file
    for (const p of parsedRows) {
        p.importState = ''; p.changed = false; p.dupOf = null;
        if (p.status === 'error' || !p.key) continue;
        const t = p.trip;
        const month = existingByMonth[t.yearMonth] || {};
        const id = `${t.date}|${t.direction}|${headOf(t.plate)}`;
        if (month[p.key]) { p.importState = 'already'; p.changed = !sameTripContent(month[p.key], t); }
        else {
            const hit = indexFor(t.yearMonth).get(id);
            const earlier = inFile.get(id);
            if (hit && hit[0] !== p.key) { p.importState = 'possible-dup'; p.dupOf = { source: tripSource(hit[0], hit[1]), destination: clean(hit[1].destination) }; }
            else if (earlier) { p.importState = 'possible-dup'; p.dupOf = { source: 'file', row: earlier.excelRow, destination: earlier.trip.destination }; }
            else p.importState = 'new';
        }
        if (!inFile.has(id)) inFile.set(id, p);
    }
    return parsedRows;
}

/**
 * Save the chosen rows with their deterministic keys: every 'new' row plus the 'possible-dup' rows whose keys the
 * user ticked. Each trip is written with a create-only write (a Firebase transaction that aborts when the key
 * already holds data), so two people importing the same file at the same moment can never overwrite each other.
 * db: { createIfAbsent(ym, key, trip) → Promise<boolean> } — true when written, false when the key already existed.
 * Returns { added, skipped, failed, savedKeys }; a failed write leaves the others in place, and importing the same
 * file again only fills in what is missing.
 */
export async function saveTemplateTrips(db, parsedRows, { includeKeys = [], concurrency = 10 } = {}) {
    const extra = new Set(includeKeys);
    const seen = new Set();
    const chosen = parsedRows.filter(p => p.status !== 'error' && p.key
        && (p.importState === 'new' || (p.importState === 'possible-dup' && extra.has(p.key)))
        && !seen.has(p.key) && seen.add(p.key));
    const { added, skipped, failed, error, addedItems } = await createAllIfAbsent(chosen, p => db.createIfAbsent(p.trip.yearMonth, p.key, p.trip), concurrency);
    return { added, skipped, failed, error, savedKeys: addedItems.map(p => p.key) };
}
