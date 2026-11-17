// Fuel audit: check each truck trip's actual fills against the official rate sheet
// ("อัตราเติมน้ำมันประจำ ปั๊ม ต่างๆ"). Pure functions (no DOM/Firebase) so they can be tested in Node.

import { parseNum, clean } from './parseUtils.js';
import { GAS_STATION_FRIENDLY_NAMES } from './constants.js';

const pumpThai = (key) => GAS_STATION_FRIENDLY_NAMES[key] || key;
// Most common value of `key` among rows (ties: first seen) — used for "the usual downhill liters" of a class.
const modeOf = (rows, key) => {
    const c = new Map();
    rows.forEach(s => c.set(s[key], (c.get(s[key]) || 0) + 1));
    return [...c.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
};

export const AUDIT_CLASSES = {
    '10W': 'สิบล้อ',
    '18W': 'พ่วง',
    '6W': 'หกล้อ',
    '6W_BOX': 'หกล้อ isuzu + หางตู้',
    'KPP_DROP': 'ถอดหางที่กำแพงเพชร',
};

// Rate-sheet pump header fragment → station key used everywhere else in the app.
const PUMP_HEADERS = [
    ['จิตตมาศ', 'jittamas_bangpain'],
    ['ศรีนคร', 'srinakorn_klongklung'],
    ['ไพศาล', 'daopaisan_nakhonsawan'],
    ['นิสสัน', 'nissan_singburi'],
    ['ชำนาญ', 'chamnan'],
];
const PAISAN = 'daopaisan_nakhonsawan';
const TOTAL_UP_HEADER = 'รวมยอดขาขึ้น/ลิตร';

// Section title → truck class. Order matters: "หกล้อใหม่ isuzu + หางตู้" before plain หกล้อ.
const classOfTitle = (title) => {
    const t = clean(title).toLowerCase();
    if (t.includes('ยกเลิก')) return null;
    if (t.includes('isuzu') && t.includes('หางตู้')) return '6W_BOX';
    if (t.includes('หกล้อ')) return '6W';
    if (t.includes('สิบล้อ')) return '10W';
    if (t.includes('พ่วง')) return '18W';
    return null;
};

// "50+100" → 150, 80 → 80, "***" / '' → 0
const cellLiters = (v) => {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    const s = clean(v);
    if (/^\d+(\s*\+\s*\d+)*$/.test(s)) return s.split('+').reduce((a, b) => a + parseInt(b, 10), 0);
    return 0;
};

// Destination text → one canonical place name shared by the rate sheet and the trips.
// Order matters: districts before their province, แม่สาย/แม่สอด before shorter names.
const DEST_CANON = [
    ['วิเชียรบุรี', 'วิเชียรบุรี'], ['บึงสามพัน', 'บึงสามพัน'], ['เพชรบูรณ์', 'เพชรบูรณ์'],
    ['แม่สาย', 'แม่สาย'], ['แม่สอด', 'แม่สอด'], ['เชียงราย', 'เชียงราย'], ['เชียงใหม่', 'เชียงใหม่'],
    ['พะเยา', 'พะเยา'], ['ลำพูน', 'ลำพูน'], ['ลำปาง', 'ลำปาง'], ['พิษณุโลก', 'พิษณุโลก'],
    ['กำแพงเพชร', 'กำแพงเพชร'], ['นครสวรรค์', 'นครสวรรค์'], ['นครสรรค์', 'นครสวรรค์'], ['นคสวรรค์', 'นครสวรรค์'],
];
const DEST_EXACT = { 'ตาก': 'ตาก', 'แพร่': 'แพร่' };
export const canonDestination = (text) => {
    const s = clean(text).replace(/\s+/g, '');
    if (!s) return '';
    if (DEST_EXACT[s]) return DEST_EXACT[s];
    for (const [frag, canon] of DEST_CANON) if (s.includes(frag)) return canon;
    return '';
};

// The rate sheet only has two origins. Trip origins are customer/warehouse names around Bangkok, the
// Nakhon Sawan branch ("ย่อย นว", "สาขา นว") or another branch ("สาขา ชม") — the last has no rate.
const OTHER_BRANCH = /(^|[\s(+])(ชม|พล|พษ|กพ|ชร|ลป|ลพ|มส)\.?($|[\s)+])/;
const canonOrigin = (label) => {
    const s = clean(label);
    if (/นครสวรรค์|นครสรรค์|นคสวรรค์/.test(s) || /(^|[^ก-๙])นว\.?($|[^ก-๙])/.test(s)) return 'นครสวรรค์';
    if (OTHER_BRANCH.test(s)) return '';
    return 'กรุงเทพ';
};

// Parse every sheet of the rate workbook. sheets = { [sheetName]: rows[][] } (SheetJS header:1 rows).
// Returns { standards, warnings, skippedSheets }. Each standard:
// { id, cls, origin, destination, destinationLabel, variant:'standard'|'sino', up:{station:L}, upTotal,
//   down, downVictor, allowanceText, allowance, distance, sheet, row, flags[] }
// A sheet value may be rows[][] or { rows, startRow } where startRow is the 0-based Excel row of rows[0]
// (SheetJS starts at the sheet's used range, which is not always row 1).
export function parseRateSheet(sheets) {
    const standards = [], warnings = [], skippedSheets = [];
    for (const [sheet, value] of Object.entries(sheets || {})) {
        const rows = Array.isArray(value) ? value : (value?.rows || []);
        const startRow = Array.isArray(value) ? 0 : (value?.startRow || 0);
        if (!(rows || []).some(r => (r || []).some(c => clean(c) === TOTAL_UP_HEADER))) {
            const drop = parseTrailerDropSheet(sheet, rows, startRow);
            if (drop.length) standards.push(...drop); else skippedSheets.push(sheet);
            continue;
        }
        let cls = null, header = null;
        rows.forEach((r, i) => {
            const a = clean(r?.[0]);
            if (a === 'สถานที่') { cls = classOfTitle(r[2]); header = null; return; }
            if (a === 'ต้นทาง') { header = (r || []).map(clean); return; }
            if (!a.startsWith('จาก') || !header || !cls) return;
            const iTotal = header.indexOf(TOTAL_UP_HEADER);
            if (iTotal < 0) return;
            const up = {};
            for (let c = 2; c < iTotal; c++) {
                const key = PUMP_HEADERS.find(([frag]) => header[c].includes(frag))?.[1];
                const v = cellLiters(r[c]);
                if (key && v > 0) up[key] = (up[key] || 0) + v;
            }
            const downRaw = clean(r[iTotal + 1]);
            const down = parseNum((downRaw.match(/^\d+/) || [''])[0]) || (typeof r[iTotal + 1] === 'number' ? r[iTotal + 1] : 0);
            const victorM = downRaw.match(/victor\s*(\d+)/i);
            const allowanceText = clean(r[iTotal + 2]);
            const iDist = header.indexOf('ระยะทาง');
            const destinationLabel = clean(r[1]);
            const std = {
                cls, origin: canonOrigin(a), destination: canonDestination(destinationLabel), destinationLabel,
                variant: /sino/i.test(allowanceText) ? 'sino' : 'standard',
                up, upTotal: Object.values(up).reduce((x, y) => x + y, 0),
                down, downVictor: victorM ? parseInt(victorM[1], 10) : down,
                allowanceText, allowance: parseNum((allowanceText.match(/เบี้ยขับ\s*([\d,]+)/) || ['', ''])[1]),
                distance: iDist >= 0 ? clean(r[iDist]) : '',
                sheet, row: startRow + i + 1, flags: [],
            };
            const sheetTotal = cellLiters(r[iTotal]);
            if (sheetTotal !== std.upTotal) std.flags.push(`ผลรวมรายปั๊ม ${std.upTotal} ไม่เท่าช่องรวม ${sheetTotal}`);
                    if (!std.destination) std.flags.push('อ่านปลายทางไม่ออก');
            if (std.origin === 'นครสวรรค์' && std.up.nissan_singburi) std.flags.push('ต้นทางนครสวรรค์แต่เติมปั๊มนิสสัน (สิงห์บุรี) — ตรวจคอลัมน์ในไฟล์');
            standards.push(std);
        });
    }
    // Same class + origin + destination + variant twice with different liters → the sheet contradicts itself.
    const byKey = {};
    for (const s of standards) (byKey[standardKey(s)] ||= []).push(s);
    for (const list of Object.values(byKey)) {
        if (list.length < 2) continue;
        const sig = (s) => JSON.stringify([s.up, s.down, s.downVictor, s.upTotal, (s.parts || []).map(p => [p.part, p.destination, p.liters, p.allowance])]);
        const same = list.every(s => sig(s) === sig(list[0]));
        list.forEach((s, n) => {
            s.flags.push(same ? 'แถวซ้ำ (ค่าเท่ากัน)' : `ขัดกับแถว ${list.filter(o => o !== s).map(o => `${o.sheet} แถว ${o.row}`).join(', ')}`);
            if (n > 0) s.duplicateOf = `${list[0].sheet}!${list[0].row}`; // audit uses the first row
        });
    }
    // Downhill liters that differ from the rest of the same class (e.g. 179 vs 180) — likely a typo.
    for (const s of standards) {
        if (!s.down || s.duplicateOf) continue;
        const peers = standards.filter(o => o.cls === s.cls && o.down && !o.duplicateOf);
        const mode = modeOf(peers, 'down');
        if (mode != null && s.down !== mode && peers.filter(o => o.down === s.down).length <= 2) s.flags.push(`ขาล่อง ${s.down} ลิตร ต่างจากแถวอื่นของรถกลุ่มนี้ (${mode})`);
    }
    standards.forEach((s, n) => { s.id = `r${String(n + 1).padStart(3, '0')}`; });
    if (!standards.length) warnings.push('ไม่พบตารางอัตรา (หาหัวคอลัมน์ "รวมยอดขาขึ้น/ลิตร" ไม่เจอ)');
    return { standards, warnings, skippedSheets };
}

export const standardKey = (s) => `${s.cls}|${s.origin}|${s.destination}|${s.variant}`;

// Sheet "ถอดหางกำแพงเพชร": the trailer (ลูก) is left at the Kamphaeng Phet branch and the head (แม่) runs on
// alone. The current table ("แบบปัจจุบัน") has blocks of  ประเภทรถ | ปลายทาง | จำนวน น้ำมัน (ลิตร) | เบี้ยขับ
// with one แม่ row and one ลูก row each. It gives total liters only — no pump — so trips are flagged for a
// manual check rather than auto-judged. The older left-hand table is only used to point out differences.
export const DROP_CLASS = 'KPP_DROP';
function parseTrailerDropSheet(sheet, rows, startRow) {
    const out = [], old = {};
    rows.forEach((r, i) => {
        (r || []).forEach((cell, c) => {
            // old table rows: ต้นทาง กำแพงเพชร | ปลายทาง | ลิตร | เบี้ยขับ
            if (clean(cell) === 'กำแพงเพชร' && typeof r[c + 2] === 'number' && clean(rows[i - 1]?.[c]) !== 'ประเภทรถ' && canonDestination(r[c + 1])) old[canonDestination(r[c + 1])] = r[c + 2];
            if (clean(cell) !== 'ประเภทรถ' || !clean(r[c + 2]).startsWith('จำนวน')) return;
            const parts = [];
            for (let k = i + 1; k < rows.length; k++) {
                const part = clean(rows[k]?.[c]);
                if (part !== 'แม่' && part !== 'ลูก') break;
                parts.push({ part, destination: clean(rows[k][c + 1]), liters: cellLiters(rows[k][c + 2]), allowance: cellLiters(rows[k][c + 3]), row: startRow + k + 1 });
            }
            const head = parts.find(p => p.part === 'แม่');
            if (!head) return;
            const total = parts.reduce((a, p) => a + p.liters, 0);
            out.push({
                cls: DROP_CLASS, origin: '', destination: canonDestination(head.destination), destinationLabel: `${head.destination} (ถอดหางที่กำแพงเพชร)`,
                variant: 'standard', up: {}, upTotal: total, down: 0, downVictor: 0, parts,
                allowanceText: parts.map(p => `${p.part} ${p.allowance}`).join(' + '), allowance: parts.reduce((a, p) => a + p.allowance, 0),
                distance: '', sheet, row: head.row, flags: [],
            });
        });
    });
    for (const s of out) {
        const head = s.parts.find(p => p.part === 'แม่');
        if (old[s.destination] != null && old[s.destination] !== head.liters) s.flags.push(`ตารางเก่าในชีตเดียวกันให้ แม่→${head.destination} ${old[s.destination]} ลิตร แต่ตารางปัจจุบัน ${head.liters} ลิตร`);
    }
    return out;
}

const tripNoteText = (trip) => `${clean(trip.jobDetail)} ${clean(trip.notes)} ${clean(trip.queueType)}`;
const DROP_AT_KPP = /(ตัดลูก|ถอดหาง|ตัดหาง)\s*(ที่)?\s*(กพ|กำแพง)/;
// "ตัดลูกนว.", "ตัด นว", "ลูกนว.", "ตัวลูกนครสวรรค์", typos "ตักลูกนว." / "นครสรรค์" — trailer left at Nakhon Sawan.
// "นว" must not run on into another Thai letter ("ลูกนวล" is a name, "ตัดนวม" a word).
const NSW_WORD = '(?:นว(?![ก-ฮ])|นครสวรรค์|นครสรรค์|นคสวรรค์)';
const DROP_AT_NSW = new RegExp(`(?:ตัด|ตัก|ถอด)\\s*(?:ลูก|หาง)?\\s*(?:ที่)?\\s*${NSW_WORD}|ลูก\\s*${NSW_WORD}`);
// Where the free text says the trailer was dropped ('นครสวรรค์' | 'กำแพงเพชร' | '').
export const textDropPlace = (trip) => {
    const text = tripNoteText(trip);
    return DROP_AT_KPP.test(text) ? 'กำแพงเพชร' : DROP_AT_NSW.test(text) ? 'นครสวรรค์' : '';
};
// Trips from the Excel template carry an explicit trailerDrop ('ไม่ตัด' | 'นครสวรรค์' | 'กำแพงเพชร'): it wins over
// guessing from the text, including 'ไม่ตัด', which turns off a false match in the job text.
const isDropAtNsw = (trip) => (clean(trip.trailerDrop) ? clean(trip.trailerDrop) === 'นครสวรรค์' : DROP_AT_NSW.test(tripNoteText(trip)));
const isDropAtKpp = (trip) => (clean(trip.trailerDrop) ? clean(trip.trailerDrop) === 'กำแพงเพชร' : DROP_AT_KPP.test(tripNoteText(trip)));
const HEAD_OF = { '18W': '10W', '6W_BOX': '6W' }; // towing class → its head running alone
const TOW_OF = { '10W': '18W', '6W': '6W_BOX' };  // head → towing class

// The same downhill fill must not be counted twice: once on the uphill board row and again on the
// คิวขาล่อง trip. Flag a downhill trip when the same driver's uphill trip in the 5 days before already
// carries a recorded downhill ไพศาล fill. results must be in trip order. allTrips (optional) = every trip
// of the same period, including downhill trips with no fill (they are not in results): a downhill trip of
// that driver in between means the uphill fill belonged to that one, not to this one.
export function flagDoubleDownhill(results, allTrips = []) {
    const downTrips = {};
    for (const t of allTrips) if (clean(t.direction) === 'ขาล่อง' && clean(t.driver)) (downTrips[clean(t.driver)] ||= []).push(t);
    const lastUp = {};
    for (const r of results) {
        const d = clean(r.trip.driver);
        if (!d) continue;
        // Only the driver's most recent uphill trip can share its downhill fill with the next downhill trip.
        if (r.leg !== 'down') { if (r.downhillRecorded) lastUp[d] = r; else delete lastUp[d]; continue; }
        const up = lastUp[d];
        const between = up && (downTrips[d] || []).some(t => t !== r.trip && t.id !== r.trip.id && t.date >= up.trip.date && t.date <= r.trip.date);
        if (up && !between && (Date.parse(r.trip.date) - Date.parse(up.trip.date)) / 86400000 <= 5) {
            // new arrays, not push: results may be shallow copies of cached ones
            r.flags = [...r.flags, `อาจบันทึกขาล่องซ้ำ: เที่ยวขาขึ้น ${up.trip.date} บันทึกไพศาลขาล่อง ${up.downActual} ลิตรไว้แล้ว`];
            up.flags = [...up.flags, `อาจบันทึกขาล่องซ้ำ: เที่ยวขาล่อง ${r.trip.date} มีการเติมไพศาล ${r.actualTotal} ลิตรด้วย`];
            delete lastUp[d]; // one downhill fill pairs with one downhill trip
        }
    }
}

// Liters carried from trip to trip per driver, like TMS DriverPayService: what a driver filled over (+) or
// under (−) the rate is set against the next trip that has a rate. results must be in trip order.
// Adds carryIn / allot / carryOut to each result and returns { [driver]: balance }.
export function applyCarry(results) {
    const bal = {};
    for (const r of results) {
        const d = clean(r.trip.driver);
        // No driver name → nobody to carry to; never pool unnamed trips into one balance.
        if (!d) { r.carryIn = null; r.allot = null; r.carryOut = null; continue; }
        const before = bal[d] || 0;
        r.carryIn = before;
        if (['ok', 'over', 'under', 'wrong-pump'].includes(r.status)) {
            r.allot = Math.max(0, r.expectedTotal - before);
            r.carryOut = before + r.actualTotal - r.expectedTotal;
            bal[d] = r.carryOut;
        } else {
            r.allot = null;
            r.carryOut = before;
        }
    }
    return bal;
}

// Vehicle master row → { cls, model, partner, ngv } (cls null when no rate class applies).
export function vehicleClassOf(vehicle) {
    if (!vehicle) return { cls: null, reason: 'ไม่พบทะเบียนในหน้าทะเบียนรถ' };
    const type = clean(vehicle.vehicleType);
    const model = clean(vehicle.model).toLowerCase();
    if (clean(vehicle.driveType).includes('ร่วม') || type.includes('ร่วม')) return { cls: null, partner: true, reason: 'รถร่วม' };
    if (/ngv/i.test(type)) return { cls: null, reason: 'รถ NGV (ไม่มีอัตราในตาราง)' };
    let cls = null;
    if (type.includes('6ล้อพ่วงตู้') || (type.includes('6') && type.includes('ตู้') && type.includes('พ่วง'))) cls = '6W_BOX';
    else if (/^18/.test(type) || type.includes('พ่วง')) cls = '18W';
    else if (/^10/.test(type) || type.includes('สิบล้อ')) cls = '10W';
    else if (/^6/.test(type) || type.includes('หกล้อ')) cls = '6W';
    return cls ? { cls, model } : { cls: null, reason: `ประเภทรถ "${type || '-'}" จับคู่กับตารางอัตราไม่ได้` };
}

// Trips carry "head-trailer" plates ("70-8157-6436"); the vehicle master lists the head ("70-8157").
// Compare the first two dash-separated parts with spaces removed — never cut by length (plates like
// "1กข-1234" are 8 characters and must stay distinct).
export const normPlate = (p) => clean(p).replace(/\s+/g, '');
export const plateKey = (p) => normPlate(p).split('-').slice(0, 2).join('-');

// Build lookup tables once per render.
export function buildAuditContext(standards, vehicles) {
    const std = {}, drop = {};
    for (const s of standards || []) {
        if (s.cls === DROP_CLASS) { if (s.destination && !s.duplicateOf) drop[s.destination] = s; continue; }
        if (!s.duplicateOf && s.destination) std[standardKey(s)] = s;
    }
    // The vehicle master lists a 10-wheel head twice: "70-8901" (10W, running alone) and "70-8901-8158"
    // (18W = head + trailer). Keep both: exact plate first, then by head.
    const vehFull = {}, vehHead = {};
    for (const v of vehicles || []) {
        const full = normPlate(v.registrationNumber);
        if (!full) continue;
        vehFull[full] = v;
        const h = vehHead[plateKey(full)] ||= { solo: null, pairs: [] };
        if (full.split('-').length >= 3) h.pairs.push(v); else h.solo = v;
    }
    const destinations = new Set((standards || []).map(s => s.destination).filter(Boolean));
    // Downhill (ขาล่อง) liters per truck class: the usual value in the sheet (พ่วง 180, สิบล้อ 120…). The victor
    // figure only counts rows that actually name one ("180 (victor190)").
    const downDefault = {};
    for (const cls of new Set((standards || []).map(s => s.cls))) {
        const rows = (standards || []).filter(s => s.cls === cls && s.down && !s.duplicateOf);
        if (!rows.length) continue;
        const down = modeOf(rows, 'down');
        const victorRows = rows.filter(s => s.downVictor && s.downVictor !== s.down);
        downDefault[cls] = { down, downVictor: victorRows.length ? modeOf(victorRows, 'downVictor') : down };
    }
    return { std, drop, vehFull, vehHead, destinations, downDefault };
}

// Which vehicle-master row describes the truck on this trip. Returns { vehicle, note }.
export function vehicleForPlate(plate, ctx) {
    const full = normPlate(plate);
    const h = ctx.vehHead?.[plateKey(full)];
    if (ctx.vehFull?.[full]) {
        const v = ctx.vehFull[full];
        const soloWithPairs = full.split('-').length < 3 && h?.pairs.length;
        return { vehicle: v, note: soloWithPairs ? `หัวนี้มีหางพ่วงในหน้าทะเบียนรถ (${h.pairs.map(p => p.registrationNumber).join(', ')}) — ถ้าเที่ยวนี้ลากหาง ให้เขียนเลขหางในกระดานคิว` : '', towOption: soloWithPairs ? h.pairs[0] : null };
    }
    if (!h) return { vehicle: null, note: '' };
    // head-trailer = exactly 3 parts ("70-8901-8158"); anything else unknown is left unmatched
    const towing = full.split('-').length === 3;
    if (towing) {
        if (h.pairs.length) return { vehicle: h.pairs[0], note: `ทะเบียนหาง "${clean(plate)}" ไม่ตรงกับในหน้าทะเบียนรถ — ใช้ประเภทของ ${h.pairs[0].registrationNumber}` };
        // A rigid truck towing a trailer runs on the trailer rate (as TMS TrailerClass: 10W → พ่วง, 6W → 6ล้อพ่วงตู้).
        const type = clean(h.solo?.vehicleType);
        const tow = /^10|สิบล้อ/.test(type) ? 'พ่วง' : /^6|หกล้อ/.test(type) ? '6ล้อพ่วงตู้' : type;
        return { vehicle: h.solo ? { ...h.solo, vehicleType: tow } : null, note: `ลากหาง ${clean(plate)} แต่ไม่มีคู่หัว-หางนี้ในหน้าทะเบียนรถ — ตรวจแบบ${tow}` };
    }
    const vehicle = h.solo || h.pairs[0];
    const both = h.solo && h.pairs.length;
    const note = both ? `หัวนี้มีหางพ่วงในหน้าทะเบียนรถ (${h.pairs.map(v => v.registrationNumber).join(', ')}) — ถ้าเที่ยวนี้ลากหาง ให้เขียนเลขหางในกระดานคิว` : '';
    return { vehicle, note, towOption: both ? h.pairs[0] : null };
}

// Downhill trip (from the Transport คิวขาล่อง sheet) with liters filled: compare with the class's downhill rate.
function auditDownhillTrip(trip, ctx, tolerance, actual, actualTotal, vc, baht, bahtTotal, plateNote) {
    const base = { trip, leg: 'down', cls: vc.cls, model: vc.model || '', origin: clean(trip.origin), destination: clean(trip.destination), actual, actualTotal, baht, bahtTotal, flags: plateNote ? [plateNote] : [], downStatus: '' };
    if (bahtTotal) base.flags.push(`เติมเป็นเงิน ${bahtTotal.toLocaleString()} บาท (ไม่มีอัตราเป็นเงิน)`);
    if (!actualTotal) return { ...base, status: 'no-rate', reason: 'เติมเป็นเงินอย่างเดียว ไม่มีลิตรให้เทียบ', diff: 0, expected: {}, expectedTotal: 0 };
    if (!vc.cls) return { ...base, status: 'no-rate', reason: vc.reason, diff: 0, expected: {}, expectedTotal: 0 };
    const def = ctx.downDefault?.[vc.cls];
    if (!def) return { ...base, status: 'no-rate', reason: `ไม่มีอัตราขาล่องของ ${AUDIT_CLASSES[vc.cls]}`, diff: 0, expected: {}, expectedTotal: 0 };
    const down = vc.model === 'victor' ? def.downVictor : def.down;
    const expected = { [PAISAN]: down };
    const perPump = [...new Set([PAISAN, ...Object.keys(actual)])].map(k => ({ key: k, expected: expected[k] || 0, actual: Number(actual[k]) || 0 })).map(p => ({ ...p, diff: p.actual - p.expected }));
    const off = perPump.filter(p => !p.expected && p.actual);
    if (off.length) base.flags.push(`เติมปั๊มอื่นแทนไพศาล: ${off.map(p => `${pumpThai(p.key)} ${p.actual} ลิตร`).join(', ')}`);
    const diff = actualTotal - down;
    const status = diff > tolerance ? 'over' : diff < -tolerance ? 'under' : perPump.some(p => Math.abs(p.diff) > tolerance) ? 'wrong-pump' : 'ok';
    return { ...base, status, expected, expectedTotal: down, perPump, diff, down };
}

// Audit one trip. Returns null when the trip is out of scope (partner truck, no fuel at all).
// When the board wrote only the head plate of a truck that also runs with a trailer, check both ways:
// if the fill matches the trailer (พ่วง) rate in total but not the solo rate, the trailer was most likely
// just left off the board — use the trailer rate and say so.
export function auditTrip(trip, ctx, tolerance = 0) {
    const { towOption } = vehicleForPlate(trip.plate, ctx);
    const dropNsw = clean(trip.direction) === 'ขาขึ้น' && isDropAtNsw(trip);
    // A trailer was dropped, so the truck WAS towing: judge it as its head+trailer pair (keeps e.g. a victor model).
    if (dropNsw && towOption) {
        const r = auditTripAs(trip, ctx, tolerance, towOption, true);
        if (r) r.flags = r.flags.filter(f => !f.startsWith('หัวนี้มีหางพ่วง'));
        return r;
    }
    const first = auditTripAs(trip, ctx, tolerance, null, dropNsw);
    if (!first || !towOption || first.leg === 'down' || first.status === 'no-rate' || first.status === 'manual' || dropNsw) return first;
    if (Math.abs(first.diff) <= tolerance) return first;
    const towed = auditTripAs(trip, ctx, tolerance, towOption, false);
    if (towed && towed.status !== 'no-rate' && Math.abs(towed.diff) <= tolerance) {
        towed.flags = [`กระดานคิวเขียนแค่ทะเบียนหัว แต่ลิตรรวมตรงกับอัตรา${AUDIT_CLASSES[towed.cls] || ''}ของ ${towOption.registrationNumber} — ตรวจแบบลากหาง (ควรเขียนเลขหางในกระดาน)`,
            ...towed.flags.filter(f => !f.startsWith('หัวนี้มีหางพ่วง'))];
        return towed;
    }
    return first;
}

function auditTripAs(trip, ctx, tolerance, forcedVehicle, dropNswIn) {
    const dir = clean(trip.direction);
    if (dir !== 'ขาขึ้น' && dir !== 'ขาล่อง') return null;
    const actual = { ...(trip.gasStationsLiters || {}) };
    const baht = { ...(trip.gasStationsBaht || {}) };
    const actualTotal = Object.values(actual).reduce((a, b) => a + (Number(b) || 0), 0);
    const bahtTotal = Object.values(baht).reduce((a, b) => a + (Number(b) || 0), 0);
    if (!actualTotal && !bahtTotal) return null;
    const found = vehicleForPlate(trip.plate, ctx);
    const vehicle = forcedVehicle || found.vehicle;
    const plateNote = forcedVehicle ? '' : found.note;
    const vc = vehicleClassOf(vehicle);
    if (vc.partner) return null;
    if (dir === 'ขาล่อง') return auditDownhillTrip(trip, ctx, tolerance, actual, actualTotal, vc, baht, bahtTotal, plateNote);
    const destination = canonDestination(trip.destination);
    const origin = canonOrigin(trip.origin);
    const base = { trip, cls: vc.cls, model: vc.model || '', origin, destination, actual, actualTotal, baht, bahtTotal, flags: plateNote ? [plateNote] : [] };
    if (bahtTotal) base.flags.push(`เติมเป็นเงิน ${bahtTotal.toLocaleString()} บาท (ไม่มีอัตราเป็นเงิน)`);
    const none = (reason) => ({ ...base, status: 'no-rate', reason, diff: 0, expected: {}, expectedTotal: 0, downStatus: '' });
    // Trailer left at Kamphaeng Phet: the special table has total liters only (no pump) → manual check.
    if (isDropAtKpp(trip)) {
        const d = ctx.drop?.[destination];
        // The drop table has no downhill figure, but the ไพศาล downhill fill is often written on this trip —
        // name it so a normal downhill fill is not read as an overfill.
        const paisan = Number(actual[PAISAN]) || 0;
        const paisanNote = paisan ? ` · ในยอดนี้มีปั๊มไพศาล ${paisan} ลิตร ซึ่งอาจเป็นการเติมขาล่อง` : '';
        const reason = d
            ? `เที่ยวถอดหางที่กำแพงเพชร — อัตราพิเศษ ${d.parts.map(p => `${p.part}→${p.destination} ${p.liters}`).join(' + ')} = ${d.upTotal} ลิตร (ไม่ระบุปั๊ม ไม่รวมขาล่อง) · เติมจริงรวม ${actualTotal} ลิตร (ต่าง ${actualTotal - d.upTotal > 0 ? '+' : ''}${actualTotal - d.upTotal})${paisanNote}`
            : `เที่ยวถอดหางที่กำแพงเพชร — ตารางถอดหางไม่มีปลายทาง "${clean(trip.destination) || '-'}" ตรวจกับหัวหน้างาน${paisanNote}`;
        return { ...base, status: 'manual', reason, diff: d ? actualTotal - d.upTotal : 0, expected: {}, expectedTotal: d ? d.upTotal : 0, downStatus: '' };
    }
    if (!vc.cls) return none(vc.reason);
    if (!origin) return none(`ต้นทาง "${clean(trip.origin)}" เป็นสาขาอื่น (ตารางมีแค่ กรุงเทพ / นครสวรรค์)`);
    if (!destination) return none(`ปลายทาง "${clean(trip.destination) || '-'}" ไม่อยู่ในตารางอัตรา`);
    // Trailer dropped at Nakhon Sawan (user rule, 17 Nov 2026): the head runs on alone at its own rate — a พ่วง's
    // head is a 10-wheeler, a 6-wheel box trailer's head a 6-wheeler — and comes back towing the trailer, so
    // the downhill fill follows the towing rate (seen on real trips: จิตตมาศ 50 + ไพศาล 180 to กำแพงเพชร).
    // Not when the trip ends at นครสวรรค์ itself: the trailer was towed the whole way.
    const pick = (cls, variant) => ctx.std[`${cls}|${origin}|${destination}|${variant}`];
    const pickV = (cls) => (vc.model === 'sino' && pick(cls, 'sino')) || pick(cls, 'standard');
    let dropNsw = !!dropNswIn && destination !== 'นครสวรรค์' && (vc.cls in HEAD_OF || vc.cls in TOW_OF);
    let upCls = dropNsw ? (HEAD_OF[vc.cls] || vc.cls) : vc.cls;
    const downCls = dropNsw ? (TOW_OF[upCls] || vc.cls) : vc.cls;
    if (dropNsw && !pickV(upCls)) {
        // No head-alone rate for this route → fall back to the towing rate, and say so.
        base.flags.push(`ตัดหางที่นครสวรรค์ แต่ตารางไม่มีอัตรา${AUDIT_CLASSES[upCls]}ไป ${destination} — ใช้อัตรา${AUDIT_CLASSES[downCls]}แทน`);
        dropNsw = false; upCls = downCls; base.rateCls = downCls;
    }
    const s = pickV(upCls);
    if (!s) return none(`ไม่มีอัตรา ${AUDIT_CLASSES[upCls]} ${origin} → ${destination}`);
    // Downhill comes from the towing class's row for THIS route only — a class-wide default may not match
    // the route, so when the row is missing the downhill fill is left unchecked (and the note says so).
    const sDown = upCls === downCls ? s : (pickV(downCls) || { down: 0, downVictor: 0 });
    if (dropNsw) {
        base.rateCls = upCls;
        base.flags.push(sDown.down
            ? `ตัดหางที่นครสวรรค์ — ขาขึ้นคิดอัตรา${AUDIT_CLASSES[upCls]} (ตัวแม่) ขาล่องคิดอัตรา${AUDIT_CLASSES[downCls]} (ลากหางกลับ)`
            : `ตัดหางที่นครสวรรค์ — ขาขึ้นคิดอัตรา${AUDIT_CLASSES[upCls]} (ตัวแม่) · ไม่มีอัตราขาล่องของ${AUDIT_CLASSES[downCls]} จึงไม่ได้ตรวจขาล่อง`);
    }

    const down = vc.model === 'victor' ? sDown.downVictor : sDown.down;
    const expected = { ...s.up };
    // The downhill ไพศาล fill is often written on the uphill trip, but not always. Assume whichever reading
    // (downhill included or not) is closer to what was filled at ไพศาล, so an uphill overfill there is not
    // mistaken for a downhill fill that came up short.
    const upPaisan = s.up[PAISAN] || 0;
    const actPaisan = Number(actual[PAISAN]) || 0;
    let downhillRecorded = false;
    if (down) {
        if (Math.abs(actPaisan - (upPaisan + down)) < Math.abs(actPaisan - upPaisan)) { expected[PAISAN] = upPaisan + down; downhillRecorded = true; }
        else base.flags.push(`ขาล่องไพศาล ${down} ลิตร ไม่ได้บันทึกในเที่ยวนี้`);
    }
    const expectedTotal = Object.values(expected).reduce((a, b) => a + b, 0);
    const pumps = [...new Set([...Object.keys(expected), ...Object.keys(actual)])];
    const perPump = pumps.map(k => ({ key: k, expected: expected[k] || 0, actual: Number(actual[k]) || 0 }))
        .map(p => ({ ...p, diff: p.actual - p.expected }));
    const offPumps = perPump.filter(p => !p.expected && p.actual);
    if (offPumps.length) base.flags.push(`เติมปั๊มที่ไม่อยู่ในอัตรา: ${offPumps.map(p => `${pumpThai(p.key)} ${p.actual} ลิตร`).join(', ')}`);
    const diff = actualTotal - expectedTotal;
    let status;
    if (diff > tolerance) status = 'over';
    else if (diff < -tolerance) status = 'under';
    else if (perPump.some(p => Math.abs(p.diff) > tolerance)) status = 'wrong-pump';
    else status = 'ok';
    if (s.flags.length) base.flags.push(`อัตราแถวนี้มีข้อสังเกต: ${s.flags.join('; ')}`);
    // Downhill (ขาล่อง) check: the only downhill liters in the data are the ไพศาล fills written on uphill trips
    // (the Transport ขาล่อง sheet carries no liters).
    const downActual = downhillRecorded ? actPaisan - upPaisan : 0;
    let downStatus = '';
    if (down) {
        if (!downhillRecorded) downStatus = 'missing';
        else if (downActual - down > tolerance) downStatus = 'over';
        else if (downActual - down < -tolerance) downStatus = 'under';
        else downStatus = 'ok';
    }
    return { ...base, status, standard: s, expected, expectedTotal, perPump, diff, down, downhillRecorded, downActual, downStatus };
}

export const DOWN_STATUS = {
    'ok': { label: 'ตรง', cls: 'text-green-400' },
    'over': { label: 'เกิน', cls: 'text-red-400' },
    'under': { label: 'ขาด', cls: 'text-amber-300' },
    'missing': { label: 'ไม่ได้บันทึก', cls: 'text-gray-400' },
};

export const AUDIT_STATUS = {
    'ok': { label: 'ตรง', cls: 'bg-green-700' },
    'over': { label: 'เติมเกิน', cls: 'bg-red-700' },
    'under': { label: 'เติมขาด', cls: 'bg-amber-700' },
    'wrong-pump': { label: 'ผิดปั๊ม', cls: 'bg-purple-700' },
    'manual': { label: 'ต้องตรวจเอง', cls: 'bg-sky-700' },
    'no-rate': { label: 'ไม่มีอัตรา', cls: 'bg-slate-600' },
};
