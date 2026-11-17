// Fuel audit: check each truck trip's actual fills against the official rate sheet
// ("อัตราเติมน้ำมันประจำ ปั๊ม ต่างๆ"). Pure functions (no DOM/Firebase) so they can be tested in Node.

import { parseNum, clean } from './parseUtils.js';

export const AUDIT_CLASSES = {
    '10W': 'สิบล้อ',
    '18W': 'พ่วง',
    '6W': 'หกล้อ',
    '6W_BOX': 'หกล้อ isuzu + หางตู้',
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
        if (!(rows || []).some(r => (r || []).some(c => clean(c) === TOTAL_UP_HEADER))) { skippedSheets.push(sheet); continue; }
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
        const same = list.every(s => JSON.stringify(s.up) === JSON.stringify(list[0].up) && s.down === list[0].down && s.downVictor === list[0].downVictor);
        list.forEach((s, n) => {
            s.flags.push(same ? 'แถวซ้ำ (ค่าเท่ากัน)' : `ขัดกับแถว ${list.filter(o => o !== s).map(o => `${o.sheet} แถว ${o.row}`).join(', ')}`);
            if (n > 0) s.duplicateOf = `${list[0].sheet}!${list[0].row}`; // audit uses the first row
        });
    }
    // Downhill liters that differ from the rest of the same class (e.g. 179 vs 180) — likely a typo.
    const downByCls = {};
    for (const s of standards) if (s.down) ((downByCls[s.cls] ||= {})[s.down] = (downByCls[s.cls][s.down] || 0) + 1);
    for (const s of standards) {
        const counts = downByCls[s.cls] || {};
        const mode = Object.entries(counts).sort((x, y) => y[1] - x[1])[0];
        if (mode && s.down && String(s.down) !== mode[0] && counts[s.down] <= 2) s.flags.push(`ขาล่อง ${s.down} ลิตร ต่างจากแถวอื่นของรถกลุ่มนี้ (${mode[0]})`);
    }
    standards.forEach((s, n) => { s.id = `r${String(n + 1).padStart(3, '0')}`; });
    if (!standards.length) warnings.push('ไม่พบตารางอัตรา (หาหัวคอลัมน์ "รวมยอดขาขึ้น/ลิตร" ไม่เจอ)');
    return { standards, warnings, skippedSheets };
}

export const standardKey = (s) => `${s.cls}|${s.origin}|${s.destination}|${s.variant}`;

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
const plateKey = (p) => clean(p).replace(/\s+/g, '').split('-').slice(0, 2).join('-');

// Build lookup tables once per render.
export function buildAuditContext(standards, vehicles) {
    const std = {};
    for (const s of standards || []) if (!s.duplicateOf && s.destination) std[standardKey(s)] = s;
    const veh = {};
    for (const v of vehicles || []) { const k = plateKey(v.registrationNumber); if (k) veh[k] = v; }
    const destinations = new Set((standards || []).map(s => s.destination).filter(Boolean));
    return { std, veh, destinations };
}

// Audit one trip. Returns null when the trip is out of scope (not uphill, partner truck, no fuel at all).
export function auditTrip(trip, ctx, tolerance = 0) {
    if (clean(trip.direction) !== 'ขาขึ้น') return null;
    const actual = { ...(trip.gasStationsLiters || {}) };
    const baht = { ...(trip.gasStationsBaht || {}) };
    const actualTotal = Object.values(actual).reduce((a, b) => a + (Number(b) || 0), 0);
    const bahtTotal = Object.values(baht).reduce((a, b) => a + (Number(b) || 0), 0);
    if (!actualTotal && !bahtTotal) return null;
    const vc = vehicleClassOf(ctx.veh[plateKey(trip.plate)]);
    if (vc.partner) return null;
    const destination = canonDestination(trip.destination);
    const origin = canonOrigin(trip.origin);
    const base = { trip, cls: vc.cls, model: vc.model || '', origin, destination, actual, actualTotal, baht, bahtTotal, flags: [] };
    if (bahtTotal) base.flags.push(`เติมเป็นเงิน ${bahtTotal.toLocaleString()} บาท (ไม่มีอัตราเป็นเงิน)`);
    const none = (reason) => ({ ...base, status: 'no-rate', reason, diff: 0, expected: {}, expectedTotal: 0 });
    if (!vc.cls) return none(vc.reason);
    if (!origin) return none(`ต้นทาง "${clean(trip.origin)}" เป็นสาขาอื่น (ตารางมีแค่ กรุงเทพ / นครสวรรค์)`);
    if (!destination) return none(`ปลายทาง "${clean(trip.destination) || '-'}" ไม่อยู่ในตารางอัตรา`);
    const pick = (variant) => ctx.std[`${vc.cls}|${origin}|${destination}|${variant}`];
    const s = (vc.model === 'sino' && pick('sino')) || pick('standard');
    if (!s) return none(`ไม่มีอัตรา ${AUDIT_CLASSES[vc.cls]} ${origin} → ${destination}`);

    const down = vc.model === 'victor' ? s.downVictor : s.down;
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
    if (offPumps.length) base.flags.push(`เติมปั๊มที่ไม่อยู่ในอัตรา: ${offPumps.map(p => `${p.key} ${p.actual} ลิตร`).join(', ')}`);
    const diff = actualTotal - expectedTotal;
    let status;
    if (diff > tolerance) status = 'over';
    else if (diff < -tolerance) status = 'under';
    else if (perPump.some(p => Math.abs(p.diff) > tolerance)) status = 'wrong-pump';
    else status = 'ok';
    if (s.flags.length) base.flags.push(`อัตราแถวนี้มีข้อสังเกต: ${s.flags.join('; ')}`);
    return { ...base, status, standard: s, expected, expectedTotal, perPump, diff, down, downhillRecorded };
}

export const AUDIT_STATUS = {
    'ok': { label: 'ตรง', cls: 'bg-green-700' },
    'over': { label: 'เติมเกิน', cls: 'bg-red-700' },
    'under': { label: 'เติมขาด', cls: 'bg-amber-700' },
    'wrong-pump': { label: 'ผิดปั๊ม', cls: 'bg-purple-700' },
    'no-rate': { label: 'ไม่มีอัตรา', cls: 'bg-slate-600' },
};
