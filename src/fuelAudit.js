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
    'EMPTY_TRAILER_KPP': 'ลากหางเปล่าไปตัดที่กำแพงเพชร',
};
// The sheet's written rule "พ่วง ลากหางเปล่า ไปตัดไว้ที่ กำแพง … สิบล้อ + เพิ่ม N ลิตร" (no data row) is kept as one
// pseudo-row of this class with extraLiters = N.
export const EMPTY_KPP_CLASS = 'EMPTY_TRAILER_KPP';

// Rate-sheet pump header fragment → station key used everywhere else in the app.
const PUMP_HEADERS = [
    ['จิตตมาศ', 'jittamas_bangpain'],
    ['ศรีนคร', 'srinakorn_klongklung'],
    ['ไพศาล', 'daopaisan_nakhonsawan'],
    ['นิสสัน', 'nissan_singburi'],
    ['ชำนาญ', 'chamnan'],
];
export const PAISAN = 'daopaisan_nakhonsawan';
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
    // Added by the team's sheet update of 17 Nov 2026. Last, so multi-drop labels such as "พิจิตร+นครสวรรค์" keep
    // the place they had before.
    ['พิจิตร', 'พิจิตร'], ['สุโขทัย', 'สุโขทัย'],
];
// "(ตีเข้า กทม)" / "(ตีเข้า นครสวรรค์)" (sheet update 17 Nov 2026): two rates for one place, by where the truck heads
// after the drop. Trips do not record it, so the audit never picks one of them (see auditTripAs).
// Any "ตีเข้า" label is a via variant — an unreadable target becomes via_other, never the plain rate.
// A sino row of the same place keeps its own key (variant + "_sino").
const VIA = { via_bkk: 'ตีเข้า กทม', via_nsw: 'ตีเข้า นครสวรรค์', via_other: 'ตีเข้า (อ่านที่หมายไม่ออก)' };
const VIA_TEXT = /\(?\s*ตีเข้า[^)]*\)?/;
const viaOf = (label) => {
    const m = clean(label).match(/ตีเข้า\s*([^)]*)/);
    if (!m) return '';
    const t = m[1].replace(/\s+/g, '');
    return /นครสวรรค์|นครสรรค์|นคสวรรค์|^นว/.test(t) ? 'via_nsw' : /กทม|กรุงเทพ/.test(t) ? 'via_bkk' : 'via_other';
};
const viaBase = (variant) => String(variant || '').replace(/_sino$/, '');
// A data row's first cell: "จาก กรุงเทพ" (older layout) or just "กรุงเทพ" / "นครสวรรค์" (layout of 17 Nov 2026),
// typos included. A blank first cell under an origin is a merged cell → same origin as the row above.
const ORIGIN_CELL = /^จาก|^(กรุงเทพ|กทม|นครสวรรค์|นครสรรค์|นคสวรรค์|นว)/;
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
// { id, cls, origin, destination, destinationLabel, variant:'standard'|'sino'|'via_bkk'|'via_nsw'|'via_other' (+'_sino'), up:{station:L}, upTotal,
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
        let cls = null, header = null, lastOrigin = '';
        rows.forEach((r, i) => {
            const a = clean(r?.[0]);
            if (a === 'สถานที่') { cls = classOfTitle(r[2]); header = null; lastOrigin = ''; return; }
            // Written rule with no data row: "พ่วง ลากหางเปล่า ไปตัดไว้ที่ กำแพง น้ำมันจ่ายเป็นรถสิบล้อ+เพิ่ม 20 ลิตร".
            // Looked for in any cell of the row (merged cells), Thai digits accepted; a rule sentence without a
            // readable number is reported instead of silently dropped.
            const ruleCell = (r || []).map(clean).find(c => /ลากหางเปล่า/.test(c) && /กพ|กำแพง/.test(c));
            if (ruleCell) {
                const text = ruleCell.replace(/[๐-๙]/g, (d) => String('๐๑๒๓๔๕๖๗๘๙'.indexOf(d))).replace(/\s+/g, ' ');
                const m = text.match(/(?:เพิ่ม|บวก|\+)\s*(\d+)\s*ลิตร/);
                if (m) standards.push({ cls: EMPTY_KPP_CLASS, origin: '', destination: '', destinationLabel: '', variant: 'standard', up: {}, upTotal: 0, down: 0, downVictor: 0,
                    extraLiters: parseInt(m[1], 10), ruleText: text, allowanceText: '', allowance: 0, distance: '', sheet, row: startRow + i + 1, flags: [] });
                else warnings.push(`${sheet} แถว ${startRow + i + 1}: มีกฎลากหางเปล่าแต่อ่านจำนวนลิตรที่เพิ่มไม่ได้`);
                return;
            }
            if (a === 'ต้นทาง') { header = (r || []).map(clean); lastOrigin = ''; return; }
            if (!header || !cls) return;
            const iTotal = header.indexOf(TOTAL_UP_HEADER);
            if (iTotal < 0) return;
            const rowNo = startRow + i + 1;
            let originText = a;
            if (!ORIGIN_CELL.test(a)) {
                // Not an origin: skip notes / blank rows; a row with a place and figures is never dropped silently.
                if (!clean(r[1]) || (clean(r[iTotal]) === '' && clean(r[iTotal + 1]) === '')) return;
                if (!a && lastOrigin) originText = lastOrigin;
                else { warnings.push(`${sheet} แถว ${rowNo}: อ่านต้นทาง "${a || '(ว่าง)'}" ไม่ออก — ไม่ได้นำเข้าแถวนี้`); return; }
            } else lastOrigin = a;
            if (!clean(r[1])) {
                warnings.push(`${sheet} แถว ${rowNo}: มีต้นทาง "${a}" แต่ไม่มีปลายทาง — ไม่ได้นำเข้าแถวนี้`);
                return;
            }
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
            const via = viaOf(destinationLabel), sino = /sino/i.test(allowanceText);
            const std = {
                // "(ตีเข้า นครสวรรค์)" is not the place: strip it before reading the destination.
                cls, origin: canonOrigin(originText), destination: canonDestination(destinationLabel.replace(VIA_TEXT, ' ')), destinationLabel,
                variant: via ? (sino ? `${via}_sino` : via) : sino ? 'sino' : 'standard',
                up, upTotal: Object.values(up).reduce((x, y) => x + y, 0),
                down, downVictor: victorM ? parseInt(victorM[1], 10) : down,
                allowanceText, allowance: parseNum((allowanceText.match(/เบี้ยขับ\s*([\d,]+)/) || ['', ''])[1]),
                distance: iDist >= 0 ? clean(r[iDist]) : '',
                sheet, row: startRow + i + 1, flags: [],
            };
            const sheetTotal = cellLiters(r[iTotal]);
            if (sheetTotal !== std.upTotal) std.flags.push(`ผลรวมรายปั๊ม ${std.upTotal} ไม่เท่าช่องรวม ${sheetTotal}`);
            if (!std.destination) std.flags.push('อ่านปลายทางไม่ออก');
            if (via === 'via_other') std.flags.push('อ่านไม่ออกว่า "ตีเข้า" ที่ไหน (กทม / นครสวรรค์) — ตรวจในไฟล์ เที่ยวปลายทางนี้จะเป็น "ตรวจเอง"');
            if (std.origin === 'นครสวรรค์' && std.up.nissan_singburi) std.flags.push('ต้นทางนครสวรรค์แต่เติมปั๊มนิสสัน (สิงห์บุรี) — ตรวจคอลัมน์ในไฟล์');
            standards.push(std);
        });
    }
    // Same class + origin + destination + variant twice with different liters → the sheet contradicts itself.
    const byKey = {};
    for (const s of standards) (byKey[standardKey(s)] ||= []).push(s);
    for (const list of Object.values(byKey)) {
        if (list.length < 2) continue;
        const sig = (s) => JSON.stringify([s.up, s.down, s.downVictor, s.upTotal, s.extraLiters || 0, (s.parts || []).map(p => [p.part, p.destination, p.liters, p.allowance])]);
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

// Rate sets by effective date (user decision 17 Nov 2026: the team's updated sheet applies from 13 Nov 2026; earlier
// trips keep the set that was in force). versions: [{ effectiveFrom: 'YYYY-MM-DD' | '', standards }], '' = since the
// beginning. A trip uses the version with the latest effectiveFrom on or before its date; a trip without a date uses
// the newest. null = no set covers that date.
export function rateVersionFor(versions, date) {
    const d = clean(date);
    let best = null;
    for (const v of versions || []) {
        if (!v) continue;
        const eff = clean(v.effectiveFrom);
        if (d && eff > d) continue;
        if (!best || eff > clean(best.effectiveFrom)) best = v;
    }
    return best;
}

// One audit context per rate version, built once: returns (date) => ctx. A date no version covers gets an empty
// context (every trip "no rate"), never another period's rates.
export function makeRateContexts(versions, vehicles, opts = {}) {
    const cache = new Map();
    const get = (v) => {
        if (!cache.has(v)) cache.set(v, buildAuditContext(v ? v.standards : [], vehicles, opts));
        return cache.get(v);
    };
    return (date) => get(rateVersionFor(versions, date));
}

// What saving a new set that starts on `effectiveFrom` does to the sets already stored: the set in force is archived
// (fuelRateStandards/history); every stored set that starts on or after the new date is fully overridden → retired (kept as a
// backup, never used by the audit). effectiveFrom '' = the new set replaces every period.
// current: { effectiveFrom } | null; history: [{ id, effectiveFrom, retired }].
export function planRateSave(current, history, effectiveFrom) {
    const eff = clean(effectiveFrom);
    return {
        archiveCurrent: !!current,
        archivedRetired: !!current && clean(current.effectiveFrom) >= eff,
        retire: (history || []).filter(h => h && !h.retired && clean(h.effectiveFrom) >= eff).map(h => h.id),
    };
}

// The whole new value of the fuelRateStandards node { meta, rows, history: { key: { meta, rows } } } — written by
// ONE transaction, so the new set, the archived set and the retire marks land together or not at all, and a second
// save started from the same set is refused (the transaction re-runs against the stored value).
// cur: the stored node value (or null); rows: the new rows by id; meta must carry saveId; archiveKey: a fresh key.
// Returns null when the stored set is not the one the preview was compared with (expectedSaveId) → abort.
// Values never contain undefined (Firebase rejects it).
export function buildRateSaveValue(cur, rows, meta, archiveKey, expectedSaveId = '') {
    if (((cur && cur.meta && cur.meta.saveId) || '') !== expectedSaveId) return null;
    const hasCur = !!(cur && cur.rows);
    const curEff = (hasCur && cur.meta?.effectiveFrom) || '';
    const history = { ...((cur && cur.history) || {}) };
    const plan = planRateSave(hasCur ? { effectiveFrom: curEff } : null,
        Object.entries(history).map(([id, h]) => ({ id, effectiveFrom: h?.meta?.effectiveFrom || '', retired: !!h?.meta?.retired })), meta.effectiveFrom || '');
    for (const id of plan.retire) history[id] = { ...history[id], meta: { ...(history[id].meta || {}), retired: true } };
    if (plan.archiveCurrent) {
        history[archiveKey] = {
            meta: { ...(cur.meta || {}), effectiveFrom: curEff, archivedAt: meta.importedAt || '', replacedBy: meta.fileName || '', replacedFrom: meta.effectiveFrom || '', retired: plan.archivedRetired },
            rows: cur.rows,
        };
    }
    const out = { meta: { ...meta, effectiveFrom: meta.effectiveFrom || '' }, rows };
    if (Object.keys(history).length) out.history = history;
    return out;
}

// Rates now in the system that an incoming file does not have — saving replaces the whole set, so these would be
// lost (e.g. the team's 17 Nov update overwrote the "ลากหางเปล่า +20 ลิตร" sentence and dropped วิเชียรบุรี).
// Returns readable labels for the import preview.
export function ratesMissingFrom(current, incoming) {
    const have = new Set((incoming || []).filter(s => !s.duplicateOf).map(standardKey));
    const seen = new Set(), out = [];
    for (const s of current || []) {
        const k = standardKey(s);
        if (s.duplicateOf || have.has(k) || seen.has(k)) continue;
        seen.add(k);
        out.push(s.cls === EMPTY_KPP_CLASS ? `กฎ${AUDIT_CLASSES[EMPTY_KPP_CLASS]} +${s.extraLiters} ลิตร`
            : `${AUDIT_CLASSES[s.cls] || s.cls} ${s.origin ? `${s.origin} → ` : ''}${s.destinationLabel || s.destination}${s.variant === 'sino' ? ' (sino)' : ''}`);
    }
    return out;
}

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
    // Returns the template's "ตัดหางที่" option the text describes (used to warn when that cell was left empty).
    return EMPTY_LEFT_AT_KPP.test(text) ? EMPTY_KPP_DROP : DROP_AT_KPP.test(text) ? 'กำแพงเพชร' : DROP_AT_NSW.test(text) ? 'นครสวรรค์' : '';
};
// Trips from the Excel template carry an explicit trailerDrop ('ไม่ตัด' | 'นครสวรรค์' | 'กำแพงเพชร'): it wins over
// guessing from the text, including 'ไม่ตัด', which turns off a false match in the job text.
const isDropAtNsw = (trip) => (clean(trip.trailerDrop) ? clean(trip.trailerDrop) === 'นครสวรรค์' : DROP_AT_NSW.test(tripNoteText(trip)));
const isDropAtKpp = (trip) => (clean(trip.trailerDrop) ? clean(trip.trailerDrop) === 'กำแพงเพชร' : DROP_AT_KPP.test(tripNoteText(trip)));
// Empty trailer towed to กพ. and LEFT there (sheet rule row "พ่วง ลากหางเปล่า ไปตัดไว้ที่ กำแพง"). The text must say both
// "empty trailer" and "left/dropped at กพ." — "ลากลูกเปล่าไปขึ้นน้ำตาล" (towed empty to load sugar) is not this rule.
export const EMPTY_KPP_DROP = 'กำแพงเพชร (หางเปล่า)'; // template "ตัดหางที่" option
// "empty trailer" FOLLOWED closely by "left/dropped at กพ." in the same phrase ("ลากหางเปล่า ไปตัดไว้ที่ กพ"), so
// "ตัดลูก กพ … ขากลับลากลูกเปล่า" (loaded trailer dropped, came back empty) does not match.
const EMPTY_LEFT_AT_KPP = /(?:หาง|ลูก)\s*เปล่า[^/,;]{0,25}?(?:ตัด|ถอด|ไว้|ทิ้ง)\s*(?:ลูก|หาง)?\s*(?:ไว้)?\s*(?:ที่)?\s*(?:กพ|กำแพง)/;
const isEmptyTrailerToKpp = (trip) => {
    const td = clean(trip.trailerDrop);
    if (td) return td === EMPTY_KPP_DROP;
    return EMPTY_LEFT_AT_KPP.test(tripNoteText(trip));
};
const HEAD_OF = { '18W': '10W', '6W_BOX': '6W' }; // towing class → its head running alone
const TOW_OF = { '10W': '18W', '6W': '6W_BOX' };  // head → towing class

// The same downhill fill must not be counted twice: once on the uphill board row and again on the
// คิวขาล่อง trip. Flag a downhill trip when the same driver's uphill trip in the 5 days before already
// carries a recorded downhill ไพศาล fill. results must be in trip order. allTrips (optional) = every trip
// of the same period, including downhill trips with no fill (they are not in results): a downhill trip of
// that driver in between means the uphill fill belonged to that one, not to this one.
// Pair by driver name AND by head plate: the daily board writes first names only while the Transport file and the
// template use full names, so the same driver can appear under two names (seen 17 Nov 2026: 0 of 38 board names
// match the driver master exactly).
// Statuses that hold a checked figure (and feed carry-over).
export const CHECKED = ['ok', 'over', 'under', 'wrong-pump'];
// A downhill fill at another pump after the uphill trip already holds the downhill ไพศาล fill → manual check.
export function markExtraDownhill(r, up) {
    Object.assign(r, {
        status: 'manual', diff: 0,
        reason: `เที่ยวขาขึ้น ${up.trip.date} บันทึกไพศาลขาล่อง ${up.downActual} ลิตรไว้แล้ว · ขาล่องเติมเพิ่ม ${r.actualTotal} ลิตร${r.trip.fuelText ? ` (${r.trip.fuelText})` : ''} — ตารางอัตรายังไม่มีกฎการเติมเพิ่มขาล่อง ตรวจเอง`,
    });
    return r;
}
const pairKeys = (t) => [clean(t.driver) && `D:${clean(t.driver)}`, plateKey(t.plate) && `P:${plateKey(t.plate)}`].filter(Boolean);
export function flagDoubleDownhill(results, allTrips = []) {
    const downTrips = {};
    for (const t of allTrips) if (clean(t.direction) === 'ขาล่อง') for (const k of pairKeys(t)) (downTrips[k] ||= []).push(t);
    const lastUp = {};
    for (const r of results) {
        const ks = pairKeys(r.trip);
        if (!ks.length) continue;
        // Only the most recent uphill trip (of this driver / truck) can share its downhill fill with the next downhill
        // trip. A new uphill replaces the previous one under ALL of that one's keys, so a driver who changed trucks
        // does not leave a stale entry on the old truck.
        if (r.leg !== 'down') {
            for (const k of ks) { const old = lastUp[k]; if (old) for (const x of pairKeys(old.trip)) if (lastUp[x] === old) delete lastUp[x]; }
            if (r.downhillRecorded) for (const k of ks) lastUp[k] = r;
            continue;
        }
        // Column L names the uphill queue this downhill belongs to ("(คิว<date>…)"): that exact reference wins over
        // guessing by "most recent uphill" — reconcileColumnL checks it. Guessing here paired many downhill trips with
        // the truck's NEXT uphill on the same day (55 false notes Jul–Nov 2026, found in the full check 17 Nov).
        // A named uphill that cannot be found gets its own note there.
        if (r.trip.fuelRef?.date) continue;
        const k = ks.find(x => lastUp[x]);
        const up = k && lastUp[k];
        // Any other downhill trip of this driver OR truck in between means the uphill fill belongs to that one —
        // except a column-L trip that names an EARLIER uphill (it is the return leg of that older queue).
        const keysBoth = [...new Set([...ks, ...(up ? pairKeys(up.trip) : [])])];
        const between = up && keysBoth.some(x => (downTrips[x] || []).some(t => t !== r.trip && t.id !== r.trip.id && t.date >= up.trip.date && t.date <= r.trip.date
            && !(t.fuelRef?.date && t.fuelRef.date < up.trip.date)));
        if (!up || between || (Date.parse(r.trip.date) - Date.parse(up.trip.date)) / 86400000 > 5) continue;
        const paisan = Number(r.actual?.[PAISAN]) || 0;
        if (paisan > 0) {
            // A ไพศาล fill on both legs may be the same fill written twice. New arrays, not push: results may be
            // shallow copies of cached ones.
            r.flags = [...r.flags, `อาจบันทึกขาล่องซ้ำ: เที่ยวขาขึ้น ${up.trip.date} บันทึกไพศาลขาล่อง ${up.downActual} ลิตรไว้แล้ว`];
            up.flags = [...up.flags, `อาจบันทึกขาล่องซ้ำ: เที่ยวขาล่อง ${r.trip.date} มีการเติมไพศาล ${paisan} ลิตรด้วย`];
        } else {
            // Fill at another pump (นิสสัน, เอกวิทัศน์) after the downhill ไพศาล was already filled on the uphill trip:
            // probably an extra fill. This pairing is a guess (driver OR truck within 5 days), so it is a NOTE only —
            // the status and the carry-over stay as audited (rule: loose matches are hints, never auto-applied).
            r.flags = [...r.flags, `อาจเป็นการเติมเพิ่มขาล่อง: เที่ยวขาขึ้น ${up.trip.date} บันทึกไพศาลขาล่อง ${up.downActual} ลิตรไว้แล้ว — ตรวจเอง`];
        }
        for (const x of pairKeys(up.trip)) if (lastUp[x] === up) delete lastUp[x]; // one downhill fill pairs with one downhill trip
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
        if (CHECKED.includes(r.status)) {
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
// Extra fill by pickup place (configurable in settings/fuelAudit.originExtras). Default from the company's old fuel
// base (Transport "ฐานคำนวนน้ำมันขาขึ้น": 53 L at เต็กย้ง on สหพัฒน์ / น้ำพริกฉั่วฮะเส็ง pickups, 228 rows) — user
// decision 17 Nov 2026. Real trips: สหพัฒน์ 111 of 122 and น้ำพริกฉั่วฮะเส็ง 10 of 10 fill it; คอลเกต alone and
// น้ำปลาฉั่วฮะเส็ง never do, and "สหพัฒน์ ลพ" never does in the base, so they are not included.
export const DEFAULT_ORIGIN_EXTRAS = [
    { label: 'รับของชลบุรี', keywords: ['สหพัฒน์', 'น้ำพริกฉั่วฮะเส็ง'], exclude: ['สหพัฒน์ ลพ'], pump: 'tekyong_chonburi', maxLiters: 53 },
];
// Clean a rules list coming from settings or the editor: drop rules without keywords / pump / a positive limit.
// ไพศาล is not allowed as a rule pump: its liters are split between uphill and downhill further down.
const words = (v) => (Array.isArray(v) ? v : String(v || '').split(',')).map(x => clean(x).replace(/\s+/g, ' ')).filter(Boolean);
export function normalizeOriginExtras(rules) {
    if (!Array.isArray(rules)) return [];
    return rules.map(r => ({
        label: clean(r?.label),
        keywords: words(r?.keywords),
        exclude: words(r?.exclude),
        pump: clean(r?.pump),
        maxLiters: Math.max(0, Number(r?.maxLiters) || 0),
    })).filter(r => r.keywords.length && r.pump && r.pump !== PAISAN && r.maxLiters > 0);
}
// Rules whose keyword appears in the trip's origin text (and no exclude word does). Spaces are ignored on both
// sides, so "สหพัฒน์ลพ" and "สหพัฒน์  ลพ" are caught by the exclude "สหพัฒน์ ลพ".
export function matchOriginExtras(origin, rules) {
    const squash = (s) => clean(s).replace(/\s+/g, '');
    const o = squash(origin);
    if (!o || !rules?.length) return [];
    return rules.filter(r => r.keywords.some(k => o.includes(squash(k))) && !r.exclude.some(x => o.includes(squash(x))));
}
// Liters allowed on top of the standard at each rule's pump: what was filled there BEYOND the standard for that
// pump, capped at the rule's limit (so a pump already in the route rate is never counted twice).
function originAllowances(trip, actual, ctx, standardUp = {}) {
    if (clean(trip.direction) !== 'ขาขึ้น') return [];
    const out = [];
    for (const rule of matchOriginExtras(trip.origin, ctx.originExtras)) {
        const beyond = Math.max(0, (Number(actual[rule.pump]) || 0) - (Number(standardUp[rule.pump]) || 0));
        const allow = Math.min(beyond, rule.maxLiters);
        if (allow > 0) out.push({ rule, allow });
    }
    return out;
}

export function buildAuditContext(standards, vehicles, opts = {}) {
    const std = {}, drop = {};
    let emptyKpp = null;
    for (const s of standards || []) {
        if (s.cls === DROP_CLASS) { if (s.destination && !s.duplicateOf) drop[s.destination] = s; continue; }
        if (s.cls === EMPTY_KPP_CLASS) { if (!emptyKpp && Number(s.extraLiters) > 0) emptyKpp = s; continue; }
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
    return { std, drop, emptyKpp, vehFull, vehHead, destinations, downDefault, originExtras: normalizeOriginExtras(opts.originExtras),
        downOnUphill: new Set(normalizeDestinationList(opts.downOnUphill === undefined ? DEFAULT_DOWN_ON_UPHILL : opts.downOnUphill)) };
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
// opts.vehicle: judge with this vehicle (entry-form suggestions: the plate or the typed class); opts.expectOnly: return
// the expected liters even when nothing was filled yet (used by suggestFill).
export function auditTrip(trip, ctx, tolerance = 0, opts = {}) {
    const { towOption } = vehicleForPlate(trip.plate, ctx);
    const uphill = clean(trip.direction) === 'ขาขึ้น';
    // Empty trailer left at กพ. (sheet rule, read from the rate data). Without the rule in the data the trip is a
    // manual check — never silently a full towing trip (the dispatcher said the trailer was left at กพ.).
    const emptyKpp = uphill && isEmptyTrailerToKpp(trip);
    const headMode = emptyKpp ? (ctx.emptyKpp ? 'kpp-empty' : 'kpp-empty-norule') : uphill && isDropAtNsw(trip) ? 'nsw' : '';
    if (opts.vehicle) {
        // A drop says the truck was towing: a solo head with a known pair is judged as the pair (same as below).
        const v = headMode && towOption && opts.vehicle !== towOption && vehicleClassOf(opts.vehicle).cls in TOW_OF ? towOption : opts.vehicle;
        return auditTripAs(trip, ctx, tolerance, v, headMode, !!opts.expectOnly);
    }
    // A trailer was dropped, so the truck WAS towing: judge it as its head+trailer pair (keeps e.g. a victor model).
    if (headMode && towOption) {
        const r = auditTripAs(trip, ctx, tolerance, towOption, headMode);
        if (r) r.flags = r.flags.filter(f => !f.startsWith('หัวนี้มีหางพ่วง'));
        return r;
    }
    const first = auditTripAs(trip, ctx, tolerance, null, headMode);
    if (!first || !towOption || first.leg === 'down' || first.status === 'no-rate' || first.status === 'manual' || headMode) return first;
    if (Math.abs(first.diff) <= tolerance) return first;
    const towed = auditTripAs(trip, ctx, tolerance, towOption, '');
    if (towed && towed.status !== 'no-rate' && Math.abs(towed.diff) <= tolerance) {
        towed.flags = [`กระดานคิวเขียนแค่ทะเบียนหัว แต่ลิตรรวมตรงกับอัตรา${AUDIT_CLASSES[towed.cls] || ''}ของ ${towOption.registrationNumber} — ตรวจแบบลากหาง (ควรเขียนเลขหางในกระดาน)`,
            ...towed.flags.filter(f => !f.startsWith('หัวนี้มีหางพ่วง'))];
        return towed;
    }
    return first;
}

// headMode: '' (normal) | 'nsw' (trailer dropped at นครสวรรค์) | 'kpp-empty' (empty trailer towed to กพ. and left
// there — sheet rule "พ่วง ลากหางเปล่า ไปตัดไว้ที่ กำแพง น้ำมันจ่ายเป็นรถสิบล้อ + เพิ่ม N ลิตร", user confirmed 17 Nov 2026)
// | 'kpp-empty-norule' (same, but the rule is not in the imported rate data yet → manual check).
// In both drop modes the uphill is charged at the head's own rate and the downhill at the towing rate.
function auditTripAs(trip, ctx, tolerance, forcedVehicle, headMode = '', expectOnly = false) {
    const dir = clean(trip.direction);
    if (dir !== 'ขาขึ้น' && dir !== 'ขาล่อง') return null;
    const actual = { ...(trip.gasStationsLiters || {}) };
    const baht = { ...(trip.gasStationsBaht || {}) };
    const actualTotal = Object.values(actual).reduce((a, b) => a + (Number(b) || 0), 0);
    const bahtTotal = Object.values(baht).reduce((a, b) => a + (Number(b) || 0), 0);
    if (!actualTotal && !bahtTotal && !expectOnly) return null;
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
    const manual = (reason) => ({ ...base, status: 'manual', reason, diff: 0, expected: {}, expectedTotal: 0, downStatus: '' });
    const EMPTY_LABEL = 'ลากหางเปล่าไปตัดไว้ที่กำแพงเพชร';
    if (headMode === 'kpp-empty-norule') return manual(`${EMPTY_LABEL} — ตารางอัตราในระบบยังไม่มีกฎ "เพิ่ม … ลิตร" (นำเข้าไฟล์อัตราใหม่) ตรวจเอง`);
    // The sheet writes the empty-trailer rule for พ่วง only ("จ่ายเป็นรถสิบล้อ"); other classes are not guessed.
    let mode = headMode;
    if (mode === 'kpp-empty' && vc.cls !== '18W') return manual(`${EMPTY_LABEL} — กฎในตารางอัตราเขียนไว้สำหรับรถพ่วงเท่านั้น (รถนี้เป็น${AUDIT_CLASSES[vc.cls] || 'ประเภทอื่น'}) ตรวจเอง`);
    // Trailer left at Kamphaeng Phet: the special table has total liters only (no pump) → manual check.
    // User rule (17 Nov 2026): "ถ้าเป็นรถพ่วง ไปสาขา กพ. จ่าย 80 ลิตร" — when the trip itself ends at กำแพงเพชร
    // (head delivers to a shop there, trailer left at the branch), it is an ordinary พ่วง trip to กพ.: the normal
    // rate row applies, so fall through. The special table is for heads that go on beyond กพ.
    const kppEndsHere = destination === 'กำแพงเพชร';
    if (isDropAtKpp(trip) && !kppEndsHere && mode !== 'kpp-empty') {
        const d = ctx.drop?.[destination];
        // The drop table has no downhill figure, but the ไพศาล downhill fill is often written on this trip —
        // name it so a normal downhill fill is not read as an overfill.
        const paisan = Number(actual[PAISAN]) || 0;
        const paisanNote = paisan ? ` · ในยอดนี้มีปั๊มไพศาล ${paisan} ลิตร ซึ่งอาจเป็นการเติมขาล่อง` : '';
        // The drop table has no pump split, so the whole fill at a rule's pump (up to the limit) is allowed on top.
        const extras = originAllowances(trip, actual, ctx);
        const extra = extras.reduce((a, x) => a + x.allow, 0);
        const extraNote = extras.map(x => ` · ${x.rule.label || 'เติมเพิ่มตามต้นทาง'}: ${pumpThai(x.rule.pump)} ได้เพิ่ม ${x.allow} ลิตร`).join('');
        const allowed = d ? d.upTotal + extra : 0;
        const reason = d
            ? `เที่ยวถอดหางที่กำแพงเพชร — อัตราพิเศษ ${d.parts.map(p => `${p.part}→${p.destination} ${p.liters}`).join(' + ')} = ${d.upTotal} ลิตร (ไม่ระบุปั๊ม ไม่รวมขาล่อง)${extraNote} · เติมจริงรวม ${actualTotal} ลิตร (ต่าง ${actualTotal - allowed > 0 ? '+' : ''}${actualTotal - allowed})${paisanNote}`
            : `เที่ยวถอดหางที่กำแพงเพชร — ตารางถอดหางไม่มีปลายทาง "${clean(trip.destination) || '-'}" ตรวจกับหัวหน้างาน${extraNote}${paisanNote}`;
        return { ...base, status: 'manual', reason, diff: d ? actualTotal - allowed : 0, expected: {}, expectedTotal: allowed, downStatus: '' };
    }
    if (!vc.cls) return none(vc.reason);
    if (!origin) return none(`ต้นทาง "${clean(trip.origin)}" เป็นสาขาอื่น (ตารางมีแค่ กรุงเทพ / นครสวรรค์)`);
    if (!destination) return none(`ปลายทาง "${clean(trip.destination) || '-'}" ไม่อยู่ในตารางอัตรา`);
    // Trailer dropped (at นครสวรรค์, or an empty trailer at กพ.) (user rules, 17 Nov 2026): the head runs on alone at
    // its own rate — a พ่วง's head is a 10-wheeler, a 6-wheel box trailer's head a 6-wheeler — and comes back towing
    // the trailer, so the downhill fill follows the towing rate (seen on real trips: จิตตมาศ 50 + ไพศาล 180 to กพ.).
    // NSW: not when the trip ends at นครสวรรค์ itself (the trailer was towed the whole way).
    const pick = (cls, variant) => ctx.std[`${cls}|${origin}|${destination}|${variant}`];
    const pickV = (cls) => (vc.model === 'sino' && pick(cls, 'sino')) || pick(cls, 'standard');
    const viaRows = (cls) => Object.keys(VIA).map(v => (vc.model === 'sino' && pick(cls, `${v}_sino`)) || pick(cls, v)).filter(Boolean);
    const dropLabel = mode === 'kpp-empty' ? EMPTY_LABEL : 'ตัดหางที่นครสวรรค์';
    let headDrop = !!mode && (mode !== 'nsw' || destination !== 'นครสวรรค์') && (vc.cls in HEAD_OF || vc.cls in TOW_OF);
    let upCls = headDrop ? (HEAD_OF[vc.cls] || vc.cls) : vc.cls;
    const downCls = headDrop ? (TOW_OF[upCls] || vc.cls) : vc.cls;
    // Only "ตีเข้า กทม / ตีเข้า นครสวรรค์" rates for this place: the trip does not say which applies → manual check
    // with both figures (never one picked for the person). Checked for the head class, the towing class after a
    // fallback, and the towing class used for the downhill.
    const viaManual = (cls) => {
        if (pickV(cls) || !viaRows(cls).length) return null;
        const lbl = (x) => `${VIA[viaBase(x.variant)]} ${x.upTotal} ลิตร (${Object.entries(x.up || {}).map(([k, v]) => `${pumpThai(k)} ${v}`).join(' + ') || '-'})`;
        return manual(`ตารางแยกอัตรา${AUDIT_CLASSES[cls] || ''} ${origin} → ${destination} ตามการตีรถเข้า: ${viaRows(cls).map(lbl).join(' / ')} — เที่ยวไม่ได้ระบุว่าตีเข้าที่ไหน ตรวจเอง · เติมจริงรวม ${actualTotal} ลิตร`);
    };
    const viaUp = viaManual(upCls);
    if (viaUp) return viaUp;
    if (headDrop && !pickV(upCls)) {
        // Empty trailer: the rule is "สิบล้อ + N" — without a 10-wheel rate there is nothing to apply → manual.
        if (mode === 'kpp-empty') return manual(`${EMPTY_LABEL} แต่ตารางไม่มีอัตรา${AUDIT_CLASSES[upCls]}ไป ${destination} — ตรวจเอง`);
        // NSW drop: no head-alone rate for this route → fall back to the towing rate, and say so.
        base.flags.push(`${dropLabel} แต่ตารางไม่มีอัตรา${AUDIT_CLASSES[upCls]}ไป ${destination} — ใช้อัตรา${AUDIT_CLASSES[downCls]}แทน`);
        headDrop = false; upCls = downCls; base.rateCls = downCls;
    }
    const viaAny = viaManual(upCls) || (upCls !== downCls && viaManual(downCls));
    if (viaAny) return viaAny;
    const s = pickV(upCls);
    if (!s) return none(`ไม่มีอัตรา ${AUDIT_CLASSES[upCls]} ${origin} → ${destination}`);
    // Downhill comes from the towing class's row for THIS route only — a class-wide default may not match
    // the route, so when the row is missing the downhill fill is left unchecked (and the note says so).
    const sDown = upCls === downCls ? s : (pickV(downCls) || { down: 0, downVictor: 0 });
    if (headDrop) {
        base.rateCls = upCls;
        base.flags.push(sDown.down
            ? `${dropLabel} — ขาขึ้นคิดอัตรา${AUDIT_CLASSES[upCls]} (ตัวแม่) ขาล่องคิดอัตรา${AUDIT_CLASSES[downCls]} (ลากหางกลับ)`
            : `${dropLabel} — ขาขึ้นคิดอัตรา${AUDIT_CLASSES[upCls]} (ตัวแม่) · ไม่มีอัตราขาล่องของ${AUDIT_CLASSES[downCls]} จึงไม่ได้ตรวจขาล่อง`);
    }
    const emptyExtra = headDrop && mode === 'kpp-empty' ? Number(ctx.emptyKpp?.extraLiters) || 0 : 0;

    const down = vc.model === 'victor' ? sDown.downVictor : sDown.down;
    const expected = { ...s.up };
    // Extra fill allowed by pickup place (settings, e.g. "รับของชลบุรี: เต็กย้ง ได้ไม่เกิน 53 ลิตร" from the company's
    // old fuel base — the official sheet has no ชลบุรี origin). An ALLOWANCE, not a requirement: what was filled at
    // that pump counts as expected up to the limit, so not filling it is not "under" and an overfill still shows.
    for (const { rule, allow } of originAllowances(trip, actual, ctx, s.up)) {
        expected[rule.pump] = (expected[rule.pump] || 0) + allow;
        base.flags.push(`${rule.label || 'เติมเพิ่มตามต้นทาง'}: ${pumpThai(rule.pump)} ได้เพิ่มไม่เกิน ${rule.maxLiters} ลิตร (นับ ${allow})`);
    }
    // Empty trailer left at กพ.: "+N ลิตร" on top of the head's rate, pump not named in the sheet → an allowance of up
    // to N liters over the standard, first at the other uphill pumps, then at ไพศาล (after its downhill share).
    let emptyLeft = emptyExtra; const emptyUsed = [];
    const takeEmpty = (k) => {
        const beyond = Math.max(0, (Number(actual[k]) || 0) - (expected[k] || 0));
        const take = Math.min(beyond, emptyLeft);
        if (take > 0) { expected[k] = (expected[k] || 0) + take; emptyLeft -= take; emptyUsed.push(`${pumpThai(k)} ${take}`); }
        return take;
    };
    if (emptyExtra) for (const k of Object.keys(actual)) if (k !== PAISAN && emptyLeft > 0) takeEmpty(k);
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
    const paisanExtra = emptyExtra && emptyLeft > 0 ? takeEmpty(PAISAN) : 0;
    if (emptyExtra) base.flags.push(`ลากหางเปล่า: ได้เพิ่มไม่เกิน ${emptyExtra} ลิตร ตามตารางอัตรา (${ctx.emptyKpp.sheet} แถว ${ctx.emptyKpp.row})${emptyUsed.length ? ` — นับ ${emptyUsed.join(', ')}` : ''}`);
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
    // Downhill (ขาล่อง) check: the ไพศาล fill written on the uphill trip, minus the uphill share and any empty-trailer
    // extra counted at ไพศาล.
    const downActual = downhillRecorded ? actPaisan - upPaisan - paisanExtra : 0;
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

// ---- Driver name aliases (board short names → full names), confirmed by the user (17 Nov 2026) ----
// Stored in settings/fuelAudit.driverAliases as [{ from, to }] (false = none); used only when auditing — the stored
// trips keep their board names.
export function normalizeDriverAliases(list) {
    const map = {};
    for (const a of Object.values(list || {})) {
        const from = clean(a?.from), to = clean(a?.to);
        if (from && to && from !== to) map[from] = to;
    }
    return map;
}
// The trip as the audit sees it: board name replaced by the confirmed full name (board name kept in driverBoard).
export function applyDriverAlias(trip, map) {
    const key = clean(trip?.driver);
    // Own keys only: a driver named like an Object property ("constructor", "toString") must not match.
    const full = map && Object.prototype.hasOwnProperty.call(map, key) ? map[key] : '';
    return full ? { ...trip, driver: full, driverBoard: trip.driver } : trip;
}

/**
 * The confirmed proposal file (sheet "จับคู่ชื่อ": ชื่อในกระดานคิว | … | ชื่อเต็มที่เสนอ | … | ยืนยัน) → aliases.
 * "ยืนยัน" = ใช่ / ถูก / yes / ✓ → the proposed full name; ไม่ / ไม่ใช่ / ผิด / no / ✗ or blank → not used; any other
 * text = the full name the user typed instead. Nothing ambiguous is saved: a full name missing from the drivers
 * master, a board name confirmed to two different people, a board name that is itself another driver's full name,
 * or a chain (a target that is also a board name) goes to `conflicts` for the user to fix in the file.
 */
export function parseDriverAliasSheet(rows, masterNames = []) {
    const hi = (rows || []).slice(0, 5).findIndex(r => { const h = (r || []).map(clean); return h.includes('ชื่อในกระดานคิว') && h.includes('ยืนยัน'); });
    const H = hi >= 0 ? rows[hi].map(clean) : [];
    const iS = H.indexOf('ชื่อในกระดานคิว'), iP = H.indexOf('ชื่อเต็มที่เสนอ'), iY = H.indexOf('ยืนยัน');
    const empty = { aliases: [], skipped: [], conflicts: [] };
    if (iS < 0 || iP < 0 || iY < 0) return { ...empty, error: 'ไม่ใช่ไฟล์จับคู่ชื่อคนขับ (ต้องมีช่อง ชื่อในกระดานคิว / ชื่อเต็มที่เสนอ / ยืนยัน)' };
    const master = new Set(masterNames.map(clean).filter(Boolean));
    if (!master.size) return { ...empty, error: 'ยังไม่มีรายชื่อคนขับในระบบ (หน้า "คนขับ") — ตรวจชื่อเต็มไม่ได้ ลองใหม่อีกครั้ง' };
    const YES = /^(ใช่|ถูก|ถูกต้อง|ยืนยัน|yes|y|ok|✓|✔|1|true)$/i;
    const NO = /^(ไม่|ไม่ใช่|ผิด|ไม่ถูก|ไม่ยืนยัน|no|n|x|✗|✘|0|false|-)$/i;
    const picked = [], skipped = [], conflicts = [];
    for (const r of rows.slice(hi + 1)) {
        const from = clean(r?.[iS]);
        if (!from) continue;
        const said = clean(r?.[iY]);
        if (!said || NO.test(said)) { skipped.push(from); continue; }
        const to = YES.test(said) ? clean(r?.[iP]) : said;
        if (!to) { skipped.push(from); continue; }
        if (!master.has(to)) { conflicts.push({ from, to, reason: 'ชื่อเต็มไม่อยู่ในหน้า "คนขับ"' }); continue; }
        if (from !== to) picked.push({ from, to });
    }
    // Ambiguity checks over the whole file.
    const targets = {}; for (const a of picked) (targets[a.from] ||= new Set()).add(a.to);
    const froms = new Set(picked.map(a => a.from));
    const aliases = [], seen = new Set();
    for (const a of picked) {
        if (targets[a.from].size > 1) { conflicts.push({ ...a, reason: `ชื่อในกระดานนี้ถูกยืนยันเป็นหลายคน (${[...targets[a.from]].join(' / ')})` }); continue; }
        if (master.has(a.from)) { conflicts.push({ ...a, reason: 'ชื่อในกระดานตรงกับชื่อเต็มของคนขับอีกคน' }); continue; }
        if (froms.has(a.to)) { conflicts.push({ ...a, reason: 'ชื่อเต็มนี้ถูกจับคู่ต่ออีกทอด' }); continue; }
        if (seen.has(a.from)) continue; // same pair listed twice
        seen.add(a.from);
        aliases.push(a);
    }
    return { aliases, skipped, conflicts, error: '' };
}

// ---- What the entry forms should suggest (user, 17 Nov 2026: "ผมว่าคุณผิดแล้ว" — the form card still read the old
// route table, e.g. 18W โกดังสาย 3 → พิษณุโลก จิตตมาศ 180, while the official sheet says จิตตมาศ 80 + ไพศาล ขาล่อง 180).
// Built on auditTrip itself (expectOnly), so the form suggests exactly what the audit will expect — trailer drops,
// head+trailer pairs, victor/sino, origin extras all included — and the two can never drift apart.

// Destinations whose downhill ไพศาล fill the board writes on the UPHILL record (settings/fuelAudit.downOnUphill).
// Default from real trips Jul–Nov 2026: 99–100% for these; 0% for เชียงใหม่ แม่สอด ลำพูน แม่สาย เชียงราย แพร่.
export const DEFAULT_DOWN_ON_UPHILL = ['นครสวรรค์', 'พิษณุโลก', 'กำแพงเพชร', 'เพชรบูรณ์', 'ตาก', 'บึงสามพัน'];
export function normalizeDestinationList(list) {
    if (list === false || list == null) return [];
    const arr = Array.isArray(list) ? list : typeof list === 'object' ? Object.values(list) : String(list).split(',');
    return [...new Set(arr.map(x => canonDestination(x) || clean(x)).filter(Boolean))];
}

/**
 * q: { vehicleType, plate, origin, destination, direction, jobDetail, notes, queueType, trailerDrop }.
 * Returns { ok:true, cls, rateCls, model, origin, destination, up, down, downOnUp, fillUp, fillWithDown, fill, extras,
 *           notes, sheet, row } or { ok:false, reason }. fill = the default (with the downhill when downOnUp).
 */
export function suggestFill(q, ctx) {
    const dir = clean(q.direction) || 'ขาขึ้น';
    // Vehicle: the plate's master row; a head typed with a towing class uses its pair; unknown plate → typed class.
    let vehicle = null; const notes = [];
    if (clean(q.plate)) {
        const f = vehicleForPlate(q.plate, ctx);
        vehicle = f.vehicle;
        const typed = clean(q.vehicleType) ? vehicleClassOf({ vehicleType: q.vehicleType, driveType: 'รถบริษัท' }).cls : null;
        if (vehicle && typed && f.towOption && vehicleClassOf(vehicle).cls !== typed && vehicleClassOf(f.towOption).cls === typed) vehicle = f.towOption;
        if (!vehicle) notes.push('ทะเบียนนี้ไม่อยู่ในหน้าทะเบียนรถ — หน้าตรวจการเติมจะตรวจเที่ยวนี้ไม่ได้ (เพิ่มทะเบียนก่อน)');
    }
    if (!vehicle) vehicle = { vehicleType: q.vehicleType, driveType: 'รถบริษัท' };
    const vc = vehicleClassOf(vehicle);
    if (vc.partner) return { ok: false, reason: 'รถร่วม — ไม่มีอัตราในตาราง' };
    if (!vc.cls) return { ok: false, reason: vc.reason };
    const model = vc.model || '';
    if (dir === 'ขาล่อง') {
        const def = ctx.downDefault?.[vc.cls];
        if (!def) return { ok: false, reason: `ไม่มีอัตราขาล่องของ ${AUDIT_CLASSES[vc.cls]}` };
        const down = model === 'victor' ? def.downVictor : def.down;
        notes.push('ถ้าเที่ยวขาขึ้นของรอบนี้บันทึกไพศาลขาล่องไว้แล้ว (เส้นเหนือล่าง เช่น นว./พิษณุโลก/กพ.) ไม่ต้องเติมซ้ำในเที่ยวขาล่อง');
        return { ok: true, cls: vc.cls, rateCls: vc.cls, model, origin: clean(q.origin), destination: clean(q.destination), up: {}, down, downOnUp: false,
            fillUp: {}, fillWithDown: { [PAISAN]: down }, fill: {}, extras: [], notes, sheet: '', row: '', downOnly: true };
    }
    if (!clean(q.origin) || !clean(q.destination)) return { ok: false, reason: 'กรอกต้นทางและปลายทางให้ครบ' };
    const trip = { direction: 'ขาขึ้น', plate: clean(q.plate), origin: q.origin, destination: q.destination, jobDetail: q.jobDetail || '',
        notes: q.notes || '', queueType: q.queueType || '', trailerDrop: q.trailerDrop || '', gasStationsLiters: {}, gasStationsBaht: {} };
    const r = auditTrip(trip, ctx, 0, { vehicle, expectOnly: true });
    if (!r) return { ok: false, reason: 'คำนวณไม่ได้' };
    if (r.status === 'no-rate' || r.status === 'manual') return { ok: false, reason: r.reason, notes };
    // With nothing filled the audit's ไพศาล reading leaves the downhill out of `expected` — exactly the uphill part.
    const up = Object.fromEntries(Object.entries(r.expected || {}).filter(([, v]) => v > 0));
    const down = r.down || 0;
    const fillWithDown = { ...up };
    if (down) fillWithDown[PAISAN] = (fillWithDown[PAISAN] || 0) + down;
    const downOnUp = !!down && ctx.downOnUphill?.has(r.destination);
    for (const f of r.flags) if (!f.startsWith('ขาล่องไพศาล') && !f.startsWith('อัตราแถวนี้มีข้อสังเกต')) notes.push(f);
    return { ok: true, cls: r.cls, rateCls: r.rateCls || r.cls, model: r.model, origin: r.origin, destination: r.destination, up, down, downOnUp,
        fillUp: up, fillWithDown, fill: downOnUp ? fillWithDown : up, extras: matchOriginExtras(q.origin, ctx.originExtras),
        notes, rateFlags: r.standard?.flags || [], sheet: r.standard?.sheet || '', row: r.standard?.row || '' };
}
