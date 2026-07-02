// Excel import parsing — pure functions, no DOM/Firebase so they can be unit-tested in Node.
// Parses the "ฐานคำนวนน้ำมันขาขึ้น" (uphill fuel-base) sheet into route objects for the app.
import { parseNum, clean } from './parseUtils.js';

// Map each fuel-base station COLUMN (by exact header text) to an app station key + unit.
// Both Liters and Baht for a station are stored independently on a route.
export const FUEL_BASE_STATION_MAP = [
    { header: 'ปั๊มจิตตมาศ',                    key: 'jittamas_bangpain',     unit: 'Liters' },
    { header: 'ปั๊มจิตตมาศ (บาท)',              key: 'jittamas_bangpain',     unit: 'Baht' },
    { header: 'ปั๊มศรีนครคลองขลุง',             key: 'srinakorn_klongklung',  unit: 'Liters' },
    { header: 'ปั๊มไพศาลน.ว',                   key: 'daopaisan_nakhonsawan', unit: 'Liters' },
    { header: 'ปั๊มสยามนิสสัน สิงห์บุรี',        key: 'nissan_singburi',       unit: 'Liters' },
    { header: 'ปั๊มคาลเท็กซ์ ดาวอุไร (บาท)',     key: 'daourai_sai3',          unit: 'Baht' },
    { header: 'ปั๊มคาลเท็กซ์ เครือสหพัฒน์',      key: 'caltex_sahapat',        unit: 'Liters' },
    { header: 'ปั๊มคาลเท็กซ์ เครือสหพัฒน์ (บาท)', key: 'caltex_sahapat',        unit: 'Baht' },
    { header: 'ปั๊มไพศาลน.ว(ขาล่อง)',           key: 'daopaisan_nakhonsawan', unit: 'Liters' },
];

const dupeKey = (r) => [r.vehicleType, r.origin, r.destination, r.direction].join('||').toLowerCase().replace(/\s+/g, ' ');

// rows: 2D array from XLSX.utils.sheet_to_json(sheet, { header: 1 }).
// Returns { routes, stats, warnings }.
export function parseFuelBaseRows(rows) {
    if (!rows || rows.length < 2) {
        return { routes: [], stats: { total: 0, withFuel: 0, zeroFuel: 0, duplicates: 0, skipped: 0 }, warnings: ['ไม่พบข้อมูลในชีต'] };
    }
    const header = rows[0].map(clean);
    const idx = (name) => header.indexOf(name);

    const iType = idx('trucktype');
    const iOrigin = idx('ต้นทาง');
    const iDest = idx('ปลายทาง');
    const iCross = idx('ค่าข้าม');
    const iWait = idx('ค่าค้าง');

    const warnings = [];
    if (iType < 0 || iOrigin < 0 || iDest < 0) {
        return { routes: [], stats: { total: 0, withFuel: 0, zeroFuel: 0, duplicates: 0, skipped: 0 },
            warnings: ['ไม่พบคอลัมน์หลัก (trucktype / ต้นทาง / ปลายทาง) — อาจเลือกชีตผิด'] };
    }
    // Resolve station columns present in this sheet.
    const stationCols = FUEL_BASE_STATION_MAP
        .map(m => ({ ...m, col: idx(m.header) }))
        .filter(m => m.col >= 0);

    // A route's "fuel signature" — used to pick the most common (mode) fuel per route.
    const sigOf = (rt) => {
        const ser = (obj) => Object.keys(obj).sort().map(k => `${k}:${obj[k]}`).join(',');
        return `L{${ser(rt.gasStationsLiters)}}B{${ser(rt.gasStationsBaht)}}`;
    };

    // Pass 1: build a candidate route per data row, grouped by (type|origin|dest|direction).
    const groups = new Map();
    let skipped = 0, rowsUsed = 0;
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.every(c => c === '' || c == null)) continue;
        const vehicleType = clean(row[iType]);
        const origin = clean(row[iOrigin]);
        const destination = clean(row[iDest]);
        if (!vehicleType || !origin || !destination) { skipped++; continue; }

        const gasStationsLiters = {};
        const gasStationsBaht = {};
        for (const s of stationCols) {
            const val = parseNum(row[s.col]);
            if (val > 0) {
                const bucket = s.unit === 'Baht' ? gasStationsBaht : gasStationsLiters;
                bucket[s.key] = (bucket[s.key] || 0) + val;
            }
        }
        const route = {
            vehicleType, origin, destination,
            direction: 'ขาขึ้น',
            crossingFee: iCross >= 0 ? clean(row[iCross]) : '',
            waitingFee: iWait >= 0 ? clean(row[iWait]) : '',
            gasStationsLiters, gasStationsBaht,
            totalLiters: Object.values(gasStationsLiters).reduce((a, b) => a + b, 0),
            totalBaht: Object.values(gasStationsBaht).reduce((a, b) => a + b, 0),
        };
        rowsUsed++;
        const k = dupeKey(route);
        if (!groups.has(k)) groups.set(k, []);
        groups.get(k).push(route);
    }

    // Pass 2: for each route, choose the most common (mode) fuel signature.
    // Tie-break: higher total (more complete), then first seen.
    const routes = [];
    let withFuel = 0, zeroFuel = 0, variedRoutes = 0;
    for (const cands of groups.values()) {
        const counts = new Map(); // sig -> { count, total, route }
        for (const c of cands) {
            const s = sigOf(c);
            if (!counts.has(s)) counts.set(s, { count: 0, total: c.totalLiters + c.totalBaht, route: c });
            counts.get(s).count++;
        }
        if (counts.size > 1) variedRoutes++;
        let best = null;
        for (const v of counts.values()) {
            if (!best || v.count > best.count || (v.count === best.count && v.total > best.total)) best = v;
        }
        const route = best.route;
        routes.push(route);
        if (route.totalLiters > 0 || route.totalBaht > 0) withFuel++; else zeroFuel++;
    }

    return {
        routes,
        stats: { total: routes.length, withFuel, zeroFuel, variedRoutes, rowsUsed, skipped, stationColumns: stationCols.length },
        warnings,
    };
}
