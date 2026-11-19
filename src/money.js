// Money rounding (user decision 19 Nov 2026: round only what is SHOWN / TOTALLED, 2 decimals, half up; the amounts
// stored on each trip are not changed). roundHalfUp ported from the subcontractor-truck project (utils/format.ts):
// shifting with exponent notation instead of multiplying keeps 10.075 → 10.08 (× 100 gives 1007.4999… → 10.07).

export function roundHalfUp(num, decimals = 2) {
    if (!Number.isFinite(num)) return 0;
    const sign = num < 0 ? -1 : 1;
    const abs = Math.abs(num);
    // below half a unit of the last decimal it is 0 — also covers float noise such as 5.55e-17, whose exponent form
    // would break the shift below — and never -0 (a "-0" in a deduction column)
    if (abs < 0.5 * 10 ** -decimals) return 0;
    if (abs >= Number.MAX_SAFE_INTEGER) return num;
    const shifted = Number(`${abs}e${decimals}`);
    if (!Number.isFinite(shifted)) return num;
    const rounded = Math.round(shifted);
    if (!Number.isSafeInteger(rounded)) return num;
    return sign * Number(`${rounded}e-${decimals}`) + 0;
}

export const round2 = (n) => roundHalfUp(Number(n) || 0, 2);

/** Sum money with the total rounded to 2 decimals (no 0.30000000000000004 in totals or exports). */
export const sumMoney = (list, get = (x) => x) => round2((list || []).reduce((s, x) => s + (Number(get(x)) || 0), 0));

/** The same object with the named money fields rounded to 2 decimals (aggregated rows). */
export const roundFields = (obj, fields) => { const o = { ...obj }; for (const f of fields) if (f in o) o[f] = round2(o[f]); return o; };
