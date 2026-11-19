// Closed pay periods (item 6, user 19 Nov 2026): rules shared by services.js and app.js. Pure functions.
// The lock is app-level only (database rules = phase 3, parked).

/** Months of `yms` that are closed (each once, in the given order). */
export const closedOf = (yms, closedMonths) => [...new Set(yms || [])].filter(ym => ym && closedMonths?.[ym]);

/** Settings parts that a restore must never take from the file: the CURRENT lock state and its history win. */
export const KEEP_ON_SETTINGS_RESTORE = ['closedMonths', 'periodLog', 'restoreLog'];

/** The settings value to write when restoring settings from a backup: the file's settings + the current lock/logs. */
export function settingsForRestore(fromFile, current) {
    const out = { ...(fromFile || {}) };
    for (const k of KEEP_ON_SETTINGS_RESTORE) {
        if (current && current[k] != null) out[k] = current[k];
        else delete out[k];
    }
    return out;
}

/** Closed months a new rate set would re-judge: every closed month from its start month on ('' = every month). */
export const rateSaveClosedHit = (effectiveFrom, closedMonths) =>
    Object.keys(closedMonths || {}).filter(ym => !effectiveFrom || ym >= String(effectiveFrom).slice(0, 7)).sort();

/** Parts whose restore would rewrite a closed month (trips, column-L liters, rate sets used by its audit). */
export const PARTS_LOCKED_BY_CLOSED_MONTHS = ['trips', 'downhillFuel', 'fuelRateStandards'];
