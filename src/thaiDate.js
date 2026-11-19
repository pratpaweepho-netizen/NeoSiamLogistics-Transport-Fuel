// Thai date display and typing (user decision 19 Nov 2026): every date the user SEES or TYPES is dd/mm/yyyy with the
// Buddhist year (19/11/2569); months are mm/yyyy (11/2569). Data stays ISO — YYYY-MM-DD dates, YYYY-MM month keys —
// so sorting, month buckets and the audit are untouched. Pure functions (no DOM).

const pad = (n) => String(n).padStart(2, '0');
const MIN_Y = 2000, MAX_Y = 2100; // a typed/stored year outside this range is a typo, never a date
const validYmd = (y, m, d) => {
    if (!(y >= MIN_Y && y <= MAX_Y) || !(m >= 1 && m <= 12) || !(d >= 1)) return false;
    return d <= new Date(Date.UTC(y, m, 0)).getUTCDate();
};

/**
 * '2026-11-19' → '19/11/2569'. An ISO date-time (stored timestamp) → its LOCAL calendar day. Anything else (blank,
 * text, an impossible date) → returned unchanged, so a bad stored value stays visible.
 */
export function fmtDate(iso) {
    const s = String(iso ?? '');
    let m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if (m) return validYmd(+m[1], +m[2], +m[3]) ? `${m[3]}/${m[2]}/${+m[1] + 543}` : s;
    if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
        const t = new Date(s);
        if (!isNaN(t)) return `${pad(t.getDate())}/${pad(t.getMonth() + 1)}/${t.getFullYear() + 543}`;
    }
    return s;
}

/** '2026-11' → '11/2569'. Anything else → unchanged. */
export function fmtMonth(ym) {
    const s = String(ym ?? '');
    const m = /^(\d{4})-(\d{2})$/.exec(s);
    return m && validYmd(+m[1], +m[2], 1) ? `${m[2]}/${+m[1] + 543}` : s;
}

/**
 * ISO dates (YYYY-MM-DD) and month keys (YYYY-MM) inside a free text — audit notes, import errors — shown the Thai
 * way. Values glued to letters/underscores/dashes (ids such as tdh_2026-07-02_…) are left alone.
 */
export function fmtDateText(text) {
    return String(text ?? '')
        .replace(/(^|[^\w-])(\d{4})-(\d{2})-(\d{2})(?![\w-])/g, (all, pre, y, m, d) => (validYmd(+y, +m, +d) ? `${pre}${d}/${m}/${+y + 543}` : all))
        .replace(/(^|[^\w-])(\d{4})-(\d{2})(?![\w-])/g, (all, pre, y, m) => (validYmd(+y, +m, 1) ? `${pre}${m}/${+y + 543}` : all));
}

/** Excel number format for a date with the Buddhist year (Excel and SheetJS both show 19/11/2569). */
export const XL_DATE_FMT = 'dd/mm/bbbb';

/**
 * '2026-11-19' → a SheetJS cell holding a REAL Excel date (sorts / filters as a date) shown as 19/11/2569.
 * Built from the calendar date itself (no time zone involved). Not a valid ISO date → returned unchanged.
 */
export function xlDate(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso ?? ''));
    if (!m || !validYmd(+m[1], +m[2], +m[3])) return iso ?? '';
    return { t: 'n', v: (Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000, z: XL_DATE_FMT };
}

// Year as typed (same rule as the Excel trip template): 4 digits ≥ 2400 = พ.ศ., other 4 digits = ค.ศ.;
// 2 digits ≥ 40 = พ.ศ. short ("69" → 2569 → 2026), 2 digits < 40 = ค.ศ. short ("26" → 2026).
const yearOf = (t) => {
    if (/^\d{2}$/.test(t)) return +t >= 40 ? 2500 + +t - 543 : 2000 + +t;
    if (!/^\d{4}$/.test(t)) return NaN;
    const y = +t;
    return y >= 2400 ? y - 543 : y;
};

/**
 * Typed '19/11/2569', '19-11-2569', '19.11.69', '1/2/2026', or without separators '19112569' / '191169' (phone
 * number pads have no "/") → '2026-11-19'; '' when it is not a real date in 2000–2100.
 */
export function parseThaiDate(text) {
    const s = String(text ?? '').trim();
    const m = /^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2}|\d{4})$/.exec(s) || /^(\d{2})(\d{2})(\d{4}|\d{2})$/.exec(s);
    if (!m) return '';
    const d = +m[1], mo = +m[2], y = yearOf(m[3]);
    return validYmd(y, mo, d) ? `${y}-${pad(mo)}-${pad(d)}` : '';
}

/** Typed '11/2569', '11-69', '1/2026', or '112569' → '2026-11'; '' when not a real month in 2000–2100. */
export function parseThaiMonth(text) {
    const s = String(text ?? '').trim();
    const m = /^(\d{1,2})[/.\-](\d{2}|\d{4})$/.exec(s) || /^(\d{2})(\d{4})$/.exec(s);
    if (!m) return '';
    const mo = +m[1], y = yearOf(m[2]);
    return validYmd(y, mo, 1) ? `${y}-${pad(mo)}` : '';
}
