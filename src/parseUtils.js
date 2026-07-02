// Shared pure helpers for the Excel importers (no DOM/Firebase → unit-testable in Node).

// Parse a cell into a finite number; blanks / non-numeric / commas → 0.
export const parseNum = (v) => {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    const s = String(v).replace(/,/g, '').trim();
    const n = parseFloat(s);
    return isFinite(n) ? n : 0;
};

// Coerce a cell to a trimmed string (null/undefined → '').
export const clean = (v) => String(v == null ? '' : v).trim();
