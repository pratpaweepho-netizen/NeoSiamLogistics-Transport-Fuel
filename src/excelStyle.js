// NeoSiam standard Excel look for the report exports (phase 4 item 10, user-approved 19 Nov 2026).
// Spec ported from the KPI project (D:\logistics-delivery-kpi-management\services\excelStyle.ts): Cordia New,
// navy 1B365D header with white text, zebra rows, gold total row with a double bottom border.
// Only the LOOK changes — the header stays on row 1 and the cells are the same values as before, so anyone who
// reads these files back (filters, sums, another import) sees the same columns. Pure: ExcelJS is passed in.

export const NEO = {
    navy: 'FF1B365D', white: 'FFFFFFFF', gold: 'FF7D6608', zebra: 'FFF9FAFC', fillTotal: 'FFFCF3CF', gridline: 'FFCCCCCC',
};
const FONT = 'Cordia New';
export const FONTS = {
    head: { name: FONT, size: 14, bold: true, color: { argb: NEO.white } },
    body: { name: FONT, size: 13 },
    total: { name: FONT, size: 14, bold: true, color: { argb: NEO.gold } },
};
const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const thin = (argb) => ({ style: 'thin', color: { argb } });
const BORDER = { left: thin(NEO.gridline), right: thin(NEO.gridline), top: thin(NEO.gridline), bottom: thin(NEO.gridline) };
const BORDER_TOTAL = { ...BORDER, top: thin(NEO.gold), bottom: { style: 'double', color: { argb: NEO.gold } } };

export const MONEY_FMT = '#,##0.00';
export const COUNT_FMT = '#,##0';
// Which columns are baht / counts, by header text (one place for every export).
const MONEY_HEAD = /บาท|รายรับ|เบี้ยขับ|เบี้ยเลี้ยง|หักเงิน|สุทธิ/;
const COUNT_HEAD = /^(จำนวนคิว|คิว|เที่ยว|จำนวนครั้ง|ตรวจได้|ตรง|ผิดปั๊ม|เกิน \(เที่ยว\)|ขาด \(เที่ยว\))$/;
export const numFmtFor = (header) => (MONEY_HEAD.test(header) ? MONEY_FMT : COUNT_HEAD.test(header) ? COUNT_FMT : null);

// A cell as the exports build it: '' / null → empty, a number, a string, or xlDate() = { t:'n', v:serial, z:fmt }.
const isDateCell = (v) => v && typeof v === 'object' && v.t === 'n' && typeof v.v === 'number';

/**
 * sheets: [{ name, headers, rows, widths?, total?, filter? }]
 *   rows / total: arrays of cells (see above); widths: column widths in characters; filter: autofilter on the header.
 * Returns the ExcelJS workbook (the caller writes it).
 */
export function buildStyledWorkbook(ExcelJS, sheets) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'NeoSiam Logistics';
    for (const sh of sheets) {
        const ws = wb.addWorksheet(sh.name, { views: [{ state: 'frozen', ySplit: 1 }] });
        const fmts = sh.headers.map(h => numFmtFor(String(h)));
        const put = (rowIdx, values, style) => {
            const row = ws.getRow(rowIdx);
            for (let i = 0; i < sh.headers.length; i++) {
                const v = values[i];
                const c = row.getCell(i + 1);
                if (isDateCell(v)) { c.value = v.v; c.numFmt = v.z; } else {
                    c.value = v === '' || v == null ? null : v;
                    if (typeof v === 'number' && fmts[i]) c.numFmt = fmts[i];
                }
                style(c, i);
            }
        };
        put(1, sh.headers, (c) => { c.font = FONTS.head; c.fill = fill(NEO.navy); c.border = BORDER; c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; });
        ws.getRow(1).height = 24;
        sh.rows.forEach((r, k) => put(k + 2, r, (c) => {
            c.font = FONTS.body; c.border = BORDER; c.alignment = { vertical: 'middle' };
            if (k % 2 === 1) c.fill = fill(NEO.zebra);
        }));
        if (sh.total) put(sh.rows.length + 2, sh.total, (c) => { c.font = FONTS.total; c.fill = fill(NEO.fillTotal); c.border = BORDER_TOTAL; c.alignment = { vertical: 'middle' }; });
        sh.headers.forEach((h, i) => { ws.getColumn(i + 1).width = sh.widths?.[i] || Math.max(10, Math.min(String(h).length + 4, 30)); });
        // header + every data row (as the old SheetJS files); the total row stays outside so filtering never hides it
        if (sh.filter && sh.headers.length) ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: sh.rows.length + 1, column: sh.headers.length } };
    }
    return wb;
}

export const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
