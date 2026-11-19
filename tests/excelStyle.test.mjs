// Run: node tests/excelStyle.test.mjs — builds a real workbook with ExcelJS, writes it, reads it back.
import ExcelJS from 'exceljs';
import { buildStyledWorkbook, NEO, MONEY_FMT, COUNT_FMT, numFmtFor } from '../src/excelStyle.js';
import { xlDate, XL_DATE_FMT } from '../src/thaiDate.js';

let fail = 0;
const eq = (name, a, b) => {
    const ok = JSON.stringify(a) === JSON.stringify(b);
    if (!ok) fail++;
    console.log(ok ? 'PASS' : 'FAIL', name, ok ? '' : `${JSON.stringify(a)} != ${JSON.stringify(b)}`);
};

eq('money / count columns by header', [numFmtFor('สุทธิโอน'), numFmtFor('น้ำมัน (บาท)'), numFmtFor('จำนวนคิว'), numFmtFor('เที่ยว'), numFmtFor('ลิตร'), numFmtFor('คนขับ')], [MONEY_FMT, MONEY_FMT, COUNT_FMT, COUNT_FMT, null, null]);

const wb = buildStyledWorkbook(ExcelJS, [
    { name: 'สรุปรายวัน', headers: ['วันที่', 'จำนวนคิว', 'น้ำมัน (ลิตร)', 'สุทธิโอน', 'หมายเหตุ'],
        rows: [[xlDate('2026-11-19'), 3, 120.5, 4700.5, ''], [xlDate('2026-11-20'), 1, 40, 1500, 'x']],
        total: ['รวม', 4, 160.5, 6200.5, ''], widths: [14, 10, 12, 12, 30] },
    { name: 'ผลตรวจ', headers: ['คนขับ', 'ผล'], rows: [['ก', 'ตรง'], ['ข', 'เกิน'], ['ค', 'ขาด']], filter: true, total: ['รวม', ''] },
]);
const buf = await wb.xlsx.writeBuffer();
const back = new ExcelJS.Workbook();
await back.xlsx.load(buf);
const ws = back.getWorksheet('สรุปรายวัน');
const c = (r, k) => ws.getRow(r).getCell(k);

eq('header stays on row 1 with the same text (files read back still work)', [1, 2, 3, 4, 5].map(k => c(1, k).value), ['วันที่', 'จำนวนคิว', 'น้ำมัน (ลิตร)', 'สุทธิโอน', 'หมายเหตุ']);
eq('header: Cordia New bold white on navy', [c(1, 1).font.name, c(1, 1).font.bold, c(1, 1).font.color.argb, c(1, 1).fill.fgColor.argb], ['Cordia New', true, NEO.white, NEO.navy]);
eq('date stays a real date cell shown as 19/11/2569', [c(2, 1).value instanceof Date && c(2, 1).value.toISOString().slice(0, 10), c(2, 1).numFmt], ['2026-11-19', XL_DATE_FMT]);
eq('values unchanged; money / count formats', [c(2, 2).value, c(2, 3).value, c(2, 4).value, c(2, 4).numFmt, c(2, 2).numFmt, c(2, 3).numFmt || null], [3, 120.5, 4700.5, MONEY_FMT, COUNT_FMT, null]);
eq('blank cell is empty (not an empty text that breaks sums)', c(2, 5).value, null);
eq('zebra on every second data row only', [c(2, 1).fill?.fgColor?.argb || null, c(3, 1).fill.fgColor.argb], [null, NEO.zebra]);
eq('total row: gold bold, light yellow, double bottom border', [c(4, 1).value, c(4, 4).value, c(4, 4).font.color.argb, c(4, 4).fill.fgColor.argb, c(4, 4).border.bottom.style], ['รวม', 6200.5, NEO.gold, NEO.fillTotal, 'double']);
eq('widths and frozen header', [ws.getColumn(1).width, ws.getColumn(5).width, ws.views[0].state, ws.views[0].ySplit], [14, 30, 'frozen', 1]);
const ws2 = back.getWorksheet('ผลตรวจ');
{
    // the filter must cover the data rows (header-only = filtering does nothing); the total row stays outside
    const { unzipSync, strFromU8 } = await import('fflate').catch(() => ({}));
    const JSZip = unzipSync ? null : (await import('jszip')).default;
    let xml = '';
    if (unzipSync) xml = strFromU8(unzipSync(new Uint8Array(buf))['xl/worksheets/sheet2.xml']);
    else xml = await (await JSZip.loadAsync(buf)).file('xl/worksheets/sheet2.xml').async('string');
    eq('second sheet: filter covers header + 3 data rows, not the total', [ws2.getRow(2).getCell(2).value, (xml.match(/<autoFilter ref="([^"]+)"/) || [])[1]], ['ตรง', 'A1:B4']);
}

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
