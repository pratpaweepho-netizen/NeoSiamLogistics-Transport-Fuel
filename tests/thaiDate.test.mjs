// Run: node tests/thaiDate.test.mjs
import { fmtDate, fmtMonth, fmtDateText, parseThaiDate, parseThaiMonth, xlDate } from '../src/thaiDate.js';
import { parseTemplateDate } from '../src/tripTemplate.js';

let fail = 0;
const eq = (name, a, b) => {
    const ok = JSON.stringify(a) === JSON.stringify(b);
    if (!ok) fail++;
    console.log(ok ? 'PASS' : 'FAIL', name, ok ? '' : `${JSON.stringify(a)} != ${JSON.stringify(b)}`);
};
eq('date → dd/mm/พ.ศ.', [fmtDate('2026-11-19'), fmtDate('2026-01-05')], ['19/11/2569', '05/01/2569']);
eq('timestamp → its LOCAL day', fmtDate(new Date(2026, 10, 17, 23, 30).toISOString()), '17/11/2569');
eq('non-ISO / impossible left as is', [fmtDate(''), fmtDate(undefined), fmtDate('ค้าง'), fmtDate('2026-02-30')], ['', '', 'ค้าง', '2026-02-30']);
eq('month → mm/พ.ศ.', [fmtMonth('2026-11'), fmtMonth('2025-01'), fmtMonth(''), fmtMonth('2026-13')], ['11/2569', '01/2568', '', '2026-13']);
eq('dates and months inside text', fmtDateText('อาจบันทึกขาล่องซ้ำ: เที่ยวขาขึ้น 2026-07-25 · (ระบบ 2026-03-04 / ไฟล์ 2026-03-05) · เดือน 2026-07, 2026-08'),
    'อาจบันทึกขาล่องซ้ำ: เที่ยวขาขึ้น 25/07/2569 · (ระบบ 04/03/2569 / ไฟล์ 05/03/2569) · เดือน 07/2569, 08/2569');
eq('ids and impossible values untouched', [fmtDateText('tdh_2026-07-02_ab12'), fmtDateText('2026-13-40'), fmtDateText('x2026-01-01'), fmtDateText('trips/2026-07/abc')], ['tdh_2026-07-02_ab12', '2026-13-40', 'x2026-01-01', 'trips/07/2569/abc']);
eq('typed พ.ศ. / ค.ศ. / short years / separators', [parseThaiDate('19/11/2569'), parseThaiDate('19/11/2026'), parseThaiDate('1.2.69'), parseThaiDate('19/11/26'), parseThaiDate('5-3-2569'), parseThaiDate(' 05/03/2569 ')],
    ['2026-11-19', '2026-11-19', '2026-02-01', '2026-11-19', '2026-03-05', '2026-03-05']);
eq('typed without separators (phone number pad)', [parseThaiDate('19112569'), parseThaiDate('191169'), parseThaiDate('19112026')], ['2026-11-19', '2026-11-19', '2026-11-19']);
eq('impossible / typo years / wrong order / text refused', [parseThaiDate('31/02/2569'), parseThaiDate('19/11/2469'), parseThaiDate('11/19/2026'), parseThaiDate('2026-11-19'), parseThaiDate('abc'), parseThaiDate('')],
    ['', '', '', '', '', '']);
eq('leap day', [parseThaiDate('29/02/2567'), parseThaiDate('29/02/2024'), parseThaiDate('29/02/2569')], ['2024-02-29', '2024-02-29', '']);
eq('typed months', [parseThaiMonth('11/2569'), parseThaiMonth('1/2026'), parseThaiMonth('03-69'), parseThaiMonth('112569'), parseThaiMonth('13/2569'), parseThaiMonth('2569')], ['2026-11', '2026-01', '2026-03', '2026-11', '', '']);
eq('same reading as the Excel trip template', ['19/11/2569', '1.2.69', '19/11/26', '19/11/2026'].map(s => parseThaiDate(s) === parseTemplateDate(s)), [true, true, true, true]);
eq('Excel date cell (serial of the calendar day, Buddhist-year format)', [xlDate('2026-11-19'), xlDate('2026-03-01').v, xlDate(''), xlDate('ค้าง')],
    [{ t: 'n', v: 46345, z: 'dd/mm/bbbb' }, 46082, '', 'ค้าง']);
eq('round trip', [parseThaiDate(fmtDate('2026-12-31')), parseThaiMonth(fmtMonth('2026-12'))], ['2026-12-31', '2026-12']);

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
