// Run: node tests/money.test.mjs
import { roundHalfUp, round2, sumMoney, roundFields } from '../src/money.js';

let fail = 0;
const eq = (name, a, b) => {
    const ok = JSON.stringify(a) === JSON.stringify(b);
    if (!ok) fail++;
    console.log(ok ? 'PASS' : 'FAIL', name, ok ? '' : `${JSON.stringify(a)} != ${JSON.stringify(b)}`);
};
eq('half up where × 100 rounds down', [roundHalfUp(10.075), roundHalfUp(1.005), roundHalfUp(2.675)], [10.08, 1.01, 2.68]);
eq('negatives round away from zero like the positive', [roundHalfUp(-10.075), roundHalfUp(-0.005)], [-10.08, -0.01]);
eq('non-numbers never leak into totals', [round2(undefined), round2('abc'), round2(null), roundHalfUp(NaN)], [0, 0, 0, 0]);
eq('float noise removed from totals', [0.1 + 0.2 === 0.3, sumMoney([0.1, 0.2]), sumMoney([{ v: 3896.5 }, { v: 0.1 }, { v: 0.2 }], x => x.v)], [false, 0.3, 3896.8]);
eq('whole baht unchanged', [sumMoney([4700, 1500, 200]), round2(18684631.5)], [6400, 18684631.5]);
eq('noise that sums to zero is 0, never "-0" or 5.55e-17', [sumMoney([0.1, 0.2, -0.3]), Object.is(round2(-0.001), 0), Object.is(sumMoney([-0.1, 0.1]), 0), round2(5.551115123125783e-17)], [0, true, true, 0]);
eq('half a satang still rounds', [round2(0.005), round2(-0.005), round2(0.0049)], [0.01, -0.01, 0]);
eq('row fields rounded, others kept', roundFields({ driver: 'ก', count: 3, net: 0.1 + 0.2, liters: 10.123 }, ['net']), { driver: 'ก', count: 3, net: 0.3, liters: 10.123 });

console.log(fail ? `FAILED ${fail}` : 'ALL PASS');
process.exit(fail ? 1 : 0);
