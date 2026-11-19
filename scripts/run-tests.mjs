// `npm test`: run every tests/*.test.mjs (each prints PASS/FAIL and exits non-zero on failure) and fail if any fails.
import { readdirSync } from 'fs';
import { spawnSync } from 'child_process';
import { fileURLToPath } from 'url';
import path from 'path';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const files = readdirSync(path.join(root, 'tests')).filter(f => f.endsWith('.test.mjs')).sort();
let failed = 0;
for (const f of files) {
    const r = spawnSync(process.execPath, [path.join('tests', f)], { cwd: root, encoding: 'utf8' });
    const out = `${r.stdout || ''}${r.stderr || ''}`;
    const fails = out.split('\n').filter(l => l.startsWith('FAIL'));
    const ok = r.status === 0;
    if (!ok) failed++;
    console.log(`${ok ? 'ok  ' : 'FAIL'} ${f}${ok ? '' : `\n${fails.join('\n') || out.slice(-800)}`}`);
}
console.log(failed ? `\n${failed} of ${files.length} test files failed` : `\nall ${files.length} test files passed`);
process.exit(failed ? 1 : 0);
