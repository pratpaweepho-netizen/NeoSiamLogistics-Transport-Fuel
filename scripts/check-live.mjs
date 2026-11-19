// `npm run check:live -- [text to find in the live bundle] …`
// After a push: waits until the live site serves the same bundle as the local `npm run build` (dist/index.html),
// then optionally checks that new texts are in it. Exit 1 = the live site is still on another version.
import { readFileSync } from 'fs';

const SITE = 'https://neo-siam-logistics-transport-fuel-s.vercel.app';
const want = (readFileSync('dist/index.html', 'utf8').match(/assets\/(index-[\w-]+\.js)/) || [])[1];
if (!want) { console.error('no bundle in dist/index.html — run npm run build first'); process.exit(1); }
const args = process.argv.slice(2);
const noWait = args.includes('--no-wait'); // check once instead of waiting up to 6 minutes for the deploy
const texts = args.filter(a => a !== '--no-wait');
let live = '';
for (let i = 0; i < (noWait ? 1 : 24); i++) {
    const html = await (await fetch(`${SITE}/?cb=${Date.now()}`, { cache: 'no-store' })).text();
    live = (html.match(/assets\/(index-[\w-]+\.js)/) || [])[1] || '';
    if (live === want) break;
    await new Promise(r => setTimeout(r, 15000));
}
console.log(`local ${want} · live ${live || '(none)'} → ${live === want ? 'SAME' : 'DIFFERENT'}`);
if (live !== want) process.exit(1);
if (texts.length) {
    const js = await (await fetch(`${SITE}/assets/${live}`)).text();
    let missing = 0;
    for (const t of texts) { const n = js.split(t).length - 1; if (!n) missing++; console.log(`${n ? 'found  ' : 'MISSING'} "${t}" (${n})`); }
    if (missing) process.exit(1);
}
