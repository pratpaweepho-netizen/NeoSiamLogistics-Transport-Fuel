// Thai date/month entry boxes (user decision 19 Nov 2026): every <input type="date"> / <input type="month"> on the
// page is shown as a typed box "19/11/2569" / "11/2569" (Buddhist year) with a calendar button, the same on every
// computer whatever the browser language. The original input stays in the DOM (hidden) and keeps the ISO value, so
// all existing code that reads/writes `.value`, listens to `change` or submits the form works unchanged.

import { fmtDate, fmtMonth, parseThaiDate, parseThaiMonth } from './thaiDate.js';

const valueDesc = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
const rawGet = (el) => valueDesc.get.call(el);
const rawSet = (el, v) => valueDesc.set.call(el, v);

export function enhanceDateInput(orig) {
    if (!orig || orig.dataset.thDone) return;
    const isMonth = orig.getAttribute('type') === 'month';
    const fmt = isMonth ? fmtMonth : fmtDate;
    const parse = isMonth ? parseThaiMonth : parseThaiDate;
    const badMsg = isMonth ? 'เดือนไม่ถูกต้อง — พิมพ์ ดด/ปปปป เช่น 11/2569' : 'วันที่ไม่ถูกต้อง — พิมพ์ วว/ดด/ปปปป เช่น 19/11/2569';
    orig.dataset.thDone = '1';
    const full = orig.classList.contains('w-full');

    const wrap = document.createElement('div');
    wrap.className = `relative ${full ? 'flex w-full' : 'inline-flex'} items-center gap-1`;
    const txt = document.createElement('input');
    txt.type = 'text';
    txt.inputMode = 'numeric'; // phone number pads have no "/": 19112569 is accepted too (parseThaiDate)
    txt.autocomplete = 'off';
    txt.placeholder = isMonth ? 'ดด/ปปปป' : 'วว/ดด/ปปปป';
    txt.className = orig.className.replace(/\bw-full\b/, '') + (full ? ' flex-1 min-w-0' : '');
    if (!full) txt.style.width = isMonth ? '7rem' : '8.5rem';
    txt.title = isMonth ? 'เดือน/ปี พ.ศ.' : 'วัน/เดือน/ปี พ.ศ.';
    if (orig.id) {
        txt.id = `${orig.id}-th`;
        document.querySelectorAll(`label[for="${orig.id}"]`).forEach(l => { l.htmlFor = txt.id; });
    }
    if (orig.required) { txt.required = true; orig.required = false; }
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.textContent = '📅';
    btn.title = 'เลือกจากปฏิทิน';
    btn.tabIndex = -1;
    btn.className = 'shrink-0 px-2 py-2 rounded-md bg-slate-700 hover:bg-slate-600 text-sm leading-none';

    orig.parentNode.insertBefore(wrap, orig);
    wrap.append(txt, btn, orig);
    // hidden but still rendered (the calendar button opens its picker)
    Object.assign(orig.style, { position: 'absolute', left: '0', bottom: '0', width: '1px', height: '1px', opacity: '0', pointerEvents: 'none' });
    orig.tabIndex = -1;
    orig.setAttribute('aria-hidden', 'true');

    const markOk = () => { txt.setCustomValidity(''); txt.classList.remove('border-red-500'); };
    const show = () => { txt.value = fmt(rawGet(orig)); markOk(); };
    // code that sets `.value` (form resets, defaults) also updates the Thai box
    Object.defineProperty(orig, 'value', { configurable: true, get() { return rawGet(this); }, set(v) { rawSet(this, v); show(); } });
    show();

    const commit = () => {
        const t = txt.value.trim();
        const iso = t === '' ? '' : parse(t);
        if (t !== '' && !iso) {
            // invalid: blocks form submission (custom validity) and empties the stored value WITHOUT a change event
            // (filters keep what they show), so code reading `.value` at save time sees "no date", never the old one
            txt.setCustomValidity(badMsg);
            txt.classList.add('border-red-500');
            rawSet(orig, '');
            return false;
        }
        markOk();
        if (rawGet(orig) !== iso) {
            rawSet(orig, iso);
            orig.dispatchEvent(new Event('input', { bubbles: true }));
            orig.dispatchEvent(new Event('change', { bubbles: true }));
        }
        txt.value = fmt(iso);
        return true;
    };
    txt.addEventListener('change', commit);
    txt.addEventListener('keydown', (e) => { if (e.key === 'Enter') commit(); });
    // validity follows every keystroke, so a submit while the box still has focus is checked too
    txt.addEventListener('input', () => { const t = txt.value.trim(); txt.setCustomValidity(t !== '' && !parse(t) ? badMsg : ''); });
    txt._thCommit = commit; // the document-level submit guard commits every box before the app's own handler runs
    orig.addEventListener('change', show); // picked from the calendar
    btn.addEventListener('click', () => {
        // picking the same date fires no change event → clear a red "invalid" state now
        if (!txt.checkValidity() || txt.classList.contains('border-red-500')) { markOk(); txt.value = fmt(rawGet(orig)); }
        try { orig.showPicker(); } catch { txt.focus(); }
    });
    const form = orig.form;
    if (form && !form.dataset.thReset) {
        form.dataset.thReset = '1';
        // form.reset() restores values without the setter → refresh every Thai box of the form
        form.addEventListener('reset', () => setTimeout(() => form.querySelectorAll('input[data-th-done]').forEach(i => { i.value = rawGet(i); }), 0));
    }
}

const SELECTOR = 'input[type="date"]:not([data-th-done]), input[type="month"]:not([data-th-done])';
export const enhanceDateInputs = (root = document) => root.querySelectorAll(SELECTOR).forEach(enhanceDateInput);

/** Enhance what is on the page now and everything rendered later (views, modals, import previews). */
export function startThaiDateInputs() {
    enhanceDateInputs();
    // Before any form's own submit handler (capture on document runs first): commit every Thai box of the form — a
    // date typed but not yet left (no change event) must be the one saved — and stop the save if any box is invalid.
    document.addEventListener('submit', (e) => {
        const boxes = [...e.target.querySelectorAll('input')].filter(i => typeof i._thCommit === 'function');
        const bad = boxes.filter(i => !i._thCommit());
        if (bad.length) { e.preventDefault(); e.stopImmediatePropagation(); bad[0].reportValidity(); }
    }, true);
    new MutationObserver((list) => {
        for (const m of list) for (const n of m.addedNodes) {
            if (n.nodeType !== 1) continue;
            if (n.matches?.(SELECTOR)) enhanceDateInput(n);
            else if (n.querySelector?.(SELECTOR)) enhanceDateInputs(n);
        }
    }).observe(document.body, { childList: true, subtree: true });
}
