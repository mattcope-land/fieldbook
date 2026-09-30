'use strict';

// Fieldbook: simple books for a one-person practice.
// Everything lives in localStorage under KEY. Amounts are whole cents; dates are 'YYYY-MM-DD'.
// Backups are ordinary .xlsx files (SheetJS, vendored), and import reads them back.

const KEY = 'fieldbook';
const VERSION = 1;

const DEFAULT_CATEGORIES = {
    income: ['Consulting fees', 'Expense reimbursements', 'Other income'],
    expense: [
        'Research & databases',
        'Software & subscriptions',
        'Professional dues & licenses',
        'Insurance',
        'Legal & accounting',
        'Continuing education',
        'Office & supplies',
        'Phone & internet',
        'Travel',
        'Meals',
        'Contract help',
        'Bank & payment fees',
        'Marketing & website',
        'Other expenses'
    ]
};
const FALLBACK = { income: 'Other income', expense: 'Other expenses' };

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MON = MONTHS.map(m => m.slice(0, 3));
const MONEY_FMT = '"$"#,##0.00;[Red]-"$"#,##0.00';
const PAGE = 250;

const $ = (sel, el = document) => el.querySelector(sel);
const $$ = (sel, el = document) => [...el.querySelectorAll(sel)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);

/* ───────────────────────── Money & dates ───────────────────────── */

const currency = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' });
const money = cents => currency.format(cents / 100);
const moneyShort = cents => {
    const d = Math.abs(cents) / 100;
    if (d >= 1000) return '$' + (d / 1000).toFixed(d >= 10000 ? 0 : 1).replace(/\.0$/, '') + 'k';
    return '$' + Math.round(d);
};
// Accounting style for statements: losses in parentheses.
const acct = cents => cents < 0 ? `(${money(-cents)})` : money(cents);

// "$1,234.50", "1234.5", "(12.00)", "-12", 12.5 → cents (signed), or NaN
function parseMoney(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? Math.round(value * 100) : NaN;
    let s = String(value ?? '').trim();
    if (!s) return NaN;
    const negative = /^\(.*\)$/.test(s) || /^-|-$|^\$\s*-/.test(s);
    s = s.replace(/[^0-9.]/g, '');
    if (!s || (s.match(/\./g) || []).length > 1) return NaN;
    const cents = Math.round(parseFloat(s) * 100 + 1e-6);
    return negative ? -cents : cents;
}

const pad = n => String(n).padStart(2, '0');
const isoOf = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const today = () => isoOf(new Date());
const yearOf = iso => +iso.slice(0, 4);
const monthOf = iso => +iso.slice(5, 7);
const prettyDate = iso => `${MON[monthOf(iso) - 1]} ${+iso.slice(8, 10)}, ${iso.slice(0, 4)}`;
const longDate = iso => `${MONTHS[monthOf(iso) - 1]} ${+iso.slice(8, 10)}, ${iso.slice(0, 4)}`;
const validISO = s => /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(Date.parse(s + 'T00:00:00'));
const lastDay = (y, m) => new Date(y, m, 0).getDate();

// Excel serial day numbers ↔ ISO dates (done in UTC so time zones can't shift a day)
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const toSerial = iso => (Date.UTC(yearOf(iso), monthOf(iso) - 1, +iso.slice(8, 10)) - EXCEL_EPOCH) / 86400000;
const fromSerial = n => new Date(EXCEL_EPOCH + Math.round(n) * 86400000).toISOString().slice(0, 10);

function parseDate(value) {
    if (value instanceof Date && !isNaN(value)) return isoOf(value);
    if (typeof value === 'number') return value > 1000 && value < 100000 ? fromSerial(value) : null;
    const s = String(value ?? '').trim();
    if (!s) return null;
    let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (m) return checked(+m[1], +m[2], +m[3]);
    m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})$/); // US month/day/year
    if (m) return checked(+m[3] < 100 ? 2000 + +m[3] : +m[3], +m[1], +m[2]);
    const t = Date.parse(s);
    return isNaN(t) ? null : isoOf(new Date(t));

    function checked(y, mo, d) {
        if (mo < 1 || mo > 12 || d < 1 || d > lastDay(y, mo)) return null;
        return `${y}-${pad(mo)}-${pad(d)}`;
    }
}

function ago(iso) {
    const days = Math.round((Date.parse(today()) - Date.parse(iso.slice(0, 10))) / 86400000);
    if (days <= 0) return 'today';
    if (days === 1) return 'yesterday';
    if (days < 14) return `${days} days ago`;
    return 'on ' + prettyDate(iso.slice(0, 10)).replace(/, \d{4}$/, m => yearOf(iso) === new Date().getFullYear() ? '' : m);
}
const daysSince = iso => (Date.now() - Date.parse(iso)) / 86400000;

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const plural = (n, word, many = word + 's') => `${n.toLocaleString()} ${n === 1 ? word : many}`;

/* ───────────────────────── The books ───────────────────────── */

function freshBooks() {
    return {
        version: VERSION,
        business: '',
        entries: [],
        categories: structuredClone(DEFAULT_CATEGORIES),
        lastBackupAt: null,
        changesSinceBackup: 0
    };
}

function validEntry(e) {
    return e && typeof e.id === 'string' && validISO(e.date) && (e.type === 'income' || e.type === 'expense')
        && Number.isInteger(e.amount) && e.amount > 0 && typeof e.category === 'string';
}

function normalize(raw) {
    const books = { ...freshBooks(), ...raw };
    books.categories = {
        income: Array.isArray(raw.categories?.income) ? raw.categories.income : [...DEFAULT_CATEGORIES.income],
        expense: Array.isArray(raw.categories?.expense) ? raw.categories.expense : [...DEFAULT_CATEGORIES.expense]
    };
    books.entries = (Array.isArray(raw.entries) ? raw.entries : []).filter(validEntry)
        .map(e => ({ party: '', description: '', ...e }));
    return books;
}

function load() {
    let raw;
    try { raw = localStorage.getItem(KEY); } catch { return freshBooks(); }
    if (!raw) return freshBooks();
    try {
        return normalize(JSON.parse(raw));
    } catch (err) {
        // Never overwrite something we can't read: set it aside so it can be recovered.
        console.error('Fieldbook: saved data could not be read', err);
        try { localStorage.setItem(`${KEY}.unreadable.${Date.now()}`, raw); } catch { /* nothing more to do */ }
        return freshBooks();
    }
}

let books = load();
let saveFailed = false;

function save() {
    try {
        localStorage.setItem(KEY, JSON.stringify(books));
        saveFailed = false;
    } catch (err) {
        console.error(err);
        if (!saveFailed) toast('This browser wouldn’t save your last change. Download a backup now to keep it.', { action: ['Download backup', downloadBackup], sticky: true });
        saveFailed = true;
    }
}

// Every change goes through here: snapshot, change, save, redraw, and offer Undo.
function change(message, fn, { counts = 1, undo = true } = {}) {
    const before = JSON.stringify(books);
    const result = fn();
    books.changesSinceBackup += counts;
    save();
    render();
    if (message) toast(message, undo ? { action: ['Undo', () => { books = normalize(JSON.parse(before)); save(); render(); toast('Undone.'); }] } : {});
    return result;
}

// Ask the browser not to evict our storage when space runs low (Chrome/Firefox honour this).
if (navigator.storage?.persist) navigator.storage.persisted().then(p => p || navigator.storage.persist()).catch(() => {});

// Another tab changed the books
addEventListener('storage', e => { if (e.key === KEY) { books = load(); render(); } });

/* ───────────────────────── Queries ───────────────────────── */

const byNewest = (a, b) => b.date.localeCompare(a.date) || (b.createdAt || 0) - (a.createdAt || 0);
const inRange = (from, to) => e => e.date >= from && e.date <= to;
const sum = list => list.reduce((t, e) => t + e.amount, 0);

function totals(list) {
    let income = 0, expense = 0;
    for (const e of list) e.type === 'income' ? income += e.amount : expense += e.amount;
    return { income, expense, net: income - expense };
}

function years() {
    const set = new Set(books.entries.map(e => yearOf(e.date)));
    set.add(new Date().getFullYear());
    return [...set].sort((a, b) => b - a);
}

// Categories in the user's order, then any stray ones (from imports) alphabetically
function categoryOrder(type, list) {
    const known = books.categories[type];
    const extra = [...new Set(list.filter(e => e.type === type).map(e => e.category))].filter(c => !known.includes(c)).sort();
    return [...known, ...extra];
}

function usage(type, category) {
    return books.entries.filter(e => e.type === type && e.category === category).length;
}

/* ───────────────────────── Routing ───────────────────────── */

const views = ['ledger', 'reports', 'case'];
let view = 'ledger';

function route() {
    const next = location.hash.slice(1);
    view = views.includes(next) ? next : 'ledger';
    for (const v of views) $(`#view-${v}`).hidden = v !== view;
    for (const a of $$('.tabs a')) a.dataset.tab === view ? a.setAttribute('aria-current', 'page') : a.removeAttribute('aria-current');
    render();
    scrollTo({ top: 0 });
}
addEventListener('hashchange', route);

function render() {
    renderChrome();
    if (view === 'ledger') renderLedger();
    if (view === 'reports') renderReports();
    if (view === 'case') renderCase();
}

function renderChrome() {
    const name = books.business.trim() || 'My Practice';
    for (const el of $$('[data-business]')) el.textContent = name;
    document.title = books.business.trim() ? `Fieldbook · ${books.business.trim()}` : 'Fieldbook';
    $('#sample-banner').hidden = !books.entries.some(e => e.sample);

    const { text, stale } = backupState();
    $('#backup-text').textContent = text;
    $('#backup-tag').classList.toggle('stale', stale);
}

function backupState() {
    const n = books.changesSinceBackup;
    if (!books.lastBackupAt) {
        if (!books.entries.length || books.entries.every(e => e.sample)) return { text: 'No backup needed yet', stale: false };
        return { text: 'Not backed up yet', stale: true };
    }
    const when = ago(books.lastBackupAt);
    const stale = n > 0 && daysSince(books.lastBackupAt) >= 7;
    return { text: n ? `Backed up ${when} · ${n} new` : `Backed up ${when}`, stale };
}

/* ───────────────────────── Ledger ───────────────────────── */

const filters = { q: '', year: String(new Date().getFullYear()), type: '', category: '' };
let shown = PAGE;
let justAdded = null;
let categoryTouched = false;

const entryForm = $('#entry-form');
const formType = () => entryForm.elements.type.value;
const lastCategory = { income: null, expense: null };

function fillCategories(select, type, keep) {
    const cats = books.categories[type];
    select.innerHTML = cats.map(c => `<option>${esc(c)}</option>`).join('');
    if (keep && !cats.includes(keep)) select.insertAdjacentHTML('beforeend', `<option>${esc(keep)}</option>`);
    select.value = keep || cats[0] || '';
}

function syncFormType(form, keepCategory) {
    const type = form.elements.type.value;
    form.dataset.type = type;
    for (const el of $$('[data-party-label]', form)) el.textContent = type === 'income' ? 'Client' : 'Paid to';
    fillCategories(form.elements.category, type, keepCategory);
}

function renderLedger() {
    // New entry form
    if (!entryForm.elements.date.value) entryForm.elements.date.value = today();
    const type = formType();
    const current = entryForm.elements.category.value;
    const cats = books.categories[type];
    syncFormType(entryForm, cats.includes(current) ? current : cats.includes(lastCategory[type]) ? lastCategory[type] : null);
    fillSuggestions();

    // Filters
    const yearSel = $('#f-year');
    yearSel.innerHTML = years().map(y => `<option value="${y}">${y}</option>`).join('') + '<option value="all">All years</option>';
    yearSel.value = filters.year;
    if (yearSel.value !== filters.year) yearSel.value = filters.year = String(new Date().getFullYear());

    const catSel = $('#f-cat');
    const group = (label, type) => `<optgroup label="${label}">${categoryOrder(type, books.entries).map(c => `<option value="${type}:${esc(c)}">${esc(c)}</option>`).join('')}</optgroup>`;
    catSel.innerHTML = '<option value="">All categories</option>' + group('Money in', 'income') + group('Money out', 'expense');
    catSel.value = filters.category;
    if (catSel.value !== filters.category) filters.category = catSel.value = '';

    // Headline numbers for the chosen year
    const yearList = filters.year === 'all' ? books.entries : books.entries.filter(e => String(yearOf(e.date)) === filters.year);
    const t = totals(yearList);
    const label = filters.year === 'all' ? 'All years' : filters.year === String(new Date().getFullYear()) ? `${filters.year} so far` : filters.year;
    const margin = t.income ? Math.round(t.net / t.income * 100) : null;
    $('#summary').innerHTML = `
        <div class="specimen" data-kind="in"><div class="specimen-label">Money in</div>
            <div class="specimen-value">${money(t.income)}</div><div class="specimen-note">${label}</div></div>
        <div class="specimen" data-kind="out"><div class="specimen-label">Money out</div>
            <div class="specimen-value">${money(t.expense)}</div><div class="specimen-note">${label}</div></div>
        <div class="specimen" data-kind="net"><div class="specimen-label">${t.net < 0 ? 'Net loss' : 'Profit'}</div>
            <div class="specimen-value ${t.net < 0 ? 'neg' : ''}">${acct(t.net)}</div>
            <div class="specimen-note">${margin === null ? label : `${margin}% of money in kept`}</div></div>`;

    // The list
    const box = $('#ledger');
    if (!books.entries.length) {
        box.innerHTML = `
            <div class="empty">
                <svg class="jar" aria-hidden="true"><use href="#i-jar"/></svg>
                <h2>Nothing collected yet</h2>
                <p>Add your first entry above. You can also bring in the spreadsheet you use now,
                   or fill the books with sample entries to see how everything works.</p>
                <div class="actions">
                    <button class="btn" type="button" data-act="import">Import my spreadsheet…</button>
                    <button class="btn" type="button" data-act="sample">Try sample entries</button>
                </div>
            </div>`;
        return;
    }

    const words = filters.q.toLowerCase().split(/\s+/).filter(Boolean);
    const [ftype, fcat] = filters.category ? [filters.category.split(':')[0], filters.category.slice(filters.category.indexOf(':') + 1)] : [];
    const list = yearList.filter(e =>
        (!filters.type || e.type === filters.type) &&
        (!ftype || (e.type === ftype && e.category === fcat)) &&
        (!words.length || words.every(w => haystack(e).includes(w)))
    ).sort(byNewest);

    if (!list.length) {
        box.innerHTML = `<div class="empty"><p class="muted">No entries match${filters.q ? ` “${esc(filters.q)}”` : ''}.</p>
            <button class="link" type="button" data-act="clear-filters">Clear the filters</button></div>`;
        return;
    }

    const filtered = filters.q || filters.type || filters.category;
    let html = '';
    if (filtered) {
        const ft = totals(list);
        html += `<p class="muted" style="margin:0 4px 14px">${plural(list.length, 'entry', 'entries')}
            ${ft.income ? ` · in ${money(ft.income)}` : ''}${ft.expense ? ` · out ${money(ft.expense)}` : ''}</p>`;
    }

    const groups = new Map();
    for (const e of list) {
        const k = e.date.slice(0, 7);
        if (!groups.has(k)) groups.set(k, []);
        groups.get(k).push(e);
    }
    let count = 0;
    for (const [ym, items] of groups) {
        if (count >= shown) break;
        const mt = totals(items);
        const visible = items.slice(0, shown - count);
        count += visible.length;
        html += `
            <section class="month">
                <div class="month-head">
                    <h2>${MONTHS[monthOf(ym + '-01') - 1]} <span class="year">${ym.slice(0, 4)}</span></h2>
                    <div class="month-sum num">
                        ${mt.income ? `<span>in <b>${money(mt.income)}</b></span>` : ''}
                        ${mt.expense ? `<span>out <b>${money(mt.expense)}</b></span>` : ''}
                        ${mt.income && mt.expense ? `<span>net <b class="${mt.net < 0 ? 'neg' : ''}">${acct(mt.net)}</b></span>` : ''}
                    </div>
                </div>
                <ul class="rows">${visible.map(rowHTML).join('')}</ul>
            </section>`;
    }
    if (list.length > shown) html += `<button class="btn more" type="button" data-act="more">Show more (${(list.length - shown).toLocaleString()} older)</button>`;
    box.innerHTML = html;

    if (justAdded) {
        const row = box.querySelector(`[data-id="${justAdded}"]`);
        row?.classList.add('flash');
        justAdded = null;
    }
}

const haystack = e => `${e.description} ${e.party} ${e.category} ${money(e.amount)} ${e.amount / 100} ${prettyDate(e.date)}`.toLowerCase();

function rowHTML(e) {
    const main = e.description || e.party || e.category;
    const sub = e.description ? e.party : '';
    return `
        <li class="row" tabindex="0" data-id="${e.id}" aria-label="${esc(`${prettyDate(e.date)}, ${main}, ${e.type === 'income' ? 'money in' : 'money out'} ${money(e.amount)}`)}">
            <div class="row-day"><b>${+e.date.slice(8, 10)}</b>${MON[monthOf(e.date) - 1]}</div>
            <div class="row-main"><div class="row-desc">${esc(main)}</div>${sub ? `<div class="row-party">${esc(sub)}</div>` : ''}</div>
            <span class="tag" title="${esc(e.category)}">${esc(e.category)}</span>
            <div class="amt num ${e.type === 'income' ? 'in' : 'out'}">${e.type === 'income' ? '+' : '−'}${money(e.amount)}</div>
        </li>`;
}

function fillSuggestions() {
    const recent = [...books.entries].sort(byNewest);
    const unique = key => [...new Set(recent.map(e => e[key]).filter(Boolean))].slice(0, 300);
    $('#parties').innerHTML = unique('party').map(p => `<option value="${esc(p)}">`).join('');
    $('#descriptions').innerHTML = unique('description').map(d => `<option value="${esc(d)}">`).join('');
}

// When a familiar client or description is typed, pick the category used last time
function guessCategory(form, touched) {
    if (touched) return;
    const type = form.elements.type.value;
    const party = form.elements.party.value.trim().toLowerCase();
    const desc = form.elements.description.value.trim().toLowerCase();
    const match = [...books.entries].sort(byNewest).find(e => e.type === type &&
        ((desc && e.description.toLowerCase() === desc) || (party && e.party.toLowerCase() === party)));
    if (match && books.categories[type].includes(match.category)) form.elements.category.value = match.category;
}

function readForm(form) {
    const f = form.elements;
    const cents = parseMoney(f.amount.value);
    if (!validISO(f.date.value)) return { error: 'Please choose a date.', focus: f.date };
    if (!(Math.abs(cents) > 0)) return { error: 'Please enter an amount, like 1250 or 1,250.00.', focus: f.amount };
    if (!f.category.value) return { error: 'Please choose a category (add one in The Case).', focus: f.category };
    return {
        entry: {
            date: f.date.value,
            type: f.type.value,
            amount: Math.abs(cents),
            category: f.category.value,
            party: f.party.value.trim(),
            description: f.description.value.trim()
        }
    };
}

entryForm.addEventListener('change', e => {
    if (e.target.name === 'type') {
        categoryTouched = false;
        const last = lastCategory[formType()];
        syncFormType(entryForm, books.categories[formType()].includes(last) ? last : null);
        guessCategory(entryForm, false);
    }
    if (e.target.name === 'category') categoryTouched = true;
    if (e.target.name === 'party' || e.target.name === 'description') guessCategory(entryForm, categoryTouched);
});

entryForm.addEventListener('submit', e => {
    e.preventDefault();
    const { entry, error, focus } = readForm(entryForm);
    $('#entry-error').textContent = error || '';
    if (error) return focus.focus();

    const record = { id: uid(), ...entry, createdAt: Date.now() };
    lastCategory[entry.type] = entry.category;
    justAdded = record.id;
    // Make sure the new entry is visible in the list
    if (filters.year !== 'all' && String(yearOf(entry.date)) !== filters.year) filters.year = String(yearOf(entry.date));
    filters.q = ''; filters.type = ''; filters.category = ''; $('#q').value = ''; $('#f-type').value = '';
    change(`Added ${entry.type === 'income' ? 'money in' : 'money out'} of ${money(entry.amount)}.`, () => books.entries.push(record));

    for (const name of ['amount', 'party', 'description']) entryForm.elements[name].value = '';
    categoryTouched = false;
    entryForm.elements.amount.focus();
});

$('#q').addEventListener('input', e => { filters.q = e.target.value; shown = PAGE; renderLedger(); });
$('#f-year').addEventListener('change', e => { filters.year = e.target.value; shown = PAGE; renderLedger(); });
$('#f-type').addEventListener('change', e => { filters.type = e.target.value; shown = PAGE; renderLedger(); });
$('#f-cat').addEventListener('change', e => { filters.category = e.target.value; shown = PAGE; renderLedger(); });

$('#ledger').addEventListener('click', e => {
    const act = e.target.closest('[data-act]')?.dataset.act;
    if (act === 'more') { shown += PAGE; renderLedger(); return; }
    if (act === 'import') return $('#import-file').click();
    if (act === 'sample') return loadSample();
    if (act === 'clear-filters') {
        Object.assign(filters, { q: '', type: '', category: '' });
        $('#q').value = ''; $('#f-type').value = '';
        return renderLedger();
    }
    const row = e.target.closest('.row');
    if (row) openEdit(row.dataset.id);
});
$('#ledger').addEventListener('keydown', e => {
    const row = e.target.closest('.row');
    if (row && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); openEdit(row.dataset.id); }
});

/* ───────────────────────── Edit dialog ───────────────────────── */

const editDialog = $('#edit-dialog');
const editForm = $('#edit-form');
let editing = null;

function openEdit(id) {
    const entry = books.entries.find(e => e.id === id);
    if (!entry) return;
    editing = entry;
    const f = editForm.elements;
    f.type.value = entry.type;
    syncFormType(editForm, entry.category);
    f.date.value = entry.date;
    f.amount.value = (entry.amount / 100).toFixed(2);
    f.party.value = entry.party;
    f.description.value = entry.description;
    $('.form-error', editForm).textContent = '';
    editDialog.showModal();
}

editForm.addEventListener('change', e => {
    if (e.target.name === 'type') syncFormType(editForm, editing?.type === editForm.elements.type.value ? editing.category : null);
});

editForm.addEventListener('submit', e => {
    const action = e.submitter?.value;
    if (action === 'save') {
        const { entry, error, focus } = readForm(editForm);
        if (error) {
            e.preventDefault();
            $('.form-error', editForm).textContent = error;
            return focus.focus();
        }
        const id = editing.id;
        const same = Object.keys(entry).every(k => entry[k] === editing[k]);
        if (!same) change('Entry saved.', () => {
            const target = books.entries.find(x => x.id === id);
            Object.assign(target, entry, { updatedAt: Date.now() });
            delete target.sample;
        });
    }
    if (action === 'delete') {
        const id = editing.id;
        change('Entry deleted.', () => { books.entries = books.entries.filter(x => x.id !== id); });
    }
    editing = null;
});

/* ───────────────────────── Reports ───────────────────────── */

const report = { year: new Date().getFullYear(), span: 'year', from: '', to: '' };
let monthTable = false;

$('#r-months').innerHTML = MONTHS.map((m, i) => `<option value="m${i + 1}">${m}</option>`).join('');

function period() {
    const y = report.year;
    const s = report.span;
    const cur = new Date().getFullYear();
    if (s === 'custom') {
        let from = validISO(report.from) ? report.from : `${y}-01-01`;
        let to = validISO(report.to) ? report.to : `${y}-12-31`;
        if (from > to) [from, to] = [to, from];
        return { from, to, label: `${longDate(from)} – ${longDate(to)}` };
    }
    if (s === 'ytd') {
        const to = y === cur ? today() : `${y}-12-31`;
        return { from: `${y}-01-01`, to, label: `January 1 – ${longDate(to)}` };
    }
    if (s[0] === 'q') {
        const q = +s[1];
        const m1 = (q - 1) * 3 + 1, m3 = m1 + 2;
        return { from: `${y}-${pad(m1)}-01`, to: `${y}-${pad(m3)}-${lastDay(y, m3)}`, label: `Q${q} ${y} · ${MONTHS[m1 - 1]} 1 – ${MONTHS[m3 - 1]} ${lastDay(y, m3)}` };
    }
    if (s[0] === 'm') {
        const m = +s.slice(1);
        return { from: `${y}-${pad(m)}-01`, to: `${y}-${pad(m)}-${lastDay(y, m)}`, label: `${MONTHS[m - 1]} ${y}` };
    }
    const label = y === cur ? `January 1 – December 31, ${y} (year in progress)` : `January 1 – December 31, ${y}`;
    return { from: `${y}-01-01`, to: `${y}-12-31`, label };
}

function profitAndLoss(p) {
    const list = books.entries.filter(inRange(p.from, p.to));
    const lines = type => categoryOrder(type, list)
        .map(c => ({ category: c, amount: sum(list.filter(e => e.type === type && e.category === c)) }))
        .filter(l => l.amount);
    return { list, incomeLines: lines('income'), expenseLines: lines('expense'), ...totals(list) };
}

function renderReports() {
    const ys = years();
    if (!ys.includes(report.year)) report.year = ys[0];
    $('#r-year').innerHTML = ys.map(y => `<option>${y}</option>`).join('');
    $('#r-year').value = report.year;
    $('#r-span').value = report.span;
    $('#r-custom').hidden = report.span !== 'custom';

    const p = period();
    if (report.span === 'custom') { $('#r-from').value = p.from; $('#r-to').value = p.to; }
    const pl = profitAndLoss(p);
    const pct = cents => pl.income ? `${(cents / pl.income * 100).toFixed(1)}%` : '';
    const lineRows = (lines, none) => lines.length
        ? lines.map(l => `<tr class="st-line"><td>${esc(l.category)}</td><td class="num">${money(l.amount)}</td><td class="num pct">${pct(l.amount)}</td></tr>`).join('')
        : `<tr class="st-none"><td colspan="3">${none}</td></tr>`;

    $('#pnl').innerHTML = `
        <header class="st-head">
            <div class="st-business">${esc(books.business.trim() || 'My Practice')}</div>
            <h2 class="st-title">Profit &amp; Loss</h2>
            <div class="st-period">${esc(p.label)}</div>
            <svg class="sprig" aria-hidden="true"><use href="#i-sprig"/></svg>
        </header>
        <table class="st-table">
            <tbody>
                <tr class="st-section"><th colspan="3">Income</th></tr>
                ${lineRows(pl.incomeLines, 'No money in during this period')}
                <tr class="st-total"><td>Total income</td><td class="num">${money(pl.income)}</td><td class="num pct"></td></tr>
                <tr class="st-section"><th colspan="3">Expenses</th></tr>
                ${lineRows(pl.expenseLines, 'No money out during this period')}
                <tr class="st-total"><td>Total expenses</td><td class="num">${money(pl.expense)}</td><td class="num pct">${pct(pl.expense)}</td></tr>
                <tr class="st-net"><td>${pl.net < 0 ? 'Net loss' : 'Net profit'}</td><td class="num ${pl.net < 0 ? 'neg' : ''}">${acct(pl.net)}</td><td class="num pct">${pct(pl.net)}</td></tr>
            </tbody>
        </table>
        <p class="st-foot">Cash basis · ${plural(pl.list.length, 'entry', 'entries')} · prepared ${prettyDate(today())}</p>`;

    renderMonths(p);
    renderClients(p);
}

function renderMonths(p) {
    const y = report.year;
    const data = MONTHS.map((name, i) => {
        const list = books.entries.filter(e => yearOf(e.date) === y && monthOf(e.date) === i + 1);
        const from = `${y}-${pad(i + 1)}-01`, to = `${y}-${pad(i + 1)}-${lastDay(y, i + 1)}`;
        return { name, ...totals(list), active: to >= p.from && from <= p.to };
    });
    const box = $('#months');
    $('#toggle-month-table').textContent = monthTable ? 'Show as chart' : 'Show as table';
    $('#toggle-month-table').setAttribute('aria-pressed', monthTable);

    if (monthTable) {
        const t = totals(books.entries.filter(e => yearOf(e.date) === y));
        box.innerHTML = `<table class="table"><thead><tr><th>${y}</th><th class="num">In</th><th class="num">Out</th><th class="num">Net</th></tr></thead><tbody>
            ${data.map(d => `<tr><td>${d.name}</td><td class="num">${money(d.income)}</td><td class="num">${money(d.expense)}</td><td class="num ${d.net < 0 ? 'neg' : ''}">${acct(d.net)}</td></tr>`).join('')}
            </tbody><tfoot><tr><td>Year</td><td class="num">${money(t.income)}</td><td class="num">${money(t.expense)}</td><td class="num ${t.net < 0 ? 'neg' : ''}">${acct(t.net)}</td></tr></tfoot></table>`;
        return;
    }

    const W = 640, H = 230, L = 46, T = 10, B = 26;
    const plotW = W - L - 4, plotH = H - T - B;
    const max = Math.max(...data.map(d => Math.max(d.income, d.expense)), 1);
    const step = niceStep(max / 4);
    const top = Math.ceil(max / step) * step;
    const yPos = v => T + plotH - v / top * plotH;
    const slot = plotW / 12;
    const bw = Math.min(16, slot * 0.32);

    let svg = '';
    for (let v = 0; v <= top + 1; v += step) {
        const yy = yPos(v).toFixed(1);
        svg += `<line class="${v ? 'grid' : 'base'}" x1="${L}" x2="${W - 4}" y1="${yy}" y2="${yy}"/>
                <text class="axis" x="${L - 8}" y="${+yy + 4}" text-anchor="end">${moneyShort(v)}</text>`;
    }
    data.forEach((d, i) => {
        const cx = L + slot * i + slot / 2;
        const op = d.active ? 1 : 0.3;
        svg += `<g opacity="${op}">${bar(cx - bw - 1, d.income, 'bar-in')}${bar(cx + 1, d.expense, 'bar-out')}</g>
                <text class="axis" x="${cx}" y="${H - 8}" text-anchor="middle" opacity="${d.active ? 1 : 0.6}">${MON[i]}</text>
                <rect class="hit" data-m="${i}" x="${L + slot * i}" y="${T}" width="${slot}" height="${plotH}" rx="3"/>`;
    });

    function bar(x, v, cls) {
        if (!v) return '';
        const h = Math.max(v / top * plotH, 1.5);
        const r = Math.min(4, h, bw / 2);
        const y0 = T + plotH, y1 = y0 - h;
        return `<path class="${cls}" d="M${x},${y0} V${y1 + r} Q${x},${y1} ${x + r},${y1} H${x + bw - r} Q${x + bw},${y1} ${x + bw},${y1 + r} V${y0} Z"/>`;
    }

    box.innerHTML = `
        <div class="legend"><span><i style="background:var(--in)"></i>Money in</span><span><i style="background:var(--out)"></i>Money out</span></div>
        <svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Money in and out by month, ${y}">${svg}</svg>`;

    const tip = $('#tooltip');
    const svgEl = $('svg.chart', box);
    const show = (rect, x, yy) => {
        const d = data[+rect.dataset.m];
        $$('.hit.on', svgEl).forEach(r => r.classList.remove('on'));
        rect.classList.add('on');
        tip.innerHTML = `<h4>${d.name} ${y}</h4>
            <div><span><i style="background:var(--in)"></i>In</span><span class="num">${money(d.income)}</span></div>
            <div><span><i style="background:var(--out)"></i>Out</span><span class="num">${money(d.expense)}</span></div>
            <div><span>Net</span><span class="num ${d.net < 0 ? 'neg' : ''}">${acct(d.net)}</span></div>`;
        tip.hidden = false;
        const w = tip.offsetWidth, h = tip.offsetHeight;
        tip.style.left = Math.min(innerWidth - w - 8, Math.max(8, x + 14)) + 'px';
        tip.style.top = Math.max(8, yy - h - 12) + 'px';
    };
    const hide = () => { tip.hidden = true; $$('.hit.on', svgEl).forEach(r => r.classList.remove('on')); };
    svgEl.addEventListener('pointermove', e => { const r = e.target.closest('.hit'); r ? show(r, e.clientX, e.clientY) : hide(); });
    svgEl.addEventListener('pointerleave', hide);
    svgEl.addEventListener('pointerdown', e => { const r = e.target.closest('.hit'); if (r) show(r, e.clientX, e.clientY); });
}

function niceStep(raw) {
    const mag = 10 ** Math.floor(Math.log10(raw));
    for (const m of [1, 2, 2.5, 5, 10]) if (m * mag >= raw) return m * mag;
    return 10 * mag;
}

function renderClients(p) {
    const list = books.entries.filter(e => e.type === 'income' && e.date >= p.from && e.date <= p.to);
    const box = $('#clients');
    if (!list.length) { box.innerHTML = '<p class="muted">No money in during this period.</p>'; return; }
    const map = new Map();
    for (const e of list) {
        const k = e.party || 'No client named';
        const row = map.get(k) || { name: k, amount: 0, count: 0 };
        row.amount += e.amount; row.count++;
        map.set(k, row);
    }
    let rows = [...map.values()].sort((a, b) => b.amount - a.amount);
    if (rows.length > 10) {
        const rest = rows.slice(9);
        rows = [...rows.slice(0, 9), { name: `${rest.length} other clients`, amount: sum(rest), count: rest.reduce((t, r) => t + r.count, 0) }];
    }
    const total = sum(list);
    box.innerHTML = `<table class="table"><thead><tr><th>Client</th><th class="num">Share</th><th class="num">Amount</th></tr></thead><tbody>
        ${rows.map(r => `<tr><td class="share">${esc(r.name)}<span class="share-bar" style="width:${(r.amount / rows[0].amount * 100).toFixed(1)}%"></span></td>
            <td class="num muted">${Math.round(r.amount / total * 100)}%</td><td class="num">${money(r.amount)}</td></tr>`).join('')}
        </tbody><tfoot><tr><td>${plural(map.size, 'client')}</td><td></td><td class="num">${money(total)}</td></tr></tfoot></table>`;
}

$('#r-year').addEventListener('change', e => { report.year = +e.target.value; report.from = report.to = ''; renderReports(); });
$('#r-span').addEventListener('change', e => { report.span = e.target.value; renderReports(); });
$('#r-from').addEventListener('change', e => { report.from = e.target.value; renderReports(); });
$('#r-to').addEventListener('change', e => { report.to = e.target.value; renderReports(); });
$('#toggle-month-table').addEventListener('click', () => { monthTable = !monthTable; renderMonths(period()); });
$('#print-pnl').addEventListener('click', () => print());
$('#export-pnl').addEventListener('click', exportPnl);

/* ───────────────────────── Excel ───────────────────────── */

function sheet(rows, { widths, money: moneyCols = [], dates = [], moneyFrom = 0 } = {}) {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    rows.forEach((row, r) => {
        row.forEach((v, c) => {
            if (typeof v !== 'number') return;
            const cell = ws[XLSX.utils.encode_cell({ r, c })];
            if (dates.includes(c)) cell.z = 'yyyy-mm-dd';
            else if (moneyCols.includes(c) || (moneyFrom && c >= moneyFrom)) cell.z = MONEY_FMT;
        });
    });
    if (widths) ws['!cols'] = widths.map(wch => ({ wch }));
    return ws;
}

const entryHeader = ['Date', 'Type', 'Amount', 'Category', 'Client / Payee', 'Description', 'ID'];
const entryRow = e => [toSerial(e.date), e.type === 'income' ? 'Money in' : 'Money out', e.amount / 100, e.category, e.party, e.description, e.id];
const entrySheet = list => sheet([entryHeader, ...[...list].sort((a, b) => a.date.localeCompare(b.date)).map(entryRow)],
    { widths: [12, 11, 13, 28, 28, 44, 14], money: [2], dates: [0] });

function pnlRows(pl, p) {
    const rows = [[books.business.trim() || 'My Practice'], ['Profit & Loss'], [p.label], ['Cash basis'], [], ['Income']];
    for (const l of pl.incomeLines) rows.push(['    ' + l.category, l.amount / 100]);
    rows.push(['Total income', pl.income / 100], [], ['Expenses']);
    for (const l of pl.expenseLines) rows.push(['    ' + l.category, l.amount / 100]);
    rows.push(['Total expenses', pl.expense / 100], [], [pl.net < 0 ? 'Net loss' : 'Net profit', pl.net / 100]);
    return rows;
}

function exportPnl() {
    const p = period();
    const pl = profitAndLoss(p);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sheet(pnlRows(pl, p), { widths: [36, 16], money: [1] }), 'Profit & Loss');
    XLSX.utils.book_append_sheet(wb, entrySheet(pl.list), 'Entries');
    XLSX.writeFile(wb, `Profit and loss ${p.from} to ${p.to}.xlsx`);
}

function yearlyRows() {
    const ys = [...years()].reverse().filter(y => books.entries.some(e => yearOf(e.date) === y));
    const rows = [['Profit & loss by year'], ['', ...ys.map(String)]];
    const section = (type, title, totalLabel) => {
        rows.push([title]);
        for (const c of categoryOrder(type, books.entries)) {
            const vals = ys.map(y => sum(books.entries.filter(e => e.type === type && e.category === c && yearOf(e.date) === y)) / 100);
            if (vals.some(Boolean)) rows.push(['    ' + c, ...vals]);
        }
        rows.push([totalLabel, ...ys.map(y => sum(books.entries.filter(e => e.type === type && yearOf(e.date) === y)) / 100)], []);
    };
    section('income', 'Income', 'Total income');
    section('expense', 'Expenses', 'Total expenses');
    rows.push(['Net profit', ...ys.map(y => totals(books.entries.filter(e => yearOf(e.date) === y)).net / 100)]);
    return { rows, cols: ys.length };
}

function downloadBackup() {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, entrySheet(books.entries), 'Entries');
    const { rows, cols } = yearlyRows();
    XLSX.utils.book_append_sheet(wb, sheet(rows, { widths: [36, ...Array(cols).fill(14)], moneyFrom: 1 }), 'P&L by year');
    XLSX.utils.book_append_sheet(wb, sheet([['Type', 'Category'],
        ...books.categories.income.map(c => ['Money in', c]),
        ...books.categories.expense.map(c => ['Money out', c])], { widths: [12, 32] }), 'Categories');
    const stamp = new Date();
    XLSX.utils.book_append_sheet(wb, sheet([
        ['Fieldbook backup'],
        ['Practice', books.business],
        ['Saved', stamp.toISOString()],
        ['Entries', books.entries.length],
        [],
        ['To restore: open Fieldbook, go to The Case, choose “Restore or import…” and pick this file.'],
        ['Keep backups somewhere safe, like a cloud folder. Each one is a complete copy of your books.']
    ], { widths: [14, 60] }), 'About');
    XLSX.writeFile(wb, `Fieldbook backup ${isoOf(stamp)}.xlsx`);
    books.lastBackupAt = stamp.toISOString();
    books.changesSinceBackup = 0;
    save();
    render();
    toast('Backup downloaded. Keep it somewhere safe.');
}

$('#backup-tag').addEventListener('click', downloadBackup);
$('#download-backup').addEventListener('click', downloadBackup);

/* ───────────────────────── Import ───────────────────────── */

const HEADERS = [
    ['date', /^(date|day|when|posted|trans(action)?\.? date|posting date|invoice date|date paid|paid on)$|\bdate\b/],
    ['type', /^(type|kind|in\s*\/\s*out|income\s*\/\s*expense|direction|money in\/out|debit\/credit)$/],
    ['category', /categ|^account$|^class$|^group$|^bucket$|^line item$|^(expense|income) type$/],
    ['party', /client|customer|payee|payer|vendor|merchant|supplier|company|^who$|^name$|paid to|received from|from\s*\/\s*to/],
    ['description', /desc|memo|detail|^item|note|particular|purpose|^what|narrative|^reference|^for$/],
    ['amount', /amount|^amt|^total$|^value$|^sum$|^\$$|^usd$|^net$/],
    ['income', /income|received|revenue|deposit|credit|money in|^in$|^fees?$|receipts/],
    ['expense', /expense|expenditure|spent|spend|debit|payment|^paid$|cost|money out|^out$|withdrawal/]
];

function mapHeader(row) {
    const map = {};
    row.forEach((cell, i) => {
        const h = String(cell ?? '').trim().toLowerCase();
        if (!h) return;
        const hit = HEADERS.find(([key, re]) => !(key in map) && re.test(h));
        if (hit) map[hit[0]] = i;
    });
    return map;
}

function findTable(wb) {
    let best = null;
    const names = wb.SheetNames.includes('Entries') ? ['Entries', ...wb.SheetNames.filter(n => n !== 'Entries')] : wb.SheetNames;
    for (const name of names) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '', blankrows: false });
        for (let r = 0; r < Math.min(rows.length, 15); r++) {
            const map = mapHeader(rows[r]);
            if (!('date' in map) || !('amount' in map || 'income' in map || 'expense' in map)) continue;
            const score = Object.keys(map).length * 1000 + rows.length - r;
            if (!best || score > best.score) best = { name, header: rows[r].map(h => String(h).trim()), rows: rows.slice(r + 1), map, score };
        }
        if (best && name === 'Entries') break;
    }
    return best;
}

function typeFromText(s) {
    s = String(s ?? '').toLowerCase();
    if (/out|exp|debit|cost|spen|paid|payment|purchase|withdraw/.test(s)) return 'expense';
    if (/in|inc|rev|credit|deposit|receiv|fee|sale|invoice/.test(s)) return 'income';
    return null;
}

function parseTable(table) {
    const { rows, map } = table;
    const cell = (row, key) => key in map ? row[map[key]] : '';
    const text = (row, key) => String(cell(row, key) ?? '').trim();
    const known = { income: new Set(books.categories.income.map(c => c.toLowerCase())), expense: new Set(books.categories.expense.map(c => c.toLowerCase())) };
    const entries = [];
    let skipped = 0, guessed = 0;

    for (const row of rows) {
        const date = parseDate(cell(row, 'date'));
        let type = null, cents = NaN;
        if (!date) {
            if (row.some(v => String(v).trim())) skipped++;
            continue;
        }
        if ('amount' in map) {
            cents = parseMoney(cell(row, 'amount'));
            type = 'type' in map ? typeFromText(cell(row, 'type')) : null;
            if (!type && Number.isFinite(cents)) {
                const cat = text(row, 'category').toLowerCase();
                type = cents < 0 ? 'expense' : known.income.has(cat) ? 'income' : known.expense.has(cat) ? 'expense' : null;
                if (!type) { type = 'income'; guessed++; }
            }
        }
        if (!(Math.abs(cents) > 0)) {
            const inc = parseMoney(cell(row, 'income')), exp = parseMoney(cell(row, 'expense'));
            if (Math.abs(inc) > 0) { type = 'income'; cents = inc; }
            else if (Math.abs(exp) > 0) { type = 'expense'; cents = exp; }
        }
        if (!date || !type || !(Math.abs(cents) > 0)) {
            if (row.some(v => String(v).trim())) skipped++;
            continue;
        }
        const id = String(cell(row, 'id') ?? '');
        entries.push({
            id: id || uid(),
            date, type,
            amount: Math.abs(cents),
            category: text(row, 'category') || FALLBACK[type],
            party: text(row, 'party'),
            description: text(row, 'description'),
            createdAt: Date.now()
        });
    }
    return { entries, skipped, guessed };
}

function readBackupExtras(wb) {
    const about = wb.Sheets.About && XLSX.utils.sheet_to_json(wb.Sheets.About, { header: 1, defval: '' });
    if (!about || String(about[0]?.[0]).trim() !== 'Fieldbook backup') return null;
    const field = name => about.find(r => r[0] === name)?.[1] ?? '';
    const categories = { income: [], expense: [] };
    if (wb.Sheets.Categories) {
        for (const [type, name] of XLSX.utils.sheet_to_json(wb.Sheets.Categories, { header: 1, defval: '' }).slice(1)) {
            const t = type === 'Money in' ? 'income' : type === 'Money out' ? 'expense' : null;
            if (t && String(name).trim()) categories[t].push(String(name).trim());
        }
    }
    return { business: String(field('Practice')), saved: String(field('Saved')), categories };
}

const dupKey = e => [e.date, e.type, e.amount, e.description.toLowerCase(), e.party.toLowerCase()].join('|');

async function importFile(file) {
    let wb;
    try {
        wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    } catch (err) {
        console.error(err);
        return ask({ title: 'That file couldn’t be opened', body: `<p>Fieldbook reads Excel (.xlsx, .xls), CSV and OpenDocument files. “${esc(file.name)}” doesn’t look like one of those, or it may be damaged.</p>`, actions: [['OK', 'ok', 'btn-primary']] });
    }
    const table = findTable(wb);
    if (!table) {
        return ask({
            title: 'No entries found',
            body: `<p>Fieldbook couldn’t find a table in “${esc(file.name)}”. It looks for a header row with a <em>Date</em> column
                and an <em>Amount</em> column (or separate <em>Income</em> and <em>Expense</em> columns). <em>Description</em>, <em>Category</em>,
                <em>Client</em> and <em>Type</em> columns are picked up too.</p><p>Rename the headers in your spreadsheet to match, save, and try again.</p>`,
            actions: [['OK', 'ok', 'btn-primary']]
        });
    }
    const backup = readBackupExtras(wb);
    if (backup && table.header.includes('ID')) table.map.id = table.header.indexOf('ID'); // backups carry their IDs
    const { entries, skipped, guessed } = parseTable(table);
    if (!entries.length) {
        return ask({ title: 'No entries found', body: `<p>The sheet “${esc(table.name)}” has the right columns, but no rows with both a date and an amount.</p>`, actions: [['OK', 'ok', 'btn-primary']] });
    }
    if (backup) return restoreBackup(entries, backup, file.name);
    return importEntries(entries, { skipped, guessed, sheetName: table.name, fileName: file.name });
}

function describe(list) {
    const t = totals(list);
    const dates = list.map(e => e.date).sort();
    const nIn = list.filter(e => e.type === 'income').length;
    return `<ul>
        <li>${plural(nIn, 'entry', 'entries')} of money in, ${money(t.income)} in all</li>
        <li>${plural(list.length - nIn, 'entry', 'entries')} of money out, ${money(t.expense)} in all</li>
        <li>dated ${prettyDate(dates[0])} to ${prettyDate(dates.at(-1))}</li></ul>`;
}

async function restoreBackup(entries, backup, fileName) {
    const current = books.entries.filter(e => !e.sample).length;
    const unsaved = current && books.changesSinceBackup > 0;
    const saved = backup.saved && !isNaN(Date.parse(backup.saved)) ? ` saved ${prettyDate(isoOf(new Date(backup.saved)))}` : '';
    const choice = await ask({
        title: 'Restore this backup?',
        body: `<p>“${esc(fileName)}” is a Fieldbook backup${saved}. It holds:</p>${describe(entries)}
            ${current ? `<p>Restoring <strong>replaces</strong> the ${plural(current, 'entry', 'entries')} in Fieldbook now.</p>` : ''}
            ${unsaved ? `<p class="warn">${plural(books.changesSinceBackup, 'change')} since your last backup would be lost, so back up first to be safe.</p>` : ''}`,
        actions: [['Cancel', 'cancel'], ...(unsaved ? [['Back up, then restore', 'backup-first']] : []), [current ? 'Replace my books' : 'Restore', 'restore', unsaved ? 'btn-danger-solid' : 'btn-primary']]
    });
    if (choice === 'cancel' || !choice) return;
    if (choice === 'backup-first') downloadBackup();

    const seen = new Set();
    const unique = entries.filter(e => !seen.has(e.id) && seen.add(e.id));
    change(`Restored ${plural(unique.length, 'entry', 'entries')} from the backup.`, () => {
        books.entries = unique;
        if (backup.business) books.business = backup.business;
        for (const type of ['income', 'expense']) {
            const cats = backup.categories[type].length ? backup.categories[type] : books.categories[type];
            books.categories[type] = [...new Set([...cats, ...unique.filter(e => e.type === type).map(e => e.category)])];
        }
    }, { counts: 0 });
    // What was just restored is itself a backup
    books.lastBackupAt = new Date().toISOString();
    books.changesSinceBackup = 0;
    save();
    filters.year = 'all';
    location.hash = '#ledger';
    render();
}

async function importEntries(entries, { skipped, guessed, sheetName, fileName }) {
    const existing = new Set(books.entries.map(dupKey));
    const fresh = entries.filter(e => !existing.has(dupKey(e)));
    const dupes = entries.length - fresh.length;
    const newCats = type => [...new Set(fresh.filter(e => e.type === type).map(e => e.category))].filter(c => !books.categories[type].includes(c));
    const added = [...newCats('income'), ...newCats('expense')];
    const sampleRows = fresh.slice(0, 4).map(e => `<tr><td>${prettyDate(e.date)}</td><td>${esc(e.description || e.party || '')}</td><td>${esc(e.category)}</td><td class="num">${e.type === 'income' ? '+' : '−'}${money(e.amount)}</td></tr>`).join('');

    if (!fresh.length) {
        return ask({ title: 'Already in your books', body: `<p>All ${plural(entries.length, 'entry', 'entries')} in “${esc(fileName)}” are already in Fieldbook, so there’s nothing new to add.</p>`, actions: [['OK', 'ok', 'btn-primary']] });
    }

    const choice = await ask({
        title: `Add ${plural(fresh.length, 'entry', 'entries')}?`,
        body: `<p>From the sheet “${esc(sheetName)}” in “${esc(fileName)}”:</p>${describe(fresh)}
            <table class="table" style="margin:12px 0">${sampleRows}</table>
            ${dupes ? `<p>${plural(dupes, 'row')} already in your books will be left out.</p>` : ''}
            ${skipped ? `<p>${plural(skipped, 'row')} without a date or amount (like totals or notes) will be left out.</p>` : ''}
            ${guessed ? `<p class="warn">There’s no Type column, so ${plural(guessed, 'positive amount')} will count as money in and negative ones as money out. If that’s wrong, cancel, add a Type column (“in” or “out”), and import again.</p>` : ''}
            ${added.length ? `<p>New categories will be added: ${added.map(esc).join(', ')}.</p>` : ''}
            <p class="muted">You can undo this right after.</p>`,
        actions: [['Cancel', 'cancel'], [`Add ${plural(fresh.length, 'entry', 'entries')}`, 'add', 'btn-primary']]
    });
    if (choice !== 'add') return;

    const hadSample = books.entries.some(e => e.sample);
    change(`Added ${plural(fresh.length, 'entry', 'entries')} from “${fileName}”.`, () => {
        if (hadSample) books.entries = books.entries.filter(e => !e.sample);
        const ids = new Set(books.entries.map(e => e.id));
        for (const e of fresh) { if (ids.has(e.id)) e.id = uid(); books.entries.push(e); }
        for (const type of ['income', 'expense']) books.categories[type].push(...newCats(type));
    }, { counts: fresh.length });
    filters.year = 'all';
    location.hash = '#ledger';
    render();
}

$('#import-file').addEventListener('change', e => {
    const file = e.target.files[0];
    e.target.value = '';
    if (file) importFile(file);
});
$('label.btn:has(#import-file)')?.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('#import-file').click(); }
});

/* ───────────────────────── The Case ───────────────────────── */

function renderCase() {
    const biz = $('#business');
    if (document.activeElement !== biz) biz.value = books.business;

    const status = $('#backup-status');
    const { stale } = backupState();
    const real = books.entries.filter(e => !e.sample).length;
    status.textContent = books.lastBackupAt
        ? `Last backup ${ago(books.lastBackupAt)} (${prettyDate(isoOf(new Date(books.lastBackupAt)))}). ${books.changesSinceBackup ? plural(books.changesSinceBackup, 'change') + ' since then.' : 'No changes since.'}`
        : real ? `No backup yet, so ${real === 1 ? 'your 1 entry exists' : `your ${real.toLocaleString()} entries exist`} only in this browser.` : 'No backup yet.';
    status.classList.toggle('stale', stale);

    for (const type of ['income', 'expense']) {
        $(`#cats-${type}`).innerHTML = books.categories[type].map((c, i) => {
            const n = usage(type, c);
            return `<li><input value="${esc(c)}" data-type="${type}" data-i="${i}" aria-label="Rename ${esc(c)}">
                <span class="count">${n ? n.toLocaleString() : ''}</span>
                <button class="x" type="button" data-type="${type}" data-i="${i}" ${n ? `disabled` : ''} aria-label="Remove ${esc(c)}" title="Remove">×</button></li>`;
        }).join('');
    }
}

let bizTimer;
$('#business').addEventListener('input', e => {
    books.business = e.target.value;
    renderChrome();
    clearTimeout(bizTimer);
    bizTimer = setTimeout(() => { books.changesSinceBackup++; save(); renderChrome(); }, 500);
});

$('#view-case').addEventListener('change', e => {
    const input = e.target.closest('.cats input');
    if (!input) return;
    const { type } = input.dataset;
    const i = +input.dataset.i;
    const old = books.categories[type][i];
    const name = input.value.trim().replace(/\s+/g, ' ');
    if (!name || name === old) { input.value = old; return; }
    if (books.categories[type].some((c, j) => j !== i && c.toLowerCase() === name.toLowerCase())) {
        input.value = old;
        return toast(`There’s already a category called “${name}”.`);
    }
    const n = usage(type, old);
    change(`Renamed “${old}” to “${name}”${n ? ` on ${plural(n, 'entry', 'entries')}` : ''}.`, () => {
        books.categories[type][i] = name;
        for (const x of books.entries) if (x.type === type && x.category === old) x.category = name;
    });
});

$('#view-case').addEventListener('keydown', e => {
    if (e.key === 'Enter' && e.target.closest('.cats input')) e.target.blur();
});

$('#view-case').addEventListener('click', e => {
    const x = e.target.closest('.cats .x');
    if (!x) return;
    const { type } = x.dataset;
    const name = books.categories[type][+x.dataset.i];
    change(`Removed “${name}”.`, () => { books.categories[type].splice(+x.dataset.i, 1); });
});

for (const form of $$('.cat-add')) {
    form.addEventListener('submit', e => {
        e.preventDefault();
        const input = $('input', form);
        const { type } = form.dataset;
        const name = input.value.trim().replace(/\s+/g, ' ');
        if (!name) return input.focus();
        if (books.categories[type].some(c => c.toLowerCase() === name.toLowerCase())) return toast(`“${name}” is already there.`);
        change(`Added the category “${name}”.`, () => books.categories[type].push(name));
        input.value = '';
        input.focus();
    });
}

$('#erase-all').addEventListener('click', async () => {
    const n = books.entries.length;
    const choice = await ask({
        title: 'Erase everything?',
        body: `<p>This removes ${plural(n, 'entry', 'entries')}, your categories and your practice name from this browser.</p>
               ${books.changesSinceBackup && n ? `<p class="warn">${plural(books.changesSinceBackup, 'change')} aren’t in any backup yet.</p>` : ''}`,
        actions: [['Cancel', 'cancel'], ...(n ? [['Back up, then erase', 'backup-first']] : []), ['Erase everything', 'erase', 'btn-danger-solid']]
    });
    if (choice === 'backup-first') downloadBackup();
    if (choice === 'erase' || choice === 'backup-first') {
        change('Everything erased.', () => { books = freshBooks(); }, { counts: 0 });
    }
});

/* ───────────────────────── Sample entries ───────────────────────── */

function loadSample() {
    let seed = 7;
    const rand = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
    const pick = list => list[Math.floor(rand() * list.length)];
    const clients = ['Halvorsen Therapeutics', 'Brightwater Pharma', 'Kestrel Biosciences', 'Meridian Oncology', 'Aldine Labs'];
    const work = ['Freedom-to-operate review', 'Prior art search', 'Claims analysis', 'Patentability opinion support', 'Office action response support', 'Formulation patent landscape'];
    const list = [];
    const add = (date, type, dollars, category, party, description) => list.push({ id: uid(), date, type, amount: Math.round(dollars * 100), category, party, description, sample: true, createdAt: 0 });

    const now = new Date();
    const start = new Date(now.getFullYear() - 1, 0, 1);
    for (let d = new Date(start); d <= now; d.setMonth(d.getMonth() + 1)) {
        const y = d.getFullYear(), m = d.getMonth() + 1;
        const day = n => {
            const iso = `${y}-${pad(m)}-${pad(Math.min(n, lastDay(y, m)))}`;
            return iso > today() ? null : iso;
        };
        const on = (n, ...rest) => { const iso = day(n); if (iso) add(iso, ...rest); };
        const invoices = 2 + Math.floor(rand() * 3);
        for (let i = 0; i < invoices; i++) on(3 + Math.floor(rand() * 25), 'income', 1200 + Math.round(rand() * 34) * 250, 'Consulting fees', pick(clients), pick(work));
        on(1, 'expense', 425, 'Research & databases', 'CAS', 'SciFinder subscription');
        on(4, 'expense', 12.99, 'Software & subscriptions', 'Microsoft', 'Microsoft 365');
        on(9, 'expense', 29, 'Software & subscriptions', 'Adobe', 'Acrobat Pro');
        on(15, 'expense', 89.5, 'Phone & internet', 'Verizon', 'Phone & internet');
        on(28, 'expense', 14.2 + Math.round(rand() * 20), 'Bank & payment fees', 'Stripe', 'Card processing fees');
        on(2, 'expense', 540, 'Insurance', 'Blue Shield', 'Health insurance premium');
        on(5, 'expense', 650, 'Office & supplies', 'Harbor Workspace', 'Shared office membership');
        if (m % 2 === 0) on(19, 'expense', 900 + Math.round(rand() * 6) * 250, 'Contract help', 'R. Okafor, PhD', 'Subcontracted literature review');
        if (m === 3) on(11, 'expense', 420, 'Marketing & website', 'Squarespace', 'Website, annual plan');
        if (m % 3 === 1) on(10, 'expense', 612, 'Insurance', 'Hiscox', 'Professional liability, quarterly');
        if (m === 2) on(6, 'expense', 295, 'Professional dues & licenses', 'American Chemical Society', 'ACS membership');
        if (m === 4) on(18, 'expense', 1450, 'Legal & accounting', 'Pine & Oak CPAs', 'Tax preparation');
        if (m === 5 || m === 10) on(20, 'expense', 540, 'Continuing education', 'PLI', 'Patent law seminar');
        if (m === 6 || m === 11) { on(12, 'expense', 486.4, 'Travel', 'United Airlines', 'Client site visit'); on(13, 'expense', 68.75, 'Meals', 'Harbor Grill', 'Dinner with client team'); }
        if (rand() > 0.55) on(22, 'expense', 18 + Math.round(rand() * 90), 'Office & supplies', 'Staples', 'Printer paper & toner');
        if (rand() > 0.8) on(24, 'income', 180 + Math.round(rand() * 300), 'Expense reimbursements', pick(clients), 'Travel reimbursement');
    }
    change('Sample entries added. Clear them any time.', () => { books.entries.push(...list); }, { counts: 0 });
    filters.year = String(now.getFullYear());
}

$('#clear-sample').addEventListener('click', () => {
    change('Sample entries cleared.', () => { books.entries = books.entries.filter(e => !e.sample); }, { counts: 0 });
});

/* ───────────────────────── Dialog & toast helpers ───────────────────────── */

const askDialog = $('#ask-dialog');
function ask({ title, body, actions }) {
    $('[data-title]', askDialog).textContent = title;
    $('[data-body]', askDialog).innerHTML = body;
    $('[data-actions]', askDialog).innerHTML = '<span class="spacer"></span>' +
        actions.map(([label, value, cls = '']) => `<button class="btn ${cls}" value="${value}">${esc(label)}</button>`).join('');
    askDialog.returnValue = '';
    askDialog.showModal();
    $('[data-actions] .btn:last-child', askDialog).focus();
    return new Promise(resolve => askDialog.addEventListener('close', () => resolve(askDialog.returnValue), { once: true }));
}

let toastTimer;
function toast(message, { action, sticky } = {}) {
    const el = $('#toast');
    el.innerHTML = `<span>${esc(message)}</span>${action ? `<button class="link" type="button">${esc(action[0])}</button>` : ''}`;
    el.hidden = false;
    el.style.animation = 'none'; el.offsetHeight; el.style.animation = '';
    if (action) $('button', el).addEventListener('click', () => { el.hidden = true; action[1](); }, { once: true });
    clearTimeout(toastTimer);
    if (!sticky) toastTimer = setTimeout(() => { el.hidden = true; }, action ? 7000 : 3500);
}

/* ───────────────────────── Start ───────────────────────── */

route();
if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js');
