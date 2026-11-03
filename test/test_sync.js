// Simulates Sheets + Tasks + Properties to test Code.gs logic
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const src = fs.readFileSync(require('path').join(__dirname, '../apps-script/Code.gs'), 'utf8');

let clock = 1_800_000_000_000;
const tick = (ms = 1000) => { clock += ms; };
class FakeDate extends Date { constructor(...a) { a.length ? super(...a) : super(clock); } static now() { return clock; } }

const sheets = {};
function mkSheet(name) {
  const rows = [];
  return sheets[name] = { rows, getLastRow: () => rows.length,
    getRange(r, c, n, w) {
      if (typeof r === 'string') return { setNumberFormat() {} };
      return {
        setValues(v) { v.forEach((row, i) => rows[r - 1 + i] = row.slice()); return { setFontWeight() {} }; },
        getValues() { return rows.slice(r - 1, r - 1 + n).map(x => x.slice()); },
        clearContent() { rows.length = r - 1; } };
    }, setFrozenRows() {} };
}
const props = {};
const tasklists = {}; let tid = 0;
const Tasks = {
  Tasklists: { list: () => ({ items: Object.entries(tasklists).map(([id, l]) => ({ id, title: l.title })) }),
               insert: ({ title }) => { const id = 'L' + (++tid); tasklists[id] = { title, tasks: {} }; return { id }; } },
  Tasks: {
    list(lid, o) { const l = tasklists[lid]; if (!l) throw new Error('404 Not Found');
      return { items: Object.values(l.tasks).filter(t => !o.updatedMin || new Date(t.updated) >= new Date(o.updatedMin)).map(t => ({ ...t })) }; },
    insert(t, lid) { const id = 'T' + (++tid); tasklists[lid].tasks[id] = { id, title: t.title, status: 'needsAction', updated: new Date(clock).toISOString() }; return { id }; },
    patch(t, lid, id) { const x = tasklists[lid].tasks[id]; if (!x || x.deleted) throw new Error('404 Not Found'); x.title = t.title; x.updated = new Date(clock).toISOString(); },
    remove(lid, id) { const x = tasklists[lid].tasks[id]; if (x) { x.deleted = true; x.updated = new Date(clock).toISOString(); } }
  }
};
const ss = { getSheetByName: n => sheets[n], insertSheet: n => mkSheet(n) };
const ctx = {
  Date: FakeDate, console, Tasks,
  SpreadsheetApp: { getActive: () => ss },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => props[k] = v }) },
  LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} }) },
  Utilities: { getUuid: () => 'abcdef12-3456-7890-abcd-ef1234567890' },
  Logger: { log() {} },
  ScriptApp: { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ everyMinutes: () => ({ create() {} }) }) }) },
  ContentService: { createTextOutput: s => ({ s, setMimeType() { return JSON.parse(s); } }), MimeType: {} }
};
vm.createContext(ctx);
vm.runInContext(src, ctx);

// Instinct already has a Tasks list "קניות" with an open + completed item
tasklists.L0 = { title: 'קניות', tasks: {
  a: { id: 'a', title: 'חלב', status: 'needsAction', updated: new Date(clock).toISOString() },
  b: { id: 'b', title: 'ישן', status: 'completed', updated: new Date(clock).toISOString() } } };
tick();
ctx.setup();
const key = props.FAMILY_KEY;
const api = ops => { tick(25000); return ctx.doPost({ parameter: {}, postData: { contents: JSON.stringify({ key, ops }) } }); };
const openTasks = () => Object.values(tasklists.L0.tasks).filter(t => !t.deleted && t.status !== 'completed').map(t => t.title).sort();
const item = (s, n) => s.items.find(i => i.name === n);

let s = api([]);
assert.deepStrictEqual(s.lists, ['קניות']);
assert.ok(item(s, 'חלב') && !item(s, 'חלב').checked, 'open task imported');
assert.ok(!item(s, 'ישן'), 'completed history not imported');
assert.ok(tasklists.L0.tasks.b && !tasklists.L0.tasks.b.deleted, 'old completed untouched');

// family adds bread in the app → appears in Tasks
s = api([{ type: 'upsert', id: 'c1', list: 'קניות', name: 'לחם', checked: false, base: 0 }]);
assert.deepStrictEqual(openTasks(), ['חלב', 'לחם']);

// Instinct adds eggs, completes milk
tick(5000);
const L = tasklists.L0.tasks;
L.e = { id: 'e', title: 'ביצים', status: 'needsAction', updated: new Date(clock).toISOString() };
L.a.status = 'completed'; L.a.updated = new Date(clock).toISOString();
tick(60000); ctx.syncTasks();
s = api([]);
assert.ok(item(s, 'ביצים') && !item(s, 'ביצים').checked, 'instinct add');
assert.ok(item(s, 'חלב').checked, 'instinct complete -> checked');
assert.ok(L.a.deleted, 'completed task cleaned');
assert.deepStrictEqual(openTasks(), ['ביצים', 'לחם']);

// in store: check bread (base = updatedAt seen)
const bread = item(s, 'לחם');
s = api([{ type: 'upsert', id: bread.id, list: 'קניות', name: 'לחם', checked: true, base: bread.updatedAt }]);
assert.ok(item(s, 'לחם').checked);
assert.deepStrictEqual(openTasks(), ['ביצים']);

// Instinct re-adds milk → existing history item reopened, not duplicated
tick(5000);
L.m2 = { id: 'm2', title: 'חלב', status: 'needsAction', updated: new Date(clock).toISOString() };
tick(60000); ctx.syncTasks();
s = api([]);
assert.strictEqual(s.items.filter(i => i.name === 'חלב').length, 1);
assert.ok(!item(s, 'חלב').checked, 'milk reopened');

// stale offline op (based on old version) is rejected
s = api([{ type: 'upsert', id: item(s, 'חלב').id, list: 'קניות', name: 'חלב', checked: true, base: 1 }]);
assert.ok(!item(s, 'חלב').checked, 'stale op ignored');

// Repeated syncs are idempotent
for (let i = 0; i < 3; i++) { tick(60000); ctx.syncTasks(); }
s = api([]);
assert.deepStrictEqual(openTasks(), ['ביצים', 'חלב']);
assert.strictEqual(s.items.length, 3);

// two phones add same name offline → merged
s = api([{ type: 'upsert', id: 'x1', list: 'קניות', name: 'עגבניות', checked: false, base: 0 },
         { type: 'upsert', id: 'x2', list: 'קניות', name: 'עגבניות ', checked: false, base: 0 }]);
assert.strictEqual(s.items.filter(i => i.name === 'עגבניות').length, 1);
assert.deepStrictEqual(openTasks(), ['ביצים', 'חלב', 'עגבניות']);

// new list → new Tasks list
s = api([{ type: 'addList', name: 'פארם' }, { type: 'upsert', id: 'p1', list: 'פארם', name: 'משחת שיניים', checked: false, base: 0 }]);
const pharm = Object.values(tasklists).find(l => l.title === 'פארם');
assert.ok(pharm && Object.values(pharm.tasks).some(t => t.title === 'משחת שיניים'));

// delete from history while open → task removed, row gone
const tom = item(s, 'עגבניות');
s = api([{ type: 'delete', id: tom.id, base: tom.updatedAt }]);
assert.ok(!item(s, 'עגבניות'));
assert.deepStrictEqual(openTasks(), ['ביצים', 'חלב']);

// Instinct deletes a task outright → treated as bought
tick(5000); L.e.deleted = true; L.e.updated = new Date(clock).toISOString();
tick(60000); ctx.syncTasks();
s = api([]);
assert.ok(item(s, 'ביצים').checked);

// name like 1/2 stays text (format set) — just check value roundtrip
console.log('ALL OK', JSON.stringify(s.items.map(i => [i.name, i.checked, i.count])));


