// Simulation: server (Vadim's Tasks) + connector (Lena's Tasks) + the app, with shared clock
const fs = require('fs'), vm = require('vm'), path = require('path'), assert = require('assert');
const read = f => fs.readFileSync(path.join(__dirname, '..', 'apps-script', f), 'utf8');

let clock = 1_800_000_000_000;
const tick = (ms = 1000) => { clock += ms; };
class FakeDate extends Date { constructor(...a) { a.length ? super(...a) : super(clock); } static now() { return clock; } }

function fakeTasks(name) {
  const lists = {}; let n = 0;
  const api = {
    lists,
    Tasklists: {
      list: () => ({ items: Object.entries(lists).map(([id, l]) => ({ id, title: l.title })) }),
      get: id => { if (!lists[id]) throw new Error('404 Not Found'); return { id }; },
      insert: ({ title }) => { const id = name + 'L' + (++n); lists[id] = { title, tasks: {} }; return { id }; } },
    Tasks: {
      list(lid, o) { const l = lists[lid]; if (!l) throw new Error('404 Not Found');
        return { items: Object.values(l.tasks)
          .filter(t => (o.showDeleted || !t.deleted) && (!o.updatedMin || new Date(t.updated) >= new Date(o.updatedMin)))
          .map(t => ({ ...t })) }; },
      insert(t, lid) { const id = name + 'T' + (++n); lists[lid].tasks[id] = { id, title: t.title, status: 'needsAction', updated: new Date(clock).toISOString() }; return { id }; },
      patch(t, lid, id) { const x = lists[lid].tasks[id]; if (!x || x.deleted) throw new Error('404 Not Found'); x.title = t.title; x.updated = new Date(clock).toISOString(); },
      remove(lid, id) { const x = lists[lid].tasks[id]; if (x) { x.deleted = true; x.updated = new Date(clock).toISOString(); } }
    }
  };
  api.open = title => { const l = Object.values(lists).find(l => l.title === title);
    return l ? Object.values(l.tasks).filter(t => !t.deleted && t.status !== 'completed').map(t => t.title).sort() : []; };
  api.add = (title, item) => { const l = Object.values(lists).find(l => l.title === title); const id = name + 'X' + (++n);
    l.tasks[id] = { id, title: item, status: 'needsAction', updated: new Date(clock).toISOString() }; };
  api.complete = (title, item) => { const l = Object.values(lists).find(l => l.title === title);
    const t = Object.values(l.tasks).find(t => t.title === item && !t.deleted && t.status !== 'completed');
    t.status = 'completed'; t.updated = new Date(clock).toISOString(); };
  api.del = (title, item) => { const l = Object.values(lists).find(l => l.title === title);
    const t = Object.values(l.tasks).find(t => t.title === item && !t.deleted); t.deleted = true; t.updated = new Date(clock).toISOString(); };
  return api;
}

function propsStore() { const p = {}; return { p, svc: { getScriptProperties: () => ({ getProperty: k => p[k] ?? null, setProperty: (k, v) => p[k] = v }) } }; }
const common = {
  Date: FakeDate, console,
  LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} }) },
  Logger: { log() {} },
  ScriptApp: { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ everyMinutes: () => ({ create() {} }) }) }) },
  ContentService: { createTextOutput: s => ({ setMimeType() { return s; } }), MimeType: {} },
  Utilities: { getUuid: () => 'abcdef12-3456-7890-abcd-ef1234567890' }
};

// --- server (Vadim) ---
const sheets = {};
function mkSheet(n) { const rows = []; return sheets[n] = { rows, getLastRow: () => rows.length,
  getRange(r, c, k, w) { if (typeof r === 'string') return { setNumberFormat() {} };
    return { setValues(v) { v.forEach((row, i) => rows[r - 1 + i] = row.slice()); return { setFontWeight() {} }; },
             getValues() { return rows.slice(r - 1, r - 1 + k).map(x => x.slice()); }, clearContent() { rows.length = r - 1; } }; },
  setFrozenRows() {} }; }
const vT = fakeTasks('v'), vP = propsStore();
const server = vm.createContext({ ...common, Tasks: vT, PropertiesService: vP.svc,
  SpreadsheetApp: { getActive: () => ({ getSheetByName: n => sheets[n], insertSheet: mkSheet }) } });
vm.runInContext(read('Code.gs'), server);
vT.Tasklists.insert({ title: 'קניות' });
vT.add('קניות', 'חלב');
server.setup();
const KEY = vP.p.FAMILY_KEY;
const post = ops => JSON.parse(server.doPost({ parameter: {}, postData: { contents: JSON.stringify({ key: KEY, ops }) } }));

// --- connector (Lena) ---
const lT = fakeTasks('l'), lP = propsStore();
const conn = vm.createContext({ ...common, Tasks: lT, PropertiesService: lP.svc,
  UrlFetchApp: { fetch: (u, o) => { tick(10); return { getContentText: () => server.doPost({ parameter: {}, postData: { contents: o.payload } }) }; } } });
vm.runInContext(read('connector/Connector.gs').replace('{{API_URL}}', 'x').replace('{{FAMILY_KEY}}', KEY), conn);

// Lena already has a "קניות" list with her own item + an old completed one
lT.Tasklists.insert({ title: 'קניות' });
lT.add('קניות', 'ביצים'); lT.add('קניות', 'ישן'); lT.complete('קניות', 'ישן');
const round = () => { tick(30000); conn.sync(); tick(30000); server.syncTasks(); };

conn.setup(); round();
assert.deepStrictEqual(lT.open('קניות'), ['ביצים', 'חלב'], 'lena sees family list');
assert.deepStrictEqual(vT.open('קניות'), ['ביצים', 'חלב'], 'vadim sees lena item');

// Lena's Instinct adds bread → Vadim + app
tick(); lT.add('קניות', 'לחם'); round();
assert.deepStrictEqual(vT.open('קניות'), ['ביצים', 'חלב', 'לחם']);
let s = post([]);
assert.ok(s.items.find(i => i.name === 'לחם' && !i.checked));

// Vadim's Instinct completes milk → disappears at Lena
tick(); vT.complete('קניות', 'חלב'); round(); round();
assert.deepStrictEqual(lT.open('קניות'), ['ביצים', 'לחם']);

// Lena's Instinct deletes eggs → bought everywhere
tick(); lT.del('קניות', 'ביצים'); round();
assert.deepStrictEqual(vT.open('קניות'), ['לחם']);
assert.ok(post([]).items.find(i => i.name === 'ביצים').checked);

// app checks bread → gone at Lena
s = post([]); const br = s.items.find(i => i.name === 'לחם');
tick(); post([{ type: 'upsert', id: br.id, list: 'קניות', name: 'לחם', checked: true, base: br.updatedAt }]); round();
assert.deepStrictEqual(lT.open('קניות'), []);

// Lena re-adds milk (history item) → reopened, no duplicate
tick(); lT.add('קניות', 'חלב'); round();
s = post([]);
assert.strictEqual(s.items.filter(i => i.name === 'חלב').length, 1);
assert.deepStrictEqual(vT.open('קניות'), ['חלב']);
assert.deepStrictEqual(lT.open('קניות'), ['חלב']);

// both Instincts add same item at once → one item, one task each
tick(); lT.add('קניות', 'עגבניות'); vT.add('קניות', 'עגבניות'); round(); round();
assert.deepStrictEqual(lT.open('קניות'), ['חלב', 'עגבניות']);
assert.deepStrictEqual(vT.open('קניות'), ['חלב', 'עגבניות']);
assert.strictEqual(post([]).items.filter(i => i.name === 'עגבניות').length, 1);

// new list in app → appears at Lena
tick(); post([{ type: 'addList', name: 'פארם' }, { type: 'upsert', id: 'p1', list: 'פארם', name: 'שמפו', checked: false, base: 0 }]); round();
assert.deepStrictEqual(lT.open('פארם'), ['שמפו']);

// stable over many rounds; old completed task untouched
for (let i = 0; i < 4; i++) round();
assert.deepStrictEqual(lT.open('קניות'), ['חלב', 'עגבניות']);
assert.ok(Object.values(Object.values(lT.lists)[0].tasks).find(t => t.title === 'ישן' && !t.deleted));
console.log('CONNECTOR OK');
