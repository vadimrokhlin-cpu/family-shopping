// רשימת קניות משפחתית - צד שרת.
// סקריפט צמוד לגיליון Google: הגיליון הוא מסד הנתונים, הסקריפט הוא ה-API של האפליקציה
// וגם מסנכרן כל דקה מול Google Tasks (שם Instinct מוסיפה ומורידה מוצרים).
// כל רשימה באפליקציה = רשימה באותו שם ב-Google Tasks. ב-Tasks נשארים רק מוצרים שעוד צריך לקנות.

const ITEMS = 'Items', LISTS = 'Lists';
const ICOLS = ['id', 'list', 'name', 'checked', 'updatedAt', 'taskId', 'dirty', 'deleted', 'count'];
const LCOLS = ['name', 'tasklistId', 'createdAt'];
const DEFAULT_LIST = 'קניות';

// ---------- התקנה (מריצים פעם אחת מהעורך) ----------

function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, ITEMS, ICOLS);
  ensureSheet_(ss, LISTS, LCOLS);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('FAMILY_KEY')) {
    props.setProperty('FAMILY_KEY', Utilities.getUuid().replace(/-/g, '').slice(0, 20));
  }
  const db = load_();
  if (!db.lists.length) db.lists.push({ name: DEFAULT_LIST, tasklistId: '', createdAt: Date.now() });
  syncDb_(db);
  save_(db);
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncTasks')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncTasks').timeBased().everyMinutes(1).create();
  Logger.log('FAMILY_KEY = ' + props.getProperty('FAMILY_KEY'));
}

function ensureSheet_(ss, name, cols) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('A:C').setNumberFormat('@');   // שמות מוצרים כטקסט, ש-"1/2" לא יהפוך לתאריך
  sh.getRange('F:F').setNumberFormat('@');
}

// ---------- טריגר כל דקה ----------

function syncTasks() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const db = load_();
    syncDb_(db);
    save_(db);
  } finally {
    lock.releaseLock();
  }
}

// ---------- API לאפליקציה ----------

function doGet(e) { return handle_(e.parameter || {}, null); }

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (x) {}
  return handle_(e.parameter || {}, body);
}

function handle_(p, body) {
  const key = (body && body.key) || p.key;
  if (!key || key !== PropertiesService.getScriptProperties().getProperty('FAMILY_KEY')) {
    return json_({ error: 'unauthorized' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const db = load_();
    const ops = (body && body.ops) || [];
    if (ops.length) applyOps_(db, ops);
    if (ops.length || syncDue_()) {
      try { syncDb_(db); } catch (x) { console.warn('sync failed: ' + x); }
      save_(db);
    }
    return json_(state_(db));
  } finally {
    lock.releaseLock();
  }
}

function syncDue_() {
  const last = PropertiesService.getScriptProperties().getProperty('LAST_SYNC');
  return !last || Date.now() - new Date(last).getTime() > 20000;
}

function state_(db) {
  return {
    lists: db.lists.slice().sort((a, b) => a.createdAt - b.createdAt).map(l => l.name),
    items: db.items.filter(it => !truthy_(it.deleted)).map(it => ({
      id: it.id, list: it.list, name: it.name, checked: truthy_(it.checked),
      updatedAt: Number(it.updatedAt) || 0, count: Number(it.count) || 0
    }))
  };
}

// שינוי מהאפליקציה נדחה אם הפריט השתנה בשרת אחרי שהטלפון ראה אותו (base) - השרת מנצח.
function applyOps_(db, ops) {
  const byId = {};
  db.items.forEach(it => byId[it.id] = it);
  const now = Date.now();
  ops.forEach(op => {
    if (op.type === 'addList') {
      const name = clean_(op.name);
      if (name && !db.lists.some(l => l.name === name)) {
        db.lists.push({ name: name, tasklistId: '', createdAt: now });
      }
      return;
    }
    let it = byId[op.id];
    if (!it && op.type === 'upsert') {
      // אותו מוצר נוסף משני טלפונים במקביל - מאחדים לפי שם
      it = db.items.find(x => x.list === op.list && norm_(x.name) === norm_(op.name) && !truthy_(x.deleted));
    } else if (it && Number(op.base) < Number(it.updatedAt)) {
      return;
    }
    if (op.type === 'delete') {
      if (it) { it.deleted = true; it.dirty = true; it.updatedAt = now; }
      return;
    }
    const name = clean_(op.name);
    if (!name) return;
    if (!it) {
      it = { id: op.id, list: op.list, name: name, checked: false, updatedAt: 0,
             taskId: '', dirty: true, deleted: false, count: 0 };
      db.items.push(it);
      byId[it.id] = it;
    }
    if (op.checked && !truthy_(it.checked)) it.count = (Number(it.count) || 0) + 1;
    it.name = name;
    it.checked = !!op.checked;
    it.updatedAt = now;
    it.dirty = true;
  });
}

// ---------- סנכרון מול Google Tasks ----------

function syncDb_(db) {
  const props = PropertiesService.getScriptProperties();
  const started = new Date();
  const last = props.getProperty('LAST_SYNC');
  const byTask = {};
  db.items.forEach(it => { if (it.taskId) byTask[it.taskId] = it; });

  // רשימה חדשה: מתחברים לרשימת Tasks באותו שם, ואם אין - יוצרים
  let tasklists = null;
  const fresh = {};
  db.lists.forEach(l => {
    if (l.tasklistId) return;
    tasklists = tasklists || listAll_(tok => Tasks.Tasklists.list({ maxResults: 100, pageToken: tok }));
    const found = tasklists.find(t => t.title === l.name);
    l.tasklistId = found ? found.id : Tasks.Tasklists.insert({ title: l.name }).id;
    fresh[l.name] = true;
  });

  // משיכה: מה השתנה ב-Tasks מאז הסנכרון הקודם (עם מרווח ביטחון של 2 דקות)
  db.lists.forEach(l => {
    const opts = { showCompleted: true, showHidden: true, showDeleted: true, maxResults: 100 };
    if (last && !fresh[l.name]) opts.updatedMin = new Date(new Date(last).getTime() - 120000).toISOString();
    try {
      listAll_(tok => Tasks.Tasks.list(l.tasklistId, Object.assign({ pageToken: tok }, opts)))
        .forEach(t => pullTask_(db, l, t, byTask));
    } catch (x) {
      // הרשימה נמחקה ב-Tasks - בפעם הבאה תיווצר מחדש
      if (String(x).indexOf('Not Found') >= 0 || String(x).indexOf('404') >= 0) l.tasklistId = '';
      console.warn('pull ' + l.name + ': ' + x);
    }
  });

  // דחיפה: שינויים מהאפליקציה
  db.items.forEach(it => { if (truthy_(it.dirty)) pushItem_(db, it, byTask); });
  db.items = db.items.filter(it => !(truthy_(it.deleted) && !it.taskId && !truthy_(it.dirty)));
  props.setProperty('LAST_SYNC', started.toISOString());
}

function pullTask_(db, l, t, byTask) {
  const it = byTask[t.id];
  const tUpd = new Date(t.updated).getTime();
  const gone = t.deleted || t.status === 'completed';

  if (it) {
    if (gone) {
      // סומן כבוצע / נמחק ב-Tasks (למשל Instinct "תוריד חלב") = נקנה
      if (!t.deleted) safeDelete_(l.tasklistId, t.id);
      if (!truthy_(it.checked) && !truthy_(it.deleted)) {
        it.checked = true;
        it.count = (Number(it.count) || 0) + 1;
        it.updatedAt = Date.now();
      }
      it.taskId = '';
      delete byTask[t.id];
      return;
    }
    if (truthy_(it.checked) || truthy_(it.deleted)) {
      if (Number(it.updatedAt) >= tUpd) return;   // השינוי באפליקציה חדש יותר - הדחיפה תמחק את המשימה
      it.checked = false;
      it.deleted = false;
      it.dirty = false;
      it.updatedAt = Date.now();
    }
    const title = clean_(t.title);
    if (title && title !== it.name && tUpd > Number(it.updatedAt)) {
      it.name = title;
      it.updatedAt = Date.now();
    }
    return;
  }

  // משימה שאינה מקושרת. היסטוריה ישנה שבוצעה ב-Tasks - לא נוגעים.
  if (gone) return;
  const name = clean_(t.title);
  if (!name) return;
  let match = db.items.find(x => x.list === l.name && norm_(x.name) === norm_(name) && !truthy_(x.deleted));
  if (match && match.taskId && match.taskId !== t.id) {
    safeDelete_(l.tasklistId, t.id);   // המוצר כבר ברשימה - כפילות
    return;
  }
  if (!match) {
    match = { id: 't' + t.id, list: l.name, name: name, checked: false, updatedAt: 0,
              taskId: '', dirty: false, deleted: false, count: 0 };
    db.items.push(match);
  }
  match.checked = false;
  match.deleted = false;
  match.dirty = false;
  match.taskId = t.id;
  match.updatedAt = Date.now();
  byTask[t.id] = match;
}

function pushItem_(db, it, byTask) {
  const l = db.lists.find(x => x.name === it.list);
  if (!l || !l.tasklistId) return;
  try {
    if (truthy_(it.deleted) || truthy_(it.checked)) {
      if (it.taskId) {
        safeDelete_(l.tasklistId, it.taskId);
        delete byTask[it.taskId];
        it.taskId = '';
      }
    } else if (!it.taskId) {
      const t = Tasks.Tasks.insert({ title: it.name }, l.tasklistId);
      it.taskId = t.id;
      byTask[t.id] = it;
    } else {
      Tasks.Tasks.patch({ title: it.name }, l.tasklistId, it.taskId);
    }
    it.dirty = false;
  } catch (x) {
    // המשימה נמחקה בינתיים - ניצור אותה מחדש בסיבוב הבא
    if (String(x).indexOf('Not Found') >= 0 || String(x).indexOf('404') >= 0) it.taskId = '';
    console.warn('push ' + it.name + ': ' + x);
  }
}

function safeDelete_(tasklistId, taskId) {
  try { Tasks.Tasks.remove(tasklistId, taskId); } catch (x) {}
}

function listAll_(fetchPage) {
  let out = [], tok;
  do {
    const res = fetchPage(tok);
    out = out.concat(res.items || []);
    tok = res.nextPageToken;
  } while (tok);
  return out;
}

// ---------- גיליון ----------

function load_() {
  const ss = SpreadsheetApp.getActive();
  return {
    items: readTable_(ss.getSheetByName(ITEMS), ICOLS),
    lists: readTable_(ss.getSheetByName(LISTS), LCOLS)
  };
}

function save_(db) {
  writeTable_(ITEMS, ICOLS, db.items);
  writeTable_(LISTS, LCOLS, db.lists);
}

function readTable_(sh, cols) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getValues().map(r => {
    const o = {};
    cols.forEach((c, i) => o[c] = r[i]);
    return o;
  });
}

function writeTable_(name, cols, rows) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const old = sh.getLastRow() - 1;
  if (old > 0) sh.getRange(2, 1, old, cols.length).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, cols.length)
      .setValues(rows.map(o => cols.map(c => o[c] === undefined || o[c] === null ? '' : o[c])));
  }
}

// ---------- עזר ----------

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function truthy_(v) { return v === true || v === 'TRUE' || v === 'true'; }
function clean_(s) { return String(s || '').replace(/\s+/g, ' ').trim(); }
function norm_(s) { return clean_(s).toLowerCase(); }
