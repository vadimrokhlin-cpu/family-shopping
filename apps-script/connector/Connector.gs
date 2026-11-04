// מחבר Instinct לרשימת הקניות המשפחתית.
// מותקן בחשבון ה-Google של בן משפחה (לא של ודים). כל דקה מסנכרן את רשימות Google Tasks שלו
// (באותם שמות כמו באפליקציה) מול השרת המשפחתי - כך ה-Instinct שלו רואה ומשנה את אותה רשימה.

const API_URL = '{{API_URL}}';
const FAMILY_KEY = '{{FAMILY_KEY}}';

// ---------- התקנה (מריצים פעם אחת מהעורך) ----------

function setup() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sync')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(1).create();
  sync();
  Logger.log('מחובר. רשימות: ' + Object.keys(loadProp_('LISTS', {})).join(', '));
}

// ---------- סנכרון ----------

function sync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    syncOnce_();
  } finally {
    lock.releaseLock();
  }
}

function syncOnce_() {
  let state = call_([]);
  const map = loadProp_('MAP', {});        // taskId -> { item, list }
  const lists = loadProp_('LISTS', {});    // שם רשימה -> tasklistId
  const byId = {};
  state.items.forEach(it => byId[it.id] = it);

  // רשימות: כל רשימה משפחתית = רשימה באותו שם ב-Tasks
  let mine = null;
  state.lists.forEach(name => {
    if (lists[name] && tasklistExists_(lists[name])) return;
    mine = mine || listAll_(tok => Tasks.Tasklists.list({ maxResults: 100, pageToken: tok }));
    const found = mine.find(t => t.title === name);
    lists[name] = found ? found.id : Tasks.Tasklists.insert({ title: name }).id;
  });

  // שלב א: מה השתנה ב-Tasks שלי -> פעולות לשרת
  const ops = [], fresh = [], tasksByList = {};
  state.lists.forEach(name => {
    const tasks = listAll_(tok => Tasks.Tasks.list(lists[name],
      { showCompleted: true, showHidden: true, maxResults: 100, pageToken: tok }));
    tasksByList[name] = tasks;
    const seen = {};
    tasks.forEach(t => {
      seen[t.id] = true;
      const m = map[t.id];
      const done = t.status === 'completed';
      const title = clean_(t.title);
      if (m) {
        const it = byId[m.item];
        if (done) {
          // סומן כבוצע אצלי (למשל Instinct "תוריד לחם") = נקנה
          if (it && !it.checked) ops.push(op_(it, { checked: true }));
          safeDelete_(lists[name], t.id);
          delete map[t.id];
        } else if (it && !it.checked && title && title !== it.name) {
          ops.push(op_(it, { name: title }));
        }
        return;
      }
      if (done || !title) return;   // היסטוריה ישנה ב-Tasks - לא נוגעים
      // משימה חדשה (Instinct הוסיפה)
      const same = state.items.find(x => x.list === name && norm_(x.name) === norm_(title));
      const taken = same && Object.keys(map).some(k => map[k].item === same.id);
      if (taken) { safeDelete_(lists[name], t.id); return; }   // כבר ברשימה - כפילות
      if (same) {
        if (same.checked) ops.push(op_(same, { checked: false }));
        map[t.id] = { item: same.id, list: name };
      } else {
        ops.push({ type: 'upsert', id: 'k' + t.id, list: name, name: title, checked: false, base: 0 });
        fresh.push({ taskId: t.id, list: name, name: title });
      }
    });
    // משימה שנמחקה אצלי ב-Tasks = נקנה
    Object.keys(map).forEach(tid => {
      if (map[tid].list !== name || seen[tid]) return;
      const it = byId[map[tid].item];
      if (it && !it.checked) ops.push(op_(it, { checked: true }));
      delete map[tid];
    });
  });

  if (ops.length) {
    state = call_(ops);
    state.items.forEach(it => byId[it.id] = it);
  }
  // משימות חדשות: לקשר למזהה שהשרת נתן (ייתכן שאוחדו עם מוצר קיים)
  fresh.forEach(f => {
    const it = state.items.find(x => x.list === f.list && norm_(x.name) === norm_(f.name));
    if (it) map[f.taskId] = { item: it.id, list: f.list };
  });

  // שלב ב: הרשימה המשפחתית -> Tasks שלי
  const live = {};
  state.items.forEach(it => live[it.id] = it);
  const mapped = {};
  Object.keys(map).forEach(tid => {
    const m = map[tid], it = live[m.item];
    if (!it || it.checked) {
      safeDelete_(lists[m.list], tid);
      delete map[tid];
      return;
    }
    mapped[it.id] = true;
    const t = (tasksByList[m.list] || []).find(x => x.id === tid);
    if (t && clean_(t.title) !== it.name) {
      try { Tasks.Tasks.patch({ title: it.name }, lists[m.list], tid); } catch (x) {}
    }
  });
  state.items.forEach(it => {
    if (it.checked || mapped[it.id] || !lists[it.list]) return;
    const t = Tasks.Tasks.insert({ title: it.name }, lists[it.list]);
    map[t.id] = { item: it.id, list: it.list };
  });

  saveProp_('MAP', map);
  saveProp_('LISTS', lists);
}

function op_(it, change) {
  return Object.assign({ type: 'upsert', id: it.id, list: it.list, name: it.name,
                         checked: it.checked, base: it.updatedAt }, change);
}

// ---------- עזר ----------

function call_(ops) {
  const res = UrlFetchApp.fetch(API_URL, {
    method: 'post', contentType: 'text/plain;charset=utf-8', muteHttpExceptions: true,
    payload: JSON.stringify({ key: FAMILY_KEY, ops: ops })
  });
  const data = JSON.parse(res.getContentText());
  if (data.error) throw new Error('server: ' + data.error);
  return data;
}

function tasklistExists_(id) {
  try { Tasks.Tasklists.get(id); return true; } catch (x) { return false; }
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

function loadProp_(k, d) {
  const v = PropertiesService.getScriptProperties().getProperty(k);
  return v ? JSON.parse(v) : d;
}
function saveProp_(k, v) { PropertiesService.getScriptProperties().setProperty(k, JSON.stringify(v)); }
function clean_(s) { return String(s || '').replace(/\s+/g, ' ').trim(); }
function norm_(s) { return clean_(s).toLowerCase(); }
