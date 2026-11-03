// רשימת קניות משפחתית - האפליקציה בטלפון.
// הכל נשמר קודם בטלפון (עובד בלי אינטרנט), ושינויים ממתינים נשלחים לשרת כשיש קליטה.

const K = k => 'fs_' + k;
const load = (k, d) => { try { return JSON.parse(localStorage.getItem(K(k))) ?? d; } catch { return d; } };
const store = (k, v) => localStorage.setItem(K(k), JSON.stringify(v));
const $ = id => document.getElementById(id);
const norm = s => String(s).replace(/\s+/g, ' ').trim().toLowerCase();

// קישור משפחתי: ...#api=<כתובת הסקריפט>&key=<מפתח>
(function readLink() {
  const h = new URLSearchParams(location.hash.slice(1));
  if (h.get('api') && h.get('key')) {
    store('cfg', { api: h.get('api'), key: h.get('key') });
    history.replaceState(null, '', location.pathname);
  }
})();

let cfg = load('cfg', null);
let db = load('db', { lists: ['קניות'], items: [] });
let pending = load('pending', {});          // מזהה -> הפעולה האחרונה שעוד לא נשלחה
let current = load('list', db.lists[0]);
let syncing = false, online = null;

function persist() { store('db', db); store('pending', pending); store('list', current); }

// ---------- פעולות ----------

function queue(it, type = 'upsert') {
  pending[it.id] = { type, id: it.id, list: it.list, name: it.name, checked: it.checked, base: it.updatedAt };
}

function addItem(name) {
  name = name.replace(/\s+/g, ' ').trim();
  if (!name) return;
  let it = db.items.find(x => x.list === current && norm(x.name) === norm(name));
  if (it) {
    if (!it.checked) return flash(it.id);
    it.checked = false;
  } else {
    it = { id: 'c' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
           list: current, name, checked: false, updatedAt: 0, count: 0 };
    db.items.push(it);
  }
  queue(it);
  changed();
}

function toggle(id) {
  const it = db.items.find(x => x.id === id);
  if (!it) return;
  it.checked = !it.checked;
  if (it.checked) it.count = (it.count || 0) + 1;
  queue(it);
  changed();
}

function removeItem(id) {
  const it = db.items.find(x => x.id === id);
  if (!it || !confirm(`למחוק את "${it.name}" מההיסטוריה?`)) return;
  db.items = db.items.filter(x => x !== it);
  queue(it, 'delete');
  changed();
}

function addList(name) {
  name = (name || '').replace(/\s+/g, ' ').trim();
  if (!name || db.lists.includes(name)) return;
  db.lists.push(name);
  pending['list:' + name] = { type: 'addList', name };
  current = name;
  changed();
}

function importText(text) {
  let n = 0;
  text.split(/\r?\n/).forEach(line => {
    const done = /^\s*(\[x\]|☑|✔|✓|☒)/i.test(line);
    const name = line.replace(/^\s*(\[[ x]?\]|☐|☑|✔|✓|☒|[-*•·])\s*/i, '').trim();
    if (!name) return;
    let it = db.items.find(x => x.list === current && norm(x.name) === norm(name));
    if (it) return;
    it = { id: 'c' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6) + n,
           list: current, name, checked: done, updatedAt: 0, count: 0 };
    db.items.push(it);
    queue(it);
    n++;
  });
  changed();
  alert(`יובאו ${n} מוצרים`);
}

function changed() { persist(); render(); sync(); }

// ---------- סנכרון ----------

function applyLocal(op) {
  if (op.type === 'addList') { if (!db.lists.includes(op.name)) db.lists.push(op.name); return; }
  const i = db.items.findIndex(x => x.id === op.id);
  if (op.type === 'delete') { if (i >= 0) db.items.splice(i, 1); return; }
  if (i >= 0) Object.assign(db.items[i], { name: op.name, checked: op.checked });
  else db.items.push({ id: op.id, list: op.list, name: op.name, checked: op.checked, updatedAt: 0, count: 0 });
}

async function sync() {
  if (!cfg || syncing) return;
  syncing = true;
  const sent = { ...pending };
  try {
    const res = await fetch(cfg.api, {
      method: 'POST',
      body: JSON.stringify({ key: cfg.key, ops: Object.values(sent) })
    });
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    for (const k in sent) if (pending[k] === sent[k]) delete pending[k];
    db = { lists: data.lists, items: data.items };
    Object.values(pending).forEach(applyLocal);   // מה שנוסף בזמן שהבקשה הייתה באוויר
    if (!db.lists.includes(current)) current = db.lists[0];
    online = true;
    persist();
    render();
  } catch (e) {
    online = false;
    renderStatus();
  } finally {
    syncing = false;
  }
}

// ---------- תצוגה ----------

function render() {
  $('setup').hidden = !!cfg;
  $('app').hidden = !cfg;
  if (!cfg) return;

  const tabs = $('tabs');
  tabs.innerHTML = '';
  db.lists.forEach(name => {
    const b = document.createElement('button');
    const n = db.items.filter(x => x.list === name && !x.checked).length;
    b.innerHTML = `${esc(name)} <span class="count">${n || ''}</span>`;
    b.className = name === current ? 'active' : '';
    b.onclick = () => { current = name; persist(); render(); };
    tabs.appendChild(b);
  });

  const mine = db.items.filter(x => x.list === current);
  const open = mine.filter(x => !x.checked).sort((a, b) => a.name.localeCompare(b.name, 'he'));
  const done = mine.filter(x => x.checked).sort((a, b) => (b.count || 0) - (a.count || 0) || a.name.localeCompare(b.name, 'he'));

  $('open').innerHTML = open.map(it =>
    `<li data-id="${it.id}"><span class="box"></span><span class="name">${esc(it.name)}</span></li>`).join('');
  $('empty').hidden = open.length > 0;
  $('doneTitle').textContent = `נקנו (${done.length})`;
  $('done').innerHTML = done.map(it =>
    `<li data-id="${it.id}"><span class="box">✓</span><span class="name">${esc(it.name)}</span><button class="del" data-del="${it.id}">✕</button></li>`).join('');
  renderSuggest();
  renderStatus();
}

function renderSuggest() {
  const q = norm($('addInput').value);
  const ul = $('suggest');
  if (!q) { ul.innerHTML = ''; return; }
  const hits = db.items
    .filter(x => x.list === current && x.checked && norm(x.name).includes(q))
    .sort((a, b) => norm(b.name).startsWith(q) - norm(a.name).startsWith(q) || (b.count || 0) - (a.count || 0))
    .slice(0, 6);
  ul.innerHTML = hits.map(it => `<li data-id="${it.id}">${esc(it.name)}</li>`).join('');
}

function renderStatus() {
  const s = $('status');
  const n = Object.keys(pending).length;
  if (online === false) { s.className = 'offline'; s.textContent = n ? `אין אינטרנט · ${n} שינויים יישלחו כשתחזור קליטה` : 'אין אינטרנט · עובד מהטלפון'; }
  else if (online) { s.className = 'ok'; s.textContent = n ? 'שולח...' : 'מסונכרן'; }
  else { s.className = ''; s.textContent = 'מתחבר...'; }
}

function flash(id) {
  const li = document.querySelector(`#open li[data-id="${id}"]`);
  if (li) { li.scrollIntoView({ block: 'center' }); li.style.background = '#fff59d'; setTimeout(() => li.style.background = '', 800); }
}

function esc(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

// ---------- אירועים ----------

$('addForm').onsubmit = e => {
  e.preventDefault();
  addItem($('addInput').value);
  $('addInput').value = '';
  renderSuggest();
};
$('addInput').oninput = renderSuggest;
$('suggest').onclick = e => {
  const li = e.target.closest('li');
  if (!li) return;
  toggle(li.dataset.id);
  $('addInput').value = '';
  renderSuggest();
};
$('open').onclick = e => {
  const li = e.target.closest('li');
  if (!li) return;
  li.classList.add('bye');
  setTimeout(() => toggle(li.dataset.id), 200);
};
$('done').onclick = e => {
  if (e.target.dataset.del) return removeItem(e.target.dataset.del);
  const li = e.target.closest('li');
  if (li) toggle(li.dataset.id);
};

$('menuBtn').onclick = () => $('menu').showModal();
$('closeMenu').onclick = () => $('menu').close();
$('newListBtn').onclick = () => { $('menu').close(); addList(prompt('שם הרשימה החדשה (למשל: פארם)')); };
$('importBtn').onclick = () => { importText($('importText').value); $('importText').value = ''; $('menu').close(); };
$('shareBtn').onclick = async () => {
  const url = `${location.origin}${location.pathname}#api=${encodeURIComponent(cfg.api)}&key=${cfg.key}`;
  if (navigator.share) navigator.share({ title: 'רשימת קניות משפחתית', url }).catch(() => {});
  else { await navigator.clipboard.writeText(url); alert('הקישור הועתק'); }
};

window.addEventListener('online', sync);
document.addEventListener('visibilitychange', () => { if (!document.hidden) sync(); });
setInterval(() => { if (!document.hidden) sync(); }, 15000);

if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js');
render();
sync();
