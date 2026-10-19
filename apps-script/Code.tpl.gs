/**
 * База «Отчёт ПТО» — ГК КРАШМАШ
 * Google Apps Script, привязанный к Google Таблице «База ПТО».
 * Отдаёт данные странице ПТО и кабинету руководителя, принимает изменения от ПТО.
 *
 * Установка: Расширения → Apps Script → вставить этот код → Сохранить →
 * Развернуть → Новое развертывание → Тип: Веб-приложение →
 * Запуск от имени: Я; У кого есть доступ: Все → Развернуть → Разрешить доступ →
 * скопировать URL веб-приложения.
 * После правки кода: Развернуть → Управление развертываниями → карандаш → Версия: Новая версия.
 */

// Ключ администратора: всё, включая сотрудников и историю изменений. Ключ руководителя (ГД): только согласованные данные.
// Ключи специалистов и начальника ПТО хранятся на листе «Сотрудники» (выдаются со страницы администратора).
const ADMIN_KEY = '{{ADMIN_KEY}}';
const DIRECTOR_KEY = '{{DIRECTOR_KEY}}';

const ROLE_NAMES = { spec: 'Специалист ПТО', head: 'Начальник ПТО', admin: 'Администратор' };
const ST = { draft: 'Черновик', sent: 'На проверке', ret: 'На доработке', ok: 'Согласован' };
const TYPE = { plan: 'План', fact: 'Факт' };
// Поля КС, которые можно менять и после согласования отчёта (движение КС: подписание, оплата)
const STATUS_FIELDS = ['agr', 'sign', 'pay', 'payDate', 'signDate', 'note'];

// Листы и столбцы базы (порядок столбцов в таблице можно менять — ищутся по заголовку).
const SCHEMA = {{SCHEMA}};
const LOG_SHEET = 'Журнал изменений';

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
// Часовой пояс таблицы; у таблицы, созданной из Excel, он бывает не задан — тогда Москва.
let TZ_ = null;
function tz_() { return TZ_ || (TZ_ = String(ss_().getSpreadsheetTimeZone() || Session.getScriptTimeZone() || '') || 'Europe/Moscow'); }

function sheet_(t) {
  const name = SCHEMA[t][0], cols = SCHEMA[t][1];
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols.map(function (c) { return c[1]; })])
      .setFontWeight('bold').setBackground('#2E671F').setFontColor('#ffffff').setWrap(true);
    sh.setFrozenRows(1);
  }
  // недостающие столбцы дописываются справа
  const hdr = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
  cols.forEach(function (c) {
    if (hdr.indexOf(c[1]) < 0) {
      const col = sh.getLastColumn() + 1;
      sh.getRange(1, col).setValue(c[1]).setFontWeight('bold').setBackground('#2E671F').setFontColor('#ffffff');
      hdr.push(c[1]);
    }
  });
  return { sh: sh, hdr: hdr, cols: cols };
}

/** Первичная настройка: создаёт недостающие листы. Можно запустить вручную из редактора. */
function setup() {
  Object.keys(SCHEMA).forEach(sheet_);
}

function fromCell_(v, typ) {
  if (v === '' || v === null) return typ === 'n' || typ === 'p' ? null : '';
  if (typ === 'd' || typ === 'm') {
    let d = v;
    if (Object.prototype.toString.call(v) !== '[object Date]') {
      const s = String(v).trim();
      let m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
      if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + (typ === 'd' ? '-' + ('0' + m[1]).slice(-2) : '');
      m = s.match(/^(\d{4})-(\d{2})(-(\d{2}))?/);
      if (m) return typ === 'm' ? m[1] + '-' + m[2] : s.slice(0, 10);
      return '';
    }
    return Utilities.formatDate(d, tz_(), typ === 'm' ? 'yyyy-MM' : 'yyyy-MM-dd');
  }
  if (typ === 'n' || typ === 'p') { const n = Number(v); return isNaN(n) ? null : n; }
  return String(v);
}
function toCell_(v, typ) {
  if (v === null || v === undefined || v === '') return '';
  if (typ === 'd') return Utilities.parseDate(String(v).slice(0, 10), tz_(), 'yyyy-MM-dd');
  if (typ === 'm') return Utilities.parseDate(String(v).slice(0, 7) + '-01', tz_(), 'yyyy-MM-dd');
  if (typ === 'n' || typ === 'p') return Number(v);
  return String(v);
}
const FMT_ = { d: 'dd.MM.yyyy', m: 'MM.yyyy', n: '#,##0.00', p: '0.0%' };

function readSheet_(t) {
  const s = sheet_(t), n = s.sh.getLastRow();
  if (n < 2) return [];
  const vals = s.sh.getRange(2, 1, n - 1, s.hdr.length).getValues();
  const idx = s.cols.map(function (c) { return s.hdr.indexOf(c[1]); });
  const rows = [];
  vals.forEach(function (v) {
    const r = {};
    let any = false;
    s.cols.forEach(function (c, i) { const x = fromCell_(v[idx[i]], c[2]); r[c[0]] = x; if (x !== '' && x !== null) any = true; });
    if (any) rows.push(r);
  });
  return rows;
}

function readAll_() {
  const lists = { resp: [], pto: [] };
  readSheet_('S').forEach(function (r) { if (r.resp) lists.resp.push(r.resp); if (r.pto) lists.pto.push(r.pto); });
  return {
    objects: readSheet_('O'), ks: readSheet_('K'), pays: readSheet_('P'), plan: readSheet_('L'),
    gar: readSheet_('G'), notes: readSheet_('N'), lists: lists
  };
}

function rowIndex_(s) {
  const n = s.sh.getLastRow(), col = s.hdr.indexOf('ID') + 1, map = {};
  if (n >= 2) s.sh.getRange(2, col, n - 1, 1).getValues().forEach(function (v, i) { if (v[0] !== '') map[String(v[0])] = i + 2; });
  return map;
}

const LOG_HDR = ['Когда', 'Кто', 'Лист', 'ID', 'Действие', 'Стало', 'Было', 'Роль'];
function logSheet_() {
  let sh = ss_().getSheetByName(LOG_SHEET);
  if (!sh) { sh = ss_().insertSheet(LOG_SHEET); sh.setFrozenRows(1); }
  if (sh.getLastColumn() < LOG_HDR.length) sh.getRange(1, 1, 1, LOG_HDR.length).setValues([LOG_HDR]).setFontWeight('bold');
  return sh;
}
/** items: [кто, лист, id, действие, стало, было, роль] */
function log_(items) {
  if (!items.length) return;
  const sh = logSheet_(), now = new Date();
  const vals = items.map(function (x) {
    return [now, x[0] || '', SCHEMA[x[1]] ? SCHEMA[x[1]][0] : x[1], String(x[2]), x[3], x[4] ? JSON.stringify(x[4]) : '', x[5] ? JSON.stringify(x[5]) : '', x[6] || ''];
  });
  sh.getRange(sh.getLastRow() + 1, 1, vals.length, LOG_HDR.length).setValues(vals);
}
function readLog_(n) {
  const sh = ss_().getSheetByName(LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const last = sh.getLastRow(), from = Math.max(2, last - (n || 1500) + 1);
  return sh.getRange(from, 1, last - from + 1, LOG_HDR.length).getValues().map(function (v) {
    const j = function (x) { try { return x ? JSON.parse(x) : null; } catch (e) { return null; } };
    return { at: Object.prototype.toString.call(v[0]) === '[object Date]' ? v[0].toISOString() : String(v[0]), by: String(v[1]), sheet: String(v[2]), id: String(v[3]), act: String(v[4]), after: j(v[5]), before: j(v[6]), role: String(v[7] || '') };
  }).reverse();
}

/** Записать строки (новые — в конец, существующие — по ID). base — «Изменено» на момент открытия: если строку уже поменяли, вернётся конфликт. */
function upsert_(t, rows, by, force, role) {
  const s = sheet_(t), map = rowIndex_(s), now = new Date().toISOString(), conflicts = [], saved = [];
  const width = s.hdr.length, maxR = s.sh.getMaxRows();
  // форматы столбцов: ID и «Изменено» — текст (иначе таблица превратит их в даты/числа)
  s.cols.forEach(function (c) {
    const f = (c[0] === 'id' || c[0] === 'upd') ? '@' : FMT_[c[2]];
    if (f && maxR > 1) s.sh.getRange(2, s.hdr.indexOf(c[1]) + 1, maxR - 1, 1).setNumberFormat(f);
  });
  const updCol = s.hdr.indexOf('Изменено'), fresh = [], logs = [];
  rows.forEach(function (r) {
    if (!r || !r.id) return;
    const at = map[String(r.id)];
    const line = at ? s.sh.getRange(at, 1, 1, width).getValues()[0] : new Array(width).fill('');
    if (at && r.base && !force && updCol >= 0) {
      const cur = String(line[updCol] || '');
      if (cur && cur !== r.base) { conflicts.push(r.id); return; }
    }
    delete r.base;
    r.upd = now; r.by = by || '';
    const before = {};
    if (at) s.cols.forEach(function (c) { if (c[0] in r && c[0] !== 'upd' && c[0] !== 'by') before[c[0]] = fromCell_(line[s.hdr.indexOf(c[1])], c[2]); });
    s.cols.forEach(function (c) { if (c[0] in r) line[s.hdr.indexOf(c[1])] = toCell_(r[c[0]], c[2]); });
    if (at) s.sh.getRange(at, 1, 1, width).setValues([line]);
    else fresh.push(line);
    logs.push([by, t, r.id, at ? 'изменено' : 'добавлено', r, at ? before : null, role]);
    saved.push({ id: r.id, upd: now });
  });
  if (fresh.length) {
    const start = s.sh.getLastRow() + 1;
    if (start + fresh.length - 1 > s.sh.getMaxRows()) s.sh.insertRowsAfter(s.sh.getMaxRows(), start + fresh.length - 1 - s.sh.getMaxRows());
    s.cols.forEach(function (c) {
      const f = (c[0] === 'id' || c[0] === 'upd') ? '@' : FMT_[c[2]];
      if (f) s.sh.getRange(start, s.hdr.indexOf(c[1]) + 1, fresh.length, 1).setNumberFormat(f);
    });
    s.sh.getRange(start, 1, fresh.length, width).setValues(fresh);
  }
  log_(logs);
  return { saved: saved, conflicts: conflicts };
}

function remove_(t, ids, by, role) {
  const s = sheet_(t), map = rowIndex_(s);
  const old = {}; readSheet_(t).forEach(function (r) { old[r.id] = r; });
  const rows = ids.map(function (id) { return map[String(id)]; }).filter(Boolean).sort(function (a, b) { return b - a; });
  rows.forEach(function (r) { s.sh.deleteRow(r); });
  log_(ids.map(function (id) { return [by, t, id, 'удалено', null, old[id] || null, role]; }));
  return rows.length;
}

/** Переименование объекта: имя меняется во всех листах. */
function renameObject_(from, to, by, role) {
  if (!from || !to || from === to) return;
  ['K', 'P', 'G', 'L', 'R'].forEach(function (t) {
    const s = sheet_(t), n = s.sh.getLastRow();
    if (n < 2) return;
    const oc = s.hdr.indexOf('Объект') + 1, ic = s.hdr.indexOf('ID') + 1;
    const objs = s.sh.getRange(2, oc, n - 1, 1).getValues(), ids = s.sh.getRange(2, ic, n - 1, 1).getValues();
    let ch = false;
    objs.forEach(function (v, i) {
      if (String(v[0]) === from) { v[0] = to; ch = true; if (t === 'L') ids[i][0] = String(ids[i][0]).replace(from + '|', to + '|'); if (t === 'R') ids[i][0] = String(ids[i][0]).replace('|' + from + '|', '|' + to + '|'); }
    });
    if (ch) { s.sh.getRange(2, oc, n - 1, 1).setValues(objs); if (t === 'L' || t === 'R') s.sh.getRange(2, ic, n - 1, 1).setValues(ids); }
  });
  log_([[by, 'O', from, 'переименовано → ' + to, null, null, role]]);
}

function saveLists_(lists) {
  const s = sheet_('S');
  const n = Math.max(lists.resp.length, lists.pto.length);
  if (s.sh.getLastRow() > 1) s.sh.getRange(2, 1, s.sh.getLastRow() - 1, s.hdr.length).clearContent();
  if (!n) return;
  const rc = s.hdr.indexOf(SCHEMA.S[1][0][1]), pc = s.hdr.indexOf(SCHEMA.S[1][1][1]);
  const vals = [];
  for (let i = 0; i < n; i++) { const line = new Array(s.hdr.length).fill(''); line[rc] = lists.resp[i] || ''; line[pc] = lists.pto[i] || ''; vals.push(line); }
  s.sh.getRange(2, 1, n, s.hdr.length).setValues(vals);
}


// =====================================================================
// РОЛИ, ОТЧЁТЫ, ПРАВА
// =====================================================================
function splitObjs_(s) { return String(s || '').split(';').map(function (x) { return x.trim(); }).filter(Boolean); }

/** Кто пришёл по ключу. null — ключ неизвестен. */
function who_(k) {
  if (!k) return null;
  if (k === ADMIN_KEY) return { role: 'admin', name: 'Администратор', objects: null };
  if (k === DIRECTOR_KEY) return { role: 'dir', name: 'Руководитель', objects: null };
  const u = readSheet_('U').filter(function (x) { return x.key && x.key === k && x.active !== 'нет' && x.mustChange !== 'да'; })[0];
  if (!u) return null;
  const role = u.role === ROLE_NAMES.head ? 'head' : (u.role === ROLE_NAMES.admin ? 'admin' : 'spec');
  return { role: role, name: u.name, id: u.id, objects: role === 'spec' ? splitObjs_(u.objects) : null };
}
function sees_(me, obj) { return !me.objects || me.objects.indexOf(obj) >= 0; }
function repId_(type, obj, month) { return type + '|' + obj + '|' + month; }
function reports_() { const m = {}; readSheet_('R').forEach(function (r) { m[r.id] = r; }); return m; }
function stOf_(reps, type, obj, month) { const r = reps[repId_(type, obj, month)]; return r && r.status ? r.status : ST.draft; }

/** '' — можно; иначе текст причины. statusOnly — меняются только поля движения КС. */
function canEdit_(me, reps, type, obj, month, statusOnly) {
  if (!obj || !month) return 'не указан объект или месяц';
  if (me.role === 'admin') return '';
  if (me.role === 'dir') return 'только просмотр';
  if (!sees_(me, obj)) return 'объект «' + obj + '» не закреплён за вами';
  const st = stOf_(reps, type, obj, month);
  if (st === ST.ok) return statusOnly ? '' : 'отчёт «' + (type === 'plan' ? 'План' : 'Факт') + ' ' + month + '» по объекту «' + obj + '» согласован — изменения закрыты';
  if (me.role === 'spec' && st === ST.sent) return statusOnly ? '' : 'отчёт отправлен на проверку — изменения закрыты до решения начальника ПТО';
  return '';
}

function same_(a, b) { return String(a === null || a === undefined ? '' : a) === String(b === null || b === undefined ? '' : b); }

/** Проверка строк перед записью. Возвращает { ok: [...], denied: [{id, why}] } */
function filterRows_(me, t, rows, reps) {
  const ok = [], denied = [];
  if (t === 'K' || t === 'L') {
    const cur = {}; readSheet_(t).forEach(function (r) { cur[r.id] = r; });
    rows.forEach(function (r) {
      const ex = cur[r.id];
      const obj = 'obj' in r ? r.obj : (ex && ex.obj), month = 'month' in r ? r.month : (ex && ex.month);
      let why = '';
      if (t === 'L') {
        why = canEdit_(me, reps, 'plan', obj, month, false);
        if (!why && ex && (ex.obj !== obj || ex.month !== month)) why = canEdit_(me, reps, 'plan', ex.obj, ex.month, false);
      } else {
        const changed = Object.keys(r).filter(function (f) { return f !== 'id' && f !== 'base' && f !== 'upd' && f !== 'by' && (!ex || !same_(ex[f], r[f])); });
        const statusOnly = !!ex && changed.every(function (f) { return STATUS_FIELDS.indexOf(f) >= 0; });
        why = canEdit_(me, reps, 'fact', obj, month, statusOnly);
        if (!why && ex && (ex.obj !== obj || ex.month !== month)) why = canEdit_(me, reps, 'fact', ex.obj, ex.month, false);
      }
      if (why) denied.push({ id: r.id, why: why }); else ok.push(r);
    });
    return { ok: ok, denied: denied };
  }
  // деньги, гарантии, объекты, «Главное за месяц», списки — начальник ПТО и администратор
  if (me.role === 'head' || me.role === 'admin') return { ok: rows, denied: [] };
  rows.forEach(function (r) { denied.push({ id: r.id, why: 'недостаточно прав' }); });
  return { ok: [], denied: denied };
}
function filterDel_(me, t, ids, reps) {
  if (t !== 'K' && t !== 'L') return (me.role === 'head' || me.role === 'admin') ? { ok: ids, denied: [] } : { ok: [], denied: ids.map(function (id) { return { id: id, why: 'недостаточно прав' }; }) };
  const cur = {}; readSheet_(t).forEach(function (r) { cur[r.id] = r; });
  const ok = [], denied = [];
  ids.forEach(function (id) {
    const ex = cur[id]; if (!ex) return;
    const why = canEdit_(me, reps, t === 'L' ? 'plan' : 'fact', ex.obj, ex.month, false);
    if (why) denied.push({ id: id, why: why }); else ok.push(id);
  });
  return { ok: ok, denied: denied };
}

/** Действие над отчётом: send | approve | return | reopen | comment */
function reportAction_(me, b) {
  const type = b.type, obj = b.obj, month = b.month, act = b.act, text = String(b.text || '').trim();
  if (!TYPE[type] || !obj || !/^\d{4}-\d{2}$/.test(month || '')) throw new Error('неверный отчёт');
  if (me.role === 'dir') throw new Error('только просмотр');
  if (!sees_(me, obj)) throw new Error('объект не закреплён за вами');
  const reps = reports_(), id = repId_(type, obj, month), r = reps[id] || { id: id, type: type, obj: obj, month: month, status: ST.draft, comments: '[]' };
  const st = r.status || ST.draft, head = me.role === 'head' || me.role === 'admin';
  let to = st;
  if (act === 'send') { if (st !== ST.draft && st !== ST.ret) throw new Error('отчёт уже отправлен или согласован'); to = ST.sent; r.sentAt = new Date().toISOString(); r.sentBy = me.name; }
  else if (act === 'approve') { if (!head) throw new Error('согласует начальник ПТО'); if (st === ST.ok) throw new Error('уже согласован'); to = ST.ok; r.decidedAt = new Date().toISOString(); r.decidedBy = me.name; }
  else if (act === 'return') {
    if (!head) throw new Error('возвращает начальник ПТО');
    if (st === ST.ok && me.role !== 'admin') throw new Error('согласованный отчёт может открыть только администратор');
    if (!text) throw new Error('напишите, что исправить');
    to = ST.ret; r.decidedAt = new Date().toISOString(); r.decidedBy = me.name;
  }
  else if (act === 'reopen') { if (me.role !== 'admin') throw new Error('только администратор'); to = ST.ret; }
  else if (act !== 'comment') throw new Error('неизвестное действие');
  if (act === 'comment' && !text) throw new Error('пустой комментарий');
  let cm = []; try { cm = JSON.parse(r.comments || '[]'); } catch (e) {}
  cm.push({ at: new Date().toISOString(), by: me.name, role: me.role, act: act, text: text, from: st, to: to });
  r.comments = JSON.stringify(cm); r.status = to; r.type = type; r.obj = obj; r.month = month;
  upsert_('R', [r], me.name, true, me.role);
  return r;
}

const TR_ = {а:'a',б:'b',в:'v',г:'g',д:'d',е:'e',ё:'e',ж:'zh',з:'z',и:'i',й:'y',к:'k',л:'l',м:'m',н:'n',о:'o',п:'p',р:'r',с:'s',т:'t',у:'u',ф:'f',х:'h',ц:'ts',ч:'ch',ш:'sh',щ:'sch',ъ:'',ы:'y',ь:'',э:'e',ю:'yu',я:'ya'};
function translit_(s) { return String(s || '').toLowerCase().split('').map(function (c) { return c in TR_ ? TR_[c] : c; }).join(''); }
/** Логин по ФИО: «Яковлева Н.Г.» → yakovleva.ng */
function loginFor_(name) {
  const p = String(name || '').trim().split(/\s+/), last = translit_(p[0] || '').replace(/[^a-z0-9]/g, '');
  const ini = translit_(p.slice(1).join('')).replace(/[^a-z]/g, '').slice(0, 2);
  return last + (ini ? '.' + ini : '');
}
function tempPassword_() {
  const a = 'abcdefghjkmnpqrstuvwxyz23456789'; let s = '';
  for (let i = 0; i < 8; i++) s += a.charAt(Math.floor(Math.random() * a.length));
  return s;
}
function newKey_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }

function saveUser_(me, u, resetPass) {
  if (me.role !== 'admin') throw new Error('только администратор');
  if (!u || !String(u.name || '').trim()) throw new Error('укажите ФИО');
  const all = readSheet_('U');
  const ex = u.id ? all.filter(function (x) { return x.id === u.id; })[0] : null;
  const login = String(u.login || '').trim().toLowerCase() || loginFor_(u.name);
  if (!/^[a-z0-9._-]{3,40}$/.test(login)) throw new Error('логин — латинские буквы, цифры, точка (от 3 символов)');
  if (all.some(function (x) { return x.id !== (ex && ex.id) && String(x.login || '').toLowerCase() === login; })) throw new Error('логин «' + login + '» уже занят');
  const row = { id: ex ? ex.id : ('u' + Date.now().toString(36)), name: String(u.name).trim(), role: ROLE_NAMES[u.role] || ROLE_NAMES.spec,
    objects: (u.objects || []).join('; '), active: u.active === false ? 'нет' : 'да', login: login };
  if (!ex || resetPass || !ex.password) { row.password = tempPassword_(); row.mustChange = 'да'; row.key = newKey_(); }
  else if (!ex.key) row.key = newKey_();
  upsert_('U', [row], me.name, true, me.role);
  return readSheet_('U').filter(function (x) { return x.id === row.id; })[0];
}

/** Вход по логину и паролю. Ключ сеанса выдаётся, только когда временный пароль уже сменён. */
function login_(b, change) {
  const login = String(b.login || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache(), ck = 'fail:' + login, fails = +(cache.get(ck) || 0);
  if (fails >= 8) throw new Error('слишком много неудачных попыток — подождите 15 минут');
  const u = readSheet_('U').filter(function (x) { return String(x.login || '').toLowerCase() === login && x.active !== 'нет'; })[0];
  if (!u || String(u.password || '') !== String(b.password || '')) { cache.put(ck, String(fails + 1), 900); throw new Error('неверный логин или пароль'); }
  cache.remove(ck);
  const patch = { id: u.id, lastLogin: new Date().toISOString() };
  if (change) {
    const np = String(b.newPassword || '');
    if (np.length < 6) throw new Error('новый пароль — не короче 6 символов');
    if (np === u.password) throw new Error('новый пароль должен отличаться от временного');
    patch.password = np; patch.mustChange = 'нет';
  } else if (u.mustChange === 'да') {
    return { mustChange: true, name: u.name };
  }
  upsert_('U', [patch], u.name, true, 'login');
  return { key: u.key, name: u.name };
}

/** Данные для роли: специалист — свои объекты; ГД — только согласованное; начальник и администратор — всё. */
function dataFor_(me) {
  const d = readAll_(), reps = reports_();
  d.reports = Object.keys(reps).map(function (k) { return reps[k]; });
  if (me.role === 'spec') {
    const mine = function (r) { return sees_(me, r.obj); };
    d.objects = d.objects.filter(function (o) { return sees_(me, o.name); });
    d.ks = d.ks.filter(mine); d.plan = d.plan.filter(mine); d.reports = d.reports.filter(mine);
    d.pays = []; d.gar = [];
  } else if (me.role === 'dir') {
    d.ks = d.ks.filter(function (k) { return stOf_(reps, 'fact', k.obj, k.month) === ST.ok; });
    d.plan = d.plan.filter(function (p) { return stOf_(reps, 'plan', p.obj, p.month) === ST.ok; });
    d.reports = [];
    d.ks.forEach(function (k) { delete k.by; }); d.plan.forEach(function (p) { delete p.by; });
  }
  return d;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const me = who_(p.k);
  if (!me) return out_({ ok: false, error: 'key' });
  const pub = { role: me.role, name: me.name, objects: me.objects };
  if (p.a === 'ping') return out_({ ok: true, role: me.role, me: pub });
  if (p.a === 'all') return out_({ ok: true, role: me.role, me: pub, data: dataFor_(me), time: new Date().toISOString() });
  if (p.a === 'log' && me.role === 'admin') return out_({ ok: true, log: readLog_(+p.n || 1500) });
  if (p.a === 'users' && me.role === 'admin') return out_({ ok: true, users: readSheet_('U') });
  return out_({ ok: false, error: 'action' });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'json' }); }
  if (b.a === 'login' || b.a === 'passwd') {
    const lk = LockService.getScriptLock();
    try { lk.waitLock(25000); return out_(Object.assign({ ok: true }, login_(b, b.a === 'passwd'))); }
    catch (err) { return out_({ ok: false, error: String(err && err.message || err) }); }
    finally { lk.releaseLock(); }
  }
  const me = who_(b.k);
  if (!me) return out_({ ok: false, error: 'key' });
  if (me.role === 'dir') return out_({ ok: false, error: 'только просмотр' });
  if (b.t && (!SCHEMA[b.t] || b.t === 'U' || b.t === 'R')) return out_({ ok: false, error: 'table' });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    let res = {};
    if (b.a === 'save') {
      const reps = reports_();
      if (b.t === 'O' && b.rename) { if (me.role !== 'head' && me.role !== 'admin') throw new Error('недостаточно прав'); renameObject_(b.rename.from, b.rename.to, me.name, me.role); }
      const f = filterRows_(me, b.t, b.rows || [], reps);
      res = upsert_(b.t, f.ok, me.name, !!b.force, me.role);
      res.denied = f.denied;
      if (b.del && b.del.length) { const fd = filterDel_(me, b.t, b.del, reps); res.deleted = remove_(b.t, fd.ok, me.name, me.role); res.denied = res.denied.concat(fd.denied); }
    } else if (b.a === 'del') {
      const fd = filterDel_(me, b.t, b.ids || [], reports_());
      res.deleted = remove_(b.t, fd.ok, me.name, me.role); res.denied = fd.denied;
    } else if (b.a === 'report') {
      res.report = reportAction_(me, b);
    } else if (b.a === 'lists') {
      if (me.role !== 'head' && me.role !== 'admin') throw new Error('недостаточно прав');
      saveLists_({ resp: (b.lists && b.lists.resp) || [], pto: (b.lists && b.lists.pto) || [] });
    } else if (b.a === 'user') {
      res.user = saveUser_(me, b.user, !!b.resetPass);
    } else if (b.a === 'wipe') {
      // Очистка рабочих данных (КС, план, деньги, гарантии, «Главное», отчёты) — только администратор
      if (me.role !== 'admin' || b.confirm !== 'ОЧИСТИТЬ') throw new Error('только администратор');
      ['K', 'L', 'P', 'G', 'N', 'R'].forEach(function (t) { const s = sheet_(t); if (s.sh.getLastRow() > 1) s.sh.deleteRows(2, s.sh.getLastRow() - 1); });
      log_([[me.name, 'K', '*', 'очистка базы: КС, план, поступления, гарантии, комментарии, отчёты', null, null, me.role]]);
    } else return out_({ ok: false, error: 'action' });
    return out_(Object.assign({ ok: true }, res));
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}
