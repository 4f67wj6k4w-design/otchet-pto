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
const ADMIN_KEY = 'ВСТАВЬТЕ_КЛЮЧ_АДМИНИСТРАТОРА';
const DIRECTOR_KEY = 'ВСТАВЬТЕ_КЛЮЧ_РУКОВОДИТЕЛЯ';

const ROLE_NAMES = { spec: 'Специалист ПТО', head: 'Начальник ПТО', admin: 'Администратор' };
const ST = { draft: 'Черновик', sent: 'На проверке', ret: 'На доработке', ok: 'Согласован' };
const TYPE = { plan: 'План', fact: 'Факт' };
// Поля КС, которые можно менять и после согласования отчёта (движение КС: подписание, оплата)
const STATUS_FIELDS = ['agr', 'sign', 'pay', 'payDate', 'signDate', 'note', 'remarks', 'dsNeed', 'dsDate', 'letterDate'];

// Листы и столбцы базы (порядок столбцов в таблице можно менять — ищутся по заголовку).
const SCHEMA = {"O": ["Объекты", [["id", "ID", "s"], ["name", "Объект", "s"], ["resp", "Ответственный", "s"], ["pto", "Специалист ПТО", "s"], ["status", "Статус", "s"], ["sum", "Сумма договора с ДС, руб", "n"], ["done0", "Выполнено до 01.01.2026, руб", "n"], ["ret", "% удержания", "p"], ["payDays", "Срок оплаты, дней", "n"], ["due", "Срок по договору", "d"], ["sched", "График подачи КС", "s"], ["contract", "№ договора", "s"], ["contractDate", "Дата договора", "d"], ["start", "Начало по договору", "d"], ["rp", "Код объекта в «Рапорте прораба»", "s"], ["note", "Примечание", "s"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "K": ["Журнал КС", [["id", "ID", "s"], ["date", "Дата КС", "d"], ["month", "Отчётный месяц", "m"], ["obj", "Объект", "s"], ["num", "№ КС / описание", "s"], ["period", "Период работ", "s"], ["ks2", "Выполнение КС-2, руб", "n"], ["ks3", "К оплате КС-3, руб", "n"], ["agr", "Согласование", "s"], ["sign", "Подписание", "s"], ["pay", "Оплата", "s"], ["payDate", "Дата оплаты", "d"], ["signDate", "Дата подписания", "d"], ["remarks", "Замечания заказчика", "s"], ["dsNeed", "Нужен ДС", "s"], ["dsDate", "ДС подан заказчику", "d"], ["letterDate", "Письмо / претензия заказчику", "d"], ["note", "Примечание", "s"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "P": ["Поступления", [["id", "ID", "s"], ["date", "Дата поступления", "d"], ["obj", "Объект", "s"], ["sum", "Сумма, руб", "n"], ["kind", "Вид поступления", "s"], ["basis", "№ КС / основание", "s"], ["note", "Примечание", "s"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "L": ["План", [["id", "ID", "s"], ["obj", "Объект", "s"], ["month", "Месяц", "m"], ["sum", "План, руб", "n"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "G": ["Гарантии", [["id", "ID", "s"], ["obj", "Объект", "s"], ["due", "Срок возврата", "d"], ["sum", "Сумма ГУ, руб", "n"], ["status", "Статус оплаты", "s"], ["party", "Контрагент", "s"], ["note", "Примечание", "s"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "N": ["Главное за месяц", [["id", "ID", "s"], ["month", "Месяц", "m"], ["text", "Комментарий начальника ПТО", "s"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "S": ["Списки", [["resp", "Ответственные (руководители объектов)", "s"], ["pto", "Специалисты ПТО", "s"]]], "U": ["Сотрудники", [["id", "ID", "s"], ["name", "ФИО", "s"], ["role", "Роль", "s"], ["objects", "Объекты (через ;)", "s"], ["login", "Логин", "s"], ["password", "Пароль", "s"], ["mustChange", "Сменить пароль при входе", "s"], ["lastLogin", "Последний вход", "s"], ["key", "Ключ ссылки", "s"], ["active", "Активен", "s"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "R": ["Отчёты", [["id", "ID", "s"], ["type", "Вид", "s"], ["obj", "Объект", "s"], ["month", "Месяц", "m"], ["status", "Статус", "s"], ["sentAt", "Отправлен", "s"], ["sentBy", "Кто отправил", "s"], ["decidedAt", "Решение", "s"], ["decidedBy", "Кто решил", "s"], ["comments", "Комментарии (JSON)", "s"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "Q": ["Настройки KPI", [["id", "Параметр", "s"], ["value", "Значение", "s"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]]};
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
    gar: readSheet_('G'), notes: readSheet_('N'), kpi: readSheet_('Q'), lists: lists
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
  // форматы столбцов: ID, «Изменено» и все текстовые поля — текст (иначе таблица превратит «9 Мая» или «15.03» в дату)
  s.cols.forEach(function (c) {
    const f = (c[0] === 'id' || c[0] === 'upd' || c[2] === 's') ? '@' : FMT_[c[2]];   // текстовые столбцы — как текст: «9 Мая», «15.03» не станут датами
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
      const f = (c[0] === 'id' || c[0] === 'upd' || c[2] === 's') ? '@' : FMT_[c[2]];   // текстовые столбцы — как текст: «9 Мая», «15.03» не станут датами
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
    if (ch) { s.sh.getRange(2, oc, n - 1, 1).setNumberFormat('@').setValues(objs); if (t === 'L' || t === 'R') s.sh.getRange(2, ic, n - 1, 1).setValues(ids); }
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
  // настройки KPI: веса и фонд — администратор; доли премии сотрудников — начальник ПТО и администратор
  if (t === 'Q') {
    rows.forEach(function (r) {
      const id = String(r.id || '');
      if (me.role === 'admin' || (me.role === 'head' && id.indexOf('share:') === 0)) ok.push(r);
      else denied.push({ id: id, why: 'настройки KPI меняет администратор' });
    });
    return { ok: ok, denied: denied };
  }
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
  const hu = readSheet_('U').filter(function (x) { return x.role === ROLE_NAMES.head && x.active !== 'нет'; })[0];
  d.headName = hu ? hu.name : '';
  if (me.role === 'spec') {
    const mine = function (r) { return sees_(me, r.obj); };
    d.objects = d.objects.filter(function (o) { return sees_(me, o.name); });
    d.ks = d.ks.filter(mine); d.plan = d.plan.filter(mine); d.reports = d.reports.filter(mine);
    d.pays = []; d.gar = [];
    d.kpi = d.kpi.filter(function (r) { return String(r.id).indexOf('w:') === 0; });   // веса — да, деньги и доли — нет
  } else if (me.role === 'dir') {
    d.ks = d.ks.filter(function (k) { return stOf_(reps, 'fact', k.obj, k.month) === ST.ok; });
    d.plan = d.plan.filter(function (p) { return stOf_(reps, 'plan', p.obj, p.month) === ST.ok; });
    d.reports = []; d.kpi = [];
    d.ks.forEach(function (k) { delete k.by; }); d.plan.forEach(function (p) { delete p.by; });
  }
  return d;
}

// ---------- «Рапорт прораба»: только чтение ----------
// Скрипт ПТО читает таблицу рапортов напрямую. В базе рапортов ничего не меняется, прорабы работают как раньше.
const RP_ID = '1pq2K-fqzqvop-KEYQV0QVNBDMJPnBIthg_JrO5Iiwp8';   // Google Таблица «Рапорты прорабов — ГК КРАШМАШ»
function rpIso_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const s = String(v || '');
  let m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})/); if (m) return m[3] + '-' + m[2] + '-' + m[1];
  m = s.match(/^\d{4}-\d{2}-\d{2}/); return m ? m[0] : '';
}
function rpAddDays_(iso, n) { const d = new Date(iso + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
/** Срок из паспорта объекта: start / due или строка «Сроки…: 01.09.2026 – 12.11.2026» (как в «Рапорте прораба»). */
function rpDue_(c) {
  const iso = function (s) { const m = /(\d{2})\.(\d{2})\.(\d{4})/.exec(s || ''); return m ? m[3] + '-' + m[2] + '-' + m[1] : null; };
  let start = c.start || null, due = c.due || null;
  (c.info || []).forEach(function (s) {
    if (due || !/срок/i.test(s)) return;
    const m = /(\d{2}\.\d{2}\.\d{4})\s*[–—-]\s*(\d{2}\.\d{2}\.\d{4})/.exec(s);
    if (m) { start = start || iso(m[1]); due = iso(m[2]); }
    else { const m2 = /до\s*(\d{2}\.\d{2}\.\d{4})/.exec(s); if (m2) due = iso(m2[1]); }
  });
  return { start: start, due: due };
}
/** Выполнено по работе на дату d (та же логика, что в «Рапорте прораба»: «выполнено ранее» — на дату d0date). */
function rpCum_(c, rs, code, d) {
  const d0 = (((c.works || []).filter(function (w) { return w.c === code; })[0]) || {}).d0 || 0, dd = c.d0date || '';
  let t = +d0 || 0;
  rs.forEach(function (r) {
    const v = r.works && r.works[code]; if (typeof v !== 'number') return;
    if (!dd) { if (r.date <= d) t += v; }
    else if (r.date > dd && r.date <= d) t += v;
    else if (r.date <= dd && r.date > d) t -= v;
  });
  return t;
}
function rpCfgs_(ss) {
  const sh = ss.getSheetByName('Объекты'), n = sh ? sh.getLastRow() : 0, out = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 4).getValues().forEach(function (v) {
    try { const c = JSON.parse(v[3]); if (c && c.pid) out[c.pid] = c; } catch (e) { if (v[0]) out[String(v[0])] = { pid: String(v[0]), name: String(v[1] || v[0]), works: [] }; }
  });
  return out;
}
/** Запустите один раз из редактора после обновления скрипта: Google попросит разрешить доступ к таблице рапортов. */
function checkRapport() {
  const n = rpObjects_().length;
  Logger.log('Доступ к «Рапорту прораба» есть, объектов: ' + n);
  return n;
}
/** Список объектов «Рапорта прораба» — для привязки в справочнике ПТО. */
function rpObjects_() {
  const c = rpCfgs_(SpreadsheetApp.openById(RP_ID));
  return Object.keys(c).map(function (p) { return { pid: p, name: c[p].name || p, closed: !!c[p].closed }; });
}
/** Рапорты по объектам, которые видит сотрудник: паспорт, сроки, выполнено по работам, вывоз, техника, рапорты за days дней. */
function rpData_(me, days) {
  const map = {};
  readSheet_('O').forEach(function (o) { if (o.rp && sees_(me, o.name)) splitObjs_(String(o.rp).replace(/,/g, ';')).forEach(function (p) { map[p] = o.name; }); });
  const pids = Object.keys(map);
  if (!pids.length) return { objects: [], today: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd') };
  const ss = SpreadsheetApp.openById(RP_ID), cfg = rpCfgs_(ss), reps = {};
  const sr = ss.getSheetByName('Рапорты'), nr = sr ? sr.getLastRow() : 0;
  if (nr >= 2) sr.getRange(2, 1, nr - 1, 12).getValues().forEach(function (v) {
    const pid = String(v[1] || ''); if (!map[pid]) return;
    try { const r = JSON.parse(v[11]); if (r && r.date) (reps[pid] = reps[pid] || []).push(r); } catch (e) {}
  });
  const fin = {}, sf = ss.getSheetByName('Завершение'), nf = sf ? sf.getLastRow() : 0;
  if (nf >= 2) sf.getRange(2, 1, nf - 1, 8).getValues().forEach(function (v) {
    const pid = String(v[0] || ''); if (!map[pid]) return;
    fin[pid] = { foreman: String(v[2] || ''), reportDate: rpIso_(v[3]), status: String(v[5] || ''), confirmed: rpIso_(v[6]) };
  });
  const today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'), from = rpAddDays_(today, -(days || 31));
  const objects = pids.map(function (p) {
    const c = cfg[p] || { pid: p, name: p, works: [] };
    const rs = (reps[p] || []).sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.shift < b.shift ? -1 : 1); });
    const dd = rpDue_(c);
    const works = (c.works || []).map(function (w) {
      return { c: w.c, n: w.n, u: w.u || '', t: +w.t || 0, done: rpCum_(c, rs, w.c, today),
        d7: rs.filter(function (r) { return r.date > rpAddDays_(today, -8); }).reduce(function (a, r) { const x = r.works && r.works[w.c]; return a + (typeof x === 'number' ? x : 0); }, 0) };
    });
    const prog = works.filter(function (w) { return w.t > 0; });
    const pct = prog.length ? prog.reduce(function (a, w) { return a + Math.max(0, Math.min(1, w.done / w.t)); }, 0) / prog.length : null;
    const haul = {}, mach = {};
    rs.forEach(function (r) {
      (r.haul || []).forEach(function (h) { const k = h.m || '—'; const x = haul[k] = haul[k] || { m: k, t: 0, v: 0, w: 0 }; x.t += +h.t || 0; x.v += +h.v || 0; x.w += +h.w || 0; });
      (r.mach || []).forEach(function (m) { const k = String(m.n || '—').trim(); const x = mach[k] = mach[k] || { n: k, h: 0, days: 0 }; x.h += +m.h || 0; if (+m.h > 0) x.days++; });
    });
    const day = function (r) {
      return { date: r.date, shift: r.shift === 'N' ? 'Ночь' : 'День', foreman: r.foreman || '', idle: !!r.idle,
        people: ((r.people && +r.people.own) || 0) + (r.subs || []).reduce(function (a, x) { return a + (+x.p || 0); }, 0),
        subs: r.subs || [], works: r.works || {}, extra: r.extra || [], haul: r.haul || [], mach: r.mach || [], downtime: r.downtime || [], other: r.other || '' };
    };
    return { pid: p, obj: map[p], name: c.name || p, start: dd.start, due: dd.due, closed: !!c.closed, closedAt: c.closedAt || '',
      finish: fin[p] || null, first: rs.length ? rs[0].date : '', last: rs.length ? rs[rs.length - 1].date : '', nRep: rs.length,
      pct: pct, works: works, haul: Object.keys(haul).map(function (k) { return haul[k]; }), mach: Object.keys(mach).map(function (k) { return mach[k]; }),
      foremen: rs.slice(-20).map(function (r) { return r.foreman; }).filter(function (x, i, a) { return x && a.indexOf(x) === i; }),
      days: rs.filter(function (r) { return r.date >= from; }).map(day) };
  });
  return { objects: objects, today: today, from: from };
}

/** Месяцы, в которые по рапортам прорабов на объекте были работы (для KPI «работы закрыты КС»). { "Объект ПТО": ["2026-09", …] } */
function rpActivity_(me) {
  const map = {};
  readSheet_('O').forEach(function (o) { if (o.rp && (!me || sees_(me, o.name))) splitObjs_(String(o.rp).replace(/,/g, ';')).forEach(function (p) { map[p] = o.name; }); });
  const out = {};
  if (!Object.keys(map).length) return out;
  const sr = SpreadsheetApp.openById(RP_ID).getSheetByName('Рапорты'), nr = sr ? sr.getLastRow() : 0;
  if (nr >= 2) sr.getRange(2, 1, nr - 1, 12).getValues().forEach(function (v) {
    const nm = map[String(v[1] || '')]; if (!nm) return;
    let r; try { r = JSON.parse(v[11]); } catch (e) { return; }
    if (!r || !r.date || r.idle) return;
    const busy = Object.keys(r.works || {}).length || (r.haul || []).length || (r.extra || []).length || (r.mach || []).some(function (m) { return +m.h > 0; });
    if (!busy) return;
    const m = String(r.date).slice(0, 7); (out[nm] = out[nm] || []); if (out[nm].indexOf(m) < 0) out[nm].push(m);
  });
  return out;
}
/** KPI для сводки генерального: отдел и начальник ПТО (без специалистов). Считается здесь же — calc.js встроен в скрипт при сборке. */
function kpiForDirector_(p) {
  const d = readAll_(); d.reports = readSheet_('R');
  let act = null; try { act = rpActivity_(null); } catch (e) {}
  const today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  const r = PTO.kpi(d, { today: today, kind: p.kind || 'month', period: p.period || '', act: act });
  return { period: r.period, list: r.list, months: r.months, done: r.done, hasAct: r.hasAct, dept: r.dept, head: r.head,
    money: r.money ? { fund: r.money.fund, dept: r.money.dept } : null };
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const me = who_(p.k);
  if (!me) return out_({ ok: false, error: 'key' });
  const pub = { role: me.role, name: me.name, objects: me.objects };
  if (p.a === 'ping') return out_({ ok: true, role: me.role, me: pub });
  if (p.a === 'all') return out_({ ok: true, role: me.role, me: pub, data: dataFor_(me), time: new Date().toISOString() });
  if (p.a === 'rp' && me.role !== 'dir') {
    try { return out_(Object.assign({ ok: true }, rpData_(me, Math.min(366, +p.days || 31)))); }
    catch (err) { return out_({ ok: false, error: 'Нет доступа к базе «Рапорт прораба»: ' + String(err && err.message || err) }); }
  }
  if (p.a === 'rpact' && me.role !== 'dir') {
    try { return out_({ ok: true, act: rpActivity_(me) }); } catch (err) { return out_({ ok: true, act: null, warn: String(err && err.message || err) }); }
  }
  if (p.a === 'kpi' && me.role === 'dir') return out_(Object.assign({ ok: true }, kpiForDirector_(p)));
  if (p.a === 'rpobj' && (me.role === 'admin' || me.role === 'head')) {
    try { return out_({ ok: true, objects: rpObjects_() }); } catch (err) { return out_({ ok: false, error: String(err && err.message || err) }); }
  }
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

// ---------- расчёты (src/calc.js, вставляются при сборке) ----------
/* Расчёты отчёта ПТО — повторяют формулы старого Excel-отчёта (листы «Расчёт», «План_расч», «Объекты», «Журнал КС»).
   Работает и в браузере (window.PTO), и в node (module.exports) — для сверки с Excel. Даты — строки ГГГГ-ММ-ДД, месяцы — ГГГГ-ММ. */
(function (root) {
"use strict";
const BASE = "2026-01";            // «Выполнено до 01.01.2026» в справочнике объектов — всё, что раньше, в журнал не входит
const INF = 1e15;
const MN = ["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];
const MNs = ["янв","фев","мар","апр","май","июн","июл","авг","сен","окт","ноя","дек"];
const QN = ["I","II","III","IV"];

// ---------- даты ----------
const D = s => { const [y, m, d] = s.split("-").map(Number); return Date.UTC(y, m - 1, d || 1); };
const S = t => new Date(t).toISOString().slice(0, 10);
const addDays = (s, n) => S(D(s) + n * 864e5);
const ym = s => s.slice(0, 7);
const m1 = m => m + "-01";
const addM = (m, n) => { let [y, mo] = m.split("-").map(Number); mo += n; y += Math.floor((mo - 1) / 12); mo = ((mo - 1) % 12 + 12) % 12 + 1; return y + "-" + String(mo).padStart(2, "0"); };
const eom = m => addDays(m1(addM(m, 1)), -1);
const monday = s => { const w = (new Date(D(s)).getUTCDay() + 6) % 7; return addDays(s, -w); };
const ru = s => s ? s.slice(8, 10) + "." + s.slice(5, 7) + "." + s.slice(0, 4) : "";
const monthLabel = m => MN[+m.slice(5, 7) - 1] + " " + m.slice(0, 4);
const num = v => (v === null || v === undefined || v === "" || isNaN(v)) ? 0 : +v;
const has = v => !(v === null || v === undefined || v === "");

// ---------- строки журнала ----------
function index(data) {
  const byName = {};
  (data.objects || []).forEach(o => { byName[o.name] = o; });
  return byName;
}
function ksStatus(k) {
  if (k.pay === "Оплачена") return "paid";
  if (k.sign === "Подписана") return "signed";
  if (k.agr === "Согласована") return "agreed";
  return "review";
}
const STATUS_LABEL = { review: "На согласовании", agreed: "Согласована, не подписана", signed: "Подписана, ждёт оплаты", paid: "Оплачена" };
function enrichKs(k, byName) {
  const o = byName[k.obj];
  const toPay = !has(k.ks2) ? null : (has(k.ks3) ? +k.ks3 : +k.ks2 * (1 - (o && has(o.ret) ? +o.ret : 0)));
  const st = ksStatus(k);
  let expect = "";
  if (has(k.ks2) && k.pay !== "Оплачена" && k.month) {
    const base = k.date || eom(k.month);
    const days = o && has(o.payDays) ? +o.payDays : 30;
    expect = addDays(base, days + (k.sign === "Подписана" ? 0 : (k.agr === "Согласована" ? 15 : 30)));
  }
  let ctrl = "";
  if (!k.obj && !has(k.ks2)) ctrl = "";
  else if (!k.obj) ctrl = "нет объекта";
  else if (!o) ctrl = "объект не из списка";
  else if (!has(k.ks2)) ctrl = "нет суммы";
  else if (!k.month) ctrl = "нет месяца";
  return Object.assign({}, k, {
    toPay, st, stLabel: STATUS_LABEL[st], expect,
    wdate: has(k.ks2) ? (k.date || (k.month ? eom(k.month) : "")) : "",
    resp: o ? o.resp || "" : "", ctrl
  });
}

// ---------- объекты (справочник с расчётными столбцами) ----------
function objectRows(data, today) {
  const byName = index(data);
  const ks = (data.ks || []).map(k => enrichKs(k, byName));
  return (data.objects || []).map(o => {
    const mine = ks.filter(k => k.obj === o.name);
    const done26 = mine.filter(k => k.month >= BASE).reduce((a, k) => a + num(k.ks2), 0);
    const doneAll = num(o.done0) + done26;
    const remain = has(o.sum) ? +o.sum - doneAll : null;
    const pct = num(o.sum) ? doneAll / +o.sum : null;
    const unpaid = mine.filter(k => k.pay !== "Оплачена").reduce((a, k) => a + num(k.toPay), 0);
    const debt = mine.filter(k => k.sign === "Подписана" && k.pay !== "Оплачена").reduce((a, k) => a + num(k.toPay), 0);
    const got = (data.pays || []).filter(p => p.obj === o.name).reduce((a, p) => a + num(p.sum), 0);
    let ctrl = "";
    if (num(o.sum) > 0 && doneAll > +o.sum) ctrl = "Выполнено больше договора";
    else if (o.due && o.due < today && num(remain) > 0 && o.status !== "Завершён") ctrl = "Срок договора истёк";
    return Object.assign({}, o, { done26, doneAll, remain, pct, unpaid, debt, got, ctrl, ksCount: mine.length });
  });
}

// ---------- периоды ----------
function dataRange(data, today) {
  const ms = (data.ks || []).filter(k => has(k.ks2) && k.month).map(k => k.month);
  const minM = ms.length ? ms.reduce((a, b) => a < b ? a : b) : BASE;
  const y0 = Math.min(+minM.slice(0, 4), +BASE.slice(0, 4));
  const y1 = Math.max(+today.slice(0, 4) + 1, ...(data.plan || []).map(p => +p.month.slice(0, 4)));
  return { minM: minM < BASE ? minM : BASE, y0, y1 };
}
function periods(kind, data, today) {
  const { minM, y0, y1 } = dataRange(data, today), out = [];
  if (kind === "week") {
    for (let s = monday(m1(minM)); s <= y1 + "-12-31"; s = addDays(s, 7)) {
      const e = addDays(s, 6);
      out.push({ key: s, start: s, end: e, label: s.slice(8, 10) + "." + s.slice(5, 7) + "–" + ru(e) });
    }
  } else if (kind === "month") {
    for (let m = minM; m <= y1 + "-12"; m = addM(m, 1)) out.push({ key: m, start: m1(m), end: eom(m), label: monthLabel(m) });
  } else if (kind === "quarter") {
    for (let y = y0; y <= y1; y++) for (let q = 0; q < 4; q++) {
      const m = y + "-" + String(q * 3 + 1).padStart(2, "0");
      if (eom(addM(m, 2)) < m1(minM)) continue;
      out.push({ key: y + "-Q" + (q + 1), start: m1(m), end: eom(addM(m, 2)), label: QN[q] + " квартал " + y });
    }
  } else {
    for (let y = y0; y <= y1; y++) out.push({ key: String(y), start: y + "-01-01", end: y + "-12-31", label: y + " год" });
  }
  return out;
}
function lastKsMonth(data, today) {             // Расчёт!B20
  const ms = (data.ks || []).filter(k => k.month).map(k => k.month);
  const cur = ym(today);
  if (!ms.length) return cur;
  const mx = ms.reduce((a, b) => a > b ? a : b);
  return mx < cur ? mx : cur;
}
function defaultPeriod(kind, data, today) {      // Расчёт!B6: последний период, начавшийся не позже последних данных
  const list = periods(kind, data, today);
  const byName = index(data);
  let ref;
  if (kind === "week") {
    const ws = (data.ks || []).map(k => enrichKs(k, byName).wdate).filter(Boolean);
    ref = ws.length ? ws.reduce((a, b) => a > b ? a : b) : today;
  } else {
    const ms = (data.ks || []).filter(k => k.month).map(k => k.month);
    ref = ms.length ? m1(ms.reduce((a, b) => a > b ? a : b)) : today;
  }
  if (today < ref) ref = today;
  let p = list[0];
  list.forEach(x => { if (x.start <= ref) p = x; });
  return p;
}

// ---------- сводка за период ----------
function compute(data, opt) {
  const today = opt.today, kind = opt.kind || "month";
  const byName = index(data);
  const ks = (data.ks || []).map(k => enrichKs(k, byName));
  const list = periods(kind, data, today);
  const P = list.find(x => x.key === opt.period) || defaultPeriod(kind, data, today);
  const week = kind === "week";
  const B10 = P.start, B11 = P.end;
  const B13 = week ? ym(addDays(B10, 3)) : null;               // месяц недели (по четвергу)
  const planStart = week ? B13 : ym(B10), planEnd = week ? B13 : ym(B11);
  const ytdEndDate = week ? eom(B13) : B11;
  const year = +ytdEndDate.slice(0, 4);
  const yStart = year + "-01";
  const B20 = lastKsMonth(data, today);
  const commentMonth = week ? B13 : ym(B11);

  const inP = k => has(k.ks2) && (week ? (k.wdate && k.wdate >= B10 && k.wdate <= B11) : (k.month && m1(k.month) >= B10 && m1(k.month) <= B11));
  const sumKs = (arr, f) => arr.reduce((a, k) => a + num(f ? f(k) : k.ks2), 0);
  const ksObj = name => ks.filter(k => k.obj === name);
  const ksMonths = (arr, a, b) => arr.filter(k => k.month && k.month >= a && k.month <= b);
  const planOf = (name, a, b) => (data.plan || []).filter(p => p.obj === name && p.month >= a && p.month <= b).reduce((s, p) => s + num(p.sum), 0);
  const cap = (o, beforeM) => !num(o.sum) ? INF : Math.max(0, +o.sum - num(o.done0) - sumKs(ksObj(o.name).filter(k => k.month && k.month >= BASE && k.month < beforeM)));

  // прошлый период
  let prevS, prevE;
  if (week) { prevS = addDays(B10, -7); prevE = addDays(B11, -7); }
  else { prevS = m1(addM(ym(B10), -({ month: 1, quarter: 3, year: 12 })[kind])); prevE = addDays(B10, -1); }
  const factPrev = week ? sumKs(ks.filter(k => has(k.ks2) && k.wdate && k.wdate >= prevS && k.wdate <= prevE))
                        : sumKs(ks.filter(k => k.month && m1(k.month) >= prevS && m1(k.month) <= prevE));

  // по объектам
  const rows = (data.objects || []).map(o => {
    const mine = ksObj(o.name);
    const planP = Math.min(planOf(o.name, planStart, planEnd), cap(o, planStart));
    const factP = sumKs(mine.filter(inP));
    const factPct = week ? sumKs(mine.filter(k => k.month === B13)) : factP;
    const planY = Math.min(planOf(o.name, yStart, ym(ytdEndDate)), cap(o, yStart));
    const factY = sumKs(ksMonths(mine, yStart, ym(ytdEndDate)));
    const doneEnd = num(o.done0) + sumKs(ksMonths(mine, BASE, ym(B11)));
    const remainEnd = (o.status === "Завершён" || num(o.sum) <= 0) ? 0 : +o.sum - doneEnd;
    const got = (data.pays || []).filter(p => p.obj === o.name && p.date >= B10 && p.date <= B11).reduce((a, p) => a + num(p.sum), 0);
    // прогноз года
    const BM = (data.plan || []).filter(p => p.obj === o.name && +p.month.slice(0, 4) === year && p.month > B20).reduce((s, p) => s + num(p.sum), 0);
    const BN = cap(o, addM(B20, 1));
    const BP = Math.min(planOf(o.name, yStart, year + "-12"), cap(o, yStart));
    const BQ = Math.min(planOf(o.name, yStart, B20 < year + "-12" ? B20 : year + "-12"), cap(o, yStart));
    const show = planP !== 0 || factP !== 0 || planY !== 0 || factY !== 0 || o.status === "В работе";
    return { name: o.name, resp: o.resp || "", status: o.status, planP, factP, factPct, pct: planP ? factPct / planP : null, dev: factPct - planP,
      planY, factY, pctY: planY ? factY / planY : null, sum: num(o.sum), doneEnd, remainEnd, pctC: num(o.sum) > 0 ? doneEnd / +o.sum : null,
      got, BO: Math.min(BM, BN), BP, BQ, show };
  });
  const tot = f => rows.reduce((a, r) => a + num(f(r)), 0);

  const E2 = tot(r => r.planP);
  const E3 = sumKs(ks.filter(inP));
  const E5 = week ? sumKs(ks.filter(k => k.month === B13)) : E3;
  const E10 = tot(r => r.planY);
  const E11 = sumKs(ksMonths(ks, yStart, ym(ytdEndDate)));
  const pKs = ks.filter(inP);
  const H10 = tot(r => r.BP), H11 = tot(r => r.BQ);
  const H12 = sumKs(ksMonths(ks, yStart, B20 < year + "-12" ? B20 : year + "-12"));
  const H13 = H11 === 0 ? 1 : H12 / H11;
  const H14 = H12 + tot(r => r.BO) * H13;
  const live = (data.objects || []).filter(o => o.status !== "Завершён");
  const portf = live.filter(o => num(o.sum) > 0).reduce((a, o) => a + +o.sum, 0);
  const E17 = rows.filter(r => r.status !== "Завершён" && r.sum > 0).reduce((a, r) => a + r.remainEnd, 0);

  // деньги
  const AO = k => k.expect ? (k.expect > today ? k.expect : today) : "";
  const curM = ym(today);
  const planM = m => (data.plan || []).filter(p => p.month === m).reduce((s, p) => s + num(p.sum), 0);
  const H2 = sumKs(ks.filter(k => AO(k) && AO(k) >= B10 && AO(k) <= B11), k => k.toPay);
  let H3 = 0;
  (data.plan || []).forEach(p => { const nx = m1(addM(p.month, 1)); if (nx >= m1(ym(B10)) && nx <= B11 && p.month > B20) H3 += num(p.sum); });
  H3 *= 0.95;
  const H5 = sumKs(ks.filter(k => AO(k) && ym(AO(k)) === curM), k => k.toPay) + (addM(curM, -1) > B20 ? planM(addM(curM, -1)) * 0.95 : 0);
  const H7 = sumKs(ks.filter(k => k.expect && k.expect < today), k => k.toPay);
  const pays = (data.pays || []).filter(p => p.date >= B10 && p.date <= B11);
  const gar = (data.gar || []).filter(g => g.status !== "Оплачена");

  const st = s => pKs.filter(k => k.st === s);
  const out = {
    period: P, kind, week, list, today, B20, year, commentMonth,
    planLabel: week ? "План месяца" : "План за период", factLabel: week ? "Факт за неделю (КС-2)" : "Факт за период (КС-2)",
    plan: E2, fact: E3, factPct: E5, factMonth: week ? E5 : null, pct: E2 ? E5 / E2 : 0, dev: E5 - E2,
    prevFact: factPrev, prevChange: (factPrev === 0 || prevS < m1(BASE)) ? null : E3 / factPrev - 1,
    ksCount: pKs.length, objPlanned: rows.filter(r => r.planP > 0).length, objLive: live.length,
    planY: E10, factY: E11, pctY: E10 ? E11 / E10 : 0, devY: E11 - E10,
    planYear: H10, rate: H13, forecast: H14, gap: H14 - H10, forecastPct: H10 ? H14 / H10 : 0,
    st: {
      review: { sum: sumKs(st("review")), n: st("review").length },
      agreed: { sum: sumKs(st("agreed")), n: st("agreed").length },
      signed: { sum: sumKs(st("signed"), k => k.toPay), ks2: sumKs(st("signed")), n: st("signed").length },
      paid: { sum: sumKs(st("paid")), n: st("paid").length }
    },
    money: {
      expectPeriod: H2 + H3, expectCur: H5, expectTile: B11 < today ? H5 : H2 + H3, expectTileCur: B11 < today,
      overdue: H7, got: pays.reduce((a, p) => a + num(p.sum), 0), adv: pays.filter(p => p.kind === "Аванс").reduce((a, p) => a + num(p.sum), 0),
      garIn: gar.filter(g => g.due && g.due >= B10 && g.due <= B11).reduce((a, g) => a + num(g.sum), 0),
      garInN: gar.filter(g => g.due && g.due >= B10 && g.due <= B11 && num(g.sum) > 0).length,
      garOver: gar.filter(g => g.due && g.due < B10).reduce((a, g) => a + num(g.sum), 0),
      garNoDate: gar.filter(g => !g.due).reduce((a, g) => a + num(g.sum), 0),
      garAll: gar.reduce((a, g) => a + num(g.sum), 0),
      remain: E17, portfolio: portf, portfolioPct: portf ? 1 - E17 / portf : 0
    },
    comment: ((data.notes || []).find(n => n.month === commentMonth) || {}).text || "",
    rows: rows.filter(r => r.show).sort((a, b) => Math.max(b.planP, b.factPct) - Math.max(a.planP, a.factPct) || b.planY - a.planY),
    ksList: pKs.slice().sort((a, b) => (b.date || b.wdate).localeCompare(a.date || a.wdate))
  };
  // по ответственным
  const by = {};
  rows.forEach(r => {
    const k = r.resp || "— не указан —";
    const g = by[k] || (by[k] = { resp: k, live: 0, planP: 0, factPct: 0, planY: 0, factY: 0, remain: 0 });
    if (r.status !== "Завершён") g.live++;
    g.planP += r.planP; g.factPct += r.factPct; g.planY += r.planY; g.factY += r.factY;
    if (r.status !== "Завершён" && r.sum > 0) g.remain += r.remainEnd;
  });
  out.byResp = Object.values(by).filter(g => g.live || g.planP || g.factPct || g.planY || g.factY)
    .map(g => Object.assign(g, { pct: g.planP ? g.factPct / g.planP : null, dev: g.factPct - g.planP, pctY: g.planY ? g.factY / g.planY : null }))
    .sort((a, b) => b.planY - a.planY || b.factY - a.factY);

  // графики: 12 месяцев года
  const cut = ym(B11 < eom(B20) ? B11 : eom(B20));
  const months = []; for (let i = 0; i < 12; i++) months.push(year + "-" + String(i + 1).padStart(2, "0"));
  const capPlan = m => (data.objects || []).reduce((acc, o) => {   // План_расч: план, урезанный остатком договора нарастающим итогом
    const c = cap(o, yStart); const a = planOf(o.name, yStart, addM(m, -1)), b = a + planOf(o.name, m, m);
    return acc + (Math.min(b, c) - Math.min(a, c));
  }, 0);
  let cp = 0, cf = 0;
  out.chartMonths = months.map(m => {
    const pl = capPlan(m), fa = m <= cut ? sumKs(ks.filter(k => k.month === m)) : null;
    cp += pl;
    const fc = m <= cut ? (cf += fa) : (cf += pl * H13);
    return { m, label: MNs[+m.slice(5) - 1], plan: pl, fact: fa, cumPlan: cp, cumFact: m <= cut ? fc : null, cumFc: m >= cut ? fc : null };
  });
  // ожидаемые поступления: просрочено + 6 месяцев вперёд
  const exp = [{ label: "Просрочено", sum: H7, over: true }];
  for (let i = 0; i < 6; i++) {
    const m = addM(curM, i);
    const byKs = sumKs(ks.filter(k => k.expect && k.expect >= today && ym(k.expect) === m), k => k.toPay);
    const byPlan = addM(m, -1) > B20 ? planM(addM(m, -1)) * 0.95 : 0;
    exp.push({ label: MNs[+m.slice(5) - 1] + " " + m.slice(2, 4), ks: byKs, plan: byPlan, sum: byKs + byPlan });
  }
  out.chartExpect = exp;
  out.topContracts = objectRows(data, today).filter(o => o.status !== "Завершён" && num(o.sum) > 0)
    .sort((a, b) => b.sum - a.sum).slice(0, 10).map(o => ({ name: o.name, done: o.doneAll, remain: Math.max(0, o.remain) }));
  return out;
}

// ---------- KPI: специалист ПТО, начальник ПТО, отдел ----------
// Плавная шкала: коэффициент показателя = достижение цели (0…120%), без порогов и обнулений.
// Показатель без данных за период не учитывается — его вес делится между остальными.
const KPI_SETS = {
  spec: [
    { k: "ksTime", n: "КС поданы в срок по графику", w: 0.30, t: 1, how: "КС за месяц с датой не позже последнего числа отчётного месяца / сколько КС положено по графику (1 или 2 в месяц) по объектам с планом или с работами по рапортам" },
    { k: "close", n: "Работы по рапортам закрыты КС", w: 0.25, t: 1, how: "Доля месяцев, когда по рапортам прорабов на объекте были работы и за этот месяц подана КС" },
    { k: "ds", n: "Допработы: ДС подан в срок", w: 0.20, t: 1, how: "КС с отметкой «Нужен ДС», по которым ДС подан заказчику не позже 30 дней от даты КС" },
    { k: "remarks", n: "КС без замечаний и возвратов", w: 0.15, t: 1, how: "Доля КС за период без отметки «Замечания» / «Возврат на исправление»" },
    { k: "disc", n: "Дисциплина данных и отчёт вовремя", w: 0.10, t: 1, how: "Среднее: у КС есть дата; у подписанных есть дата подписания; отчёт «Факт» за месяц отправлен до 5-го числа следующего месяца" }
  ],
  head: [
    { k: "plan", n: "Выполнение плана закрытия отделом", w: 0.30, t: 1, how: "Принятый факт (КС согласована / подписана / оплачена) / план периода по всем объектам" },
    { k: "review", n: "Скорость проверки отчётов", w: 0.15, t: 1, how: "Доля отчётов, по которым решение принято не позже 2 рабочих дней после отправки" },
    { k: "planOnTime", n: "План на месяц согласован до 1-го числа", w: 0.10, t: 1, how: "Доля объектов с планом на месяц, у которых план согласован до начала месяца" },
    { k: "overdue", n: "Работа с просроченной оплатой", w: 0.15, t: 1, how: "Доля КС, не оплаченных в срок по договору, по которым есть дата письма или претензии заказчику" },
    { k: "ds", n: "Допработы: ДС подан в срок (отдел)", w: 0.15, t: 1, how: "Как у специалистов, по всем объектам отдела" },
    { k: "reopen", n: "Качество согласования", w: 0.10, t: 0, how: "Доля согласованных отчётов, которые пришлось открыть заново (цель — 0)" },
    { k: "note", n: "«Главное за месяц» к 5-му числу", w: 0.05, t: 1, how: "Комментарий для генерального за месяц заполнен до 5-го числа следующего месяца" }
  ],
  dept: [
    { k: "plan", n: "Выполнение плана закрытия", w: 0.40, t: 1, how: "Принятый факт / план периода (план не больше остатка договора)" },
    { k: "ksTime", n: "КС поданы в срок по графику", w: 0.20, t: 1, how: "По всем объектам отдела" },
    { k: "close", n: "Работы по рапортам закрыты КС", w: 0.15, t: 1, how: "По всем объектам отдела" },
    { k: "hang", n: "Без «зависших» КС", w: 0.15, t: 0, how: "Доля КС-2, которые больше 30 дней не согласованы заказчиком (цель — 0)" },
    { k: "disc", n: "Дисциплина данных", w: 0.10, t: 1, how: "По всем объектам отдела" }
  ]
};
const accepted = k => k.agr === "Согласована" || k.sign === "Подписана" || k.pay === "Оплачена";
const yes = v => /^да$/i.test(String(v || "").trim());
const badRemark = v => /замеч|возврат/i.test(String(v || ""));
const msk = iso => iso ? new Date(Date.parse(iso) + 3 * 3600e3).toISOString().slice(0, 10) : "";
function workDays(a, b) { let n = 0; for (let d = addDays(a, 1); d <= b; d = addDays(d, 1)) { const w = new Date(D(d)).getUTCDay(); if (w !== 0 && w !== 6) n++; } return n; }   // рабочие дни после a до b включительно, пн–пт
function kpiSettings(data) {
  const s = { w: {}, fund: 0, share: {} };
  (data.kpi || []).forEach(r => {
    if (!r || !r.id) return; let v = r.value; try { v = JSON.parse(r.value); } catch (e) {}
    if (r.id.indexOf("share:") === 0) s.share[r.id.slice(6)] = +v || 0;
    else if (r.id.indexOf("w:") === 0) s.w[r.id.slice(2)] = +v;
    else if (r.id === "fund") s.fund = +v || 0;
  });
  return s;
}
/** opt: { today, kind, period, act: { "Объект": ["2026-09", …] } — месяцы с работами по рапортам (не обязательно) } */
function kpi(data, opt) {
  const today = opt.today, kind = opt.kind || "month", S = kpiSettings(data), act = opt.act || null;
  const byName = index(data);
  const ks = (data.ks || []).map(k => enrichKs(k, byName));
  const list = periods(kind, data, today);
  const P = list.find(x => x.key === opt.period) || defaultPeriod(kind, data, today);
  const a = ym(P.start), b = ym(P.end);
  const months = []; for (let m = a; m <= b; m = addM(m, 1)) months.push(m);
  const done = months.filter(m => eom(m) < today);                 // законченные месяцы периода
  const inP = k => k.month && k.month >= a && k.month <= b && num(k.ks2) > 0;
  const sum = (arr, f) => arr.reduce((s, x) => s + num(f ? f(x) : x.ks2), 0);
  const planRaw = (o, x, y) => (data.plan || []).filter(p => p.obj === o && p.month >= x && p.month <= y).reduce((s, p) => s + num(p.sum), 0);
  const cap = o => !num(o.sum) ? INF : Math.max(0, +o.sum - num(o.done0) - sum(ks.filter(k => k.obj === o.name && k.month && k.month >= BASE && k.month < a)));
  const reps = data.reports || [];
  const rep = (type, obj, m) => reps.find(r => r.type === type && r.obj === obj && r.month === m);
  const cmts = r => { try { return JSON.parse(r.comments || "[]"); } catch (e) { return []; } };
  const R = (num_, den, issues) => ({ num: num_, den, value: den ? num_ / den : null, issues: issues || [] });
  const objs = data.objects || [];

  // ---- показатели по набору объектов ----
  function ksTime(os) {           // КС в срок по графику
    let n = 0, d = 0; const iss = [];
    os.forEach(o => done.forEach(m => {
      const worked = planRaw(o.name, m, m) > 0 || (act && (act[o.name] || []).indexOf(m) >= 0);
      if (!worked) return;
      const need = /^Дважды/.test(o.sched || "") ? 2 : 1;
      const mine = ks.filter(k => k.obj === o.name && k.month === m && num(k.ks2) > 0);
      const ok = Math.min(need, mine.filter(k => k.date && k.date <= eom(m)).length);
      d += need; n += ok;
      if (ok < need) iss.push(o.name + ", " + monthLabel(m).toLowerCase() + ": " + (mine.length ? "КС подана позже срока или без даты" : "КС нет"));
    }));
    return R(n, d, iss);
  }
  function close(os) {            // работы по рапортам закрыты КС
    if (!act) return R(0, 0);
    let n = 0, d = 0; const iss = [];
    os.forEach(o => (act[o.name] || []).filter(m => done.indexOf(m) >= 0).forEach(m => {
      d++; if (ks.some(k => k.obj === o.name && k.month === m && num(k.ks2) > 0)) n++; else iss.push(o.name + ", " + monthLabel(m).toLowerCase() + ": по рапортам были работы, КС нет");
    }));
    return R(n, d, iss);
  }
  function ds(os) {               // ДС в срок
    const names = os.map(o => o.name); let n = 0, d = 0; const iss = [];
    ks.filter(k => names.indexOf(k.obj) >= 0 && inP(k) && yes(k.dsNeed)).forEach(k => {
      const due = addDays(k.date || eom(k.month), 30);
      if (k.dsDate && k.dsDate <= due) { n++; d++; }
      else if (k.dsDate || due < today) { d++; iss.push(k.obj + " · " + (k.num || "КС") + ": ДС " + (k.dsDate ? "подан " + ru(k.dsDate) + ", срок был " + ru(due) : "не подан, срок " + ru(due))); }
    });
    return R(n, d, iss);
  }
  function remarks(os) {
    const names = os.map(o => o.name), pk = ks.filter(k => names.indexOf(k.obj) >= 0 && inP(k)), bad = pk.filter(k => badRemark(k.remarks));
    return R(pk.length - bad.length, pk.length, bad.map(k => k.obj + " · " + (k.num || "КС") + ": " + k.remarks));
  }
  function disc(os) {
    const names = os.map(o => o.name), pk = ks.filter(k => names.indexOf(k.obj) >= 0 && inP(k)), iss = [], parts = [];
    if (pk.length) { const w = pk.filter(k => !k.date); parts.push(1 - w.length / pk.length); w.slice(0, 8).forEach(k => iss.push(k.obj + " · " + (k.num || "КС") + ": нет даты КС")); }
    const sg = pk.filter(k => k.sign === "Подписана");
    if (sg.length) { const w = sg.filter(k => !k.signDate); parts.push(1 - w.length / sg.length); w.slice(0, 8).forEach(k => iss.push(k.obj + " · " + (k.num || "КС") + ": нет даты подписания")); }
    let rn = 0, rd = 0;
    os.forEach(o => done.forEach(m => {
      if (!ks.some(k => k.obj === o.name && k.month === m && num(k.ks2) > 0)) return;
      const r = rep("fact", o.name, m); if (r && !r.sentAt && !r.decidedAt) return;      // перенесённые данные — не оцениваем
      rd++; const dl = addM(m, 1) + "-05";
      if (r && r.sentAt && msk(r.sentAt) <= dl) rn++; else iss.push(o.name + ", " + monthLabel(m).toLowerCase() + ": отчёт «Факт» " + (r && r.sentAt ? "отправлен " + ru(msk(r.sentAt)) : "не отправлен") + " (срок " + ru(dl) + ")");
    }));
    if (rd) parts.push(rn / rd);
    return { num: null, den: parts.length, value: parts.length ? parts.reduce((s, x) => s + x, 0) / parts.length : null, issues: iss };
  }
  // план и «зависшие» — только по законченным месяцам периода (текущий месяц ещё идёт)
  const da = done[0], db = done[done.length - 1], inDone = k => k.month && da && k.month >= da && k.month <= db && num(k.ks2) > 0;
  const capD = o => !num(o.sum) ? INF : Math.max(0, +o.sum - num(o.done0) - sum(ks.filter(k => k.obj === o.name && k.month && k.month >= BASE && k.month < da)));
  function plan(os) {             // выполнение плана закрытия
    if (!done.length) return R(0, 0);
    let pl = 0, fa = 0; const iss = [];
    os.forEach(o => {
      const p = Math.min(planRaw(o.name, da, db), capD(o)), f = sum(ks.filter(k => k.obj === o.name && inDone(k) && accepted(k)));
      pl += p; fa += f;
      if (p > 0 && f / p < 0.8) iss.push(o.name + ": план " + (p / 1e6).toFixed(1).replace(".", ",") + " млн, принято " + (f / 1e6).toFixed(1).replace(".", ",") + " млн (" + Math.round(f / p * 100) + "%)");
    });
    return { num: fa, den: pl, value: pl ? fa / pl : null, issues: iss };
  }
  function hang(os) {
    if (!done.length) return R(0, 0);
    const names = os.map(o => o.name), wd = k => k.date || eom(k.month), end = eom(db);
    const h = ks.filter(k => names.indexOf(k.obj) >= 0 && num(k.ks2) > 0 && k.month && k.month <= db && wd(k) <= addDays(end, -30) && !accepted(k));
    const base = sum(ks.filter(k => names.indexOf(k.obj) >= 0 && inDone(k)));
    if (!h.length && !base) return R(0, 0);
    return { num: sum(h), den: base, value: Math.min(1, sum(h) / Math.max(1, base)), issues: h.map(k => k.obj + " · " + (k.num || "КС") + ": не согласована с " + ru(wd(k))) };
  }
  // ---- начальник ----
  function review() {
    let n = 0, d = 0; const iss = [];
    reps.forEach(r => {
      if (!r.sentAt || !r.decidedAt) return; const s = msk(r.sentAt), dd = msk(r.decidedAt);
      if (ym(dd) < a || ym(dd) > b || dd < s) return;
      d++; const w = workDays(s, dd); if (w <= 2) n++; else iss.push((r.type === "plan" ? "План" : "Факт") + " · " + r.obj + " · " + monthLabel(r.month).toLowerCase() + ": решение через " + w + " раб. дн.");
    });
    return R(n, d, iss);
  }
  function planOnTime() {
    let n = 0, d = 0; const iss = [];
    months.forEach(m => objs.forEach(o => {
      if (planRaw(o.name, m, m) <= 0) return;
      const r = rep("plan", o.name, m); if (r && r.status === "Согласован" && !r.sentAt && !r.decidedAt) return;   // перенесённые
      if (m1(m) > today) return;                                                      // месяц ещё не начался
      d++; const ap = r ? cmts(r).filter(c => c.act === "approve").map(c => msk(c.at)).sort()[0] : "";
      if (ap && ap < m1(m)) n++; else iss.push(o.name + ", " + monthLabel(m).toLowerCase() + ": план " + (ap ? "согласован " + ru(ap) : "не согласован"));
    }));
    return R(n, d, iss);
  }
  function overdue() {
    const end = P.end < today ? P.end : today;
    const od = ks.filter(k => num(k.ks2) > 0 && k.pay !== "Оплачена" && k.expect && k.expect >= P.start && k.expect < end);   // стали просроченными в этом периоде
    const ok = od.filter(k => k.letterDate);
    return R(ok.length, od.length, od.filter(k => !k.letterDate).map(k => k.obj + " · " + (k.num || "КС") + ": оплата ожидалась " + ru(k.expect) + ", письма нет"));
  }
  function reopen() {
    let ap = 0, ro = 0; const iss = [];
    reps.forEach(r => cmts(r).forEach(c => {
      const dd = msk(c.at); if (ym(dd) < a || ym(dd) > b) return;
      if (c.act === "approve") ap++;
      if (c.act === "reopen") { ro++; iss.push((r.type === "plan" ? "План" : "Факт") + " · " + r.obj + " · " + monthLabel(r.month).toLowerCase() + ": открыт заново " + ru(dd)); }
    }));
    return ap || ro ? { num: ro, den: ap, value: Math.min(1, ro / Math.max(1, ap)), issues: iss } : R(0, 0);
  }
  function note() {
    let n = 0, d = 0; const iss = [];
    done.forEach(m => {
      d++; const x = (data.notes || []).find(z => z.month === m && String(z.text || "").trim()), dl = addM(m, 1) + "-05";
      if (x && (!x.upd || msk(x.upd) <= dl)) n++; else iss.push(monthLabel(m) + ": " + (x ? "заполнен " + ru(msk(x.upd)) : "не заполнен"));
    });
    return R(n, d, iss);
  }
  // ---- сборка ----
  const FN = { ksTime, close, ds, remarks, disc, plan, hang, review, planOnTime, overdue, reopen, note };
  function build(set, os) {
    const items = KPI_SETS[set].map(def => {
      const r = FN[def.k](os), wk = set + "." + def.k, w = (wk in S.w && isFinite(S.w[wk])) ? S.w[wk] : def.w;
      const ach = r.value === null ? null : (def.t === 0 ? Math.max(0, 1 - r.value) : r.value / def.t);
      return Object.assign({}, def, r, { w, ach, coef: ach === null ? null : Math.max(0, Math.min(1.2, ach)), issues: r.issues.slice(0, 30), more: Math.max(0, r.issues.length - 30) });
    });
    const ws = items.filter(i => i.coef !== null).reduce((s, i) => s + i.w, 0);
    return { items, coef: ws ? items.filter(i => i.coef !== null).reduce((s, i) => s + i.coef * i.w, 0) / ws : null };
  }
  const names = []; objs.forEach(o => { if (o.pto && names.indexOf(o.pto) < 0) names.push(o.pto); });
  const people = names.map(nm => { const os = objs.filter(o => o.pto === nm); return Object.assign({ name: nm, objects: os.length }, build("spec", os)); });
  const dept = build("dept", objs), head = build("head", objs);
  const n = months.length, fund = S.fund * n;
  const money = fund > 0 ? {
    fund, dept: dept.coef === null ? null : fund * dept.coef,
    people: people.map(p => { const sh = p.name in S.share ? S.share[p.name] : 0; return { name: p.name, share: sh, pay: p.coef === null ? 0 : fund * sh * p.coef }; }),
    headShare: S.share["@head"] || 0
  } : null;
  if (money) { money.head = head.coef === null ? 0 : fund * money.headShare * head.coef; money.shareSum = money.people.reduce((s, x) => s + x.share, 0) + money.headShare; }
  return { period: P, list, months: n, done: done.length, hasAct: !!act, dept, head, people, money, settings: S };
}

const api = { BASE, compute, kpi, kpiSettings, KPI_SETS, periods, defaultPeriod, objectRows, enrichKs, index, ksStatus, STATUS_LABEL, monthLabel, ru, ym, eom, addM, addDays, MN };
root.PTO = api;
if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);

