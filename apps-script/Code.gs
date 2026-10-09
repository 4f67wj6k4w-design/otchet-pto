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

// Ключ ПТО: просмотр и ввод данных. Ключ руководителя: только просмотр сводки.
const PTO_KEY = 'ВСТАВЬТЕ_КЛЮЧ_ПТО';
const DIRECTOR_KEY = 'ВСТАВЬТЕ_КЛЮЧ_РУКОВОДИТЕЛЯ';

// Листы и столбцы базы (порядок столбцов в таблице можно менять — ищутся по заголовку).
const SCHEMA = {"O": ["Объекты", [["id", "ID", "s"], ["name", "Объект", "s"], ["resp", "Ответственный", "s"], ["pto", "Специалист ПТО", "s"], ["status", "Статус", "s"], ["sum", "Сумма договора с ДС, руб", "n"], ["done0", "Выполнено до 01.01.2026, руб", "n"], ["ret", "% удержания", "p"], ["payDays", "Срок оплаты, дней", "n"], ["due", "Срок по договору", "d"], ["sched", "График подачи КС", "s"], ["note", "Примечание", "s"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "K": ["Журнал КС", [["id", "ID", "s"], ["date", "Дата КС", "d"], ["month", "Отчётный месяц", "m"], ["obj", "Объект", "s"], ["num", "№ КС / описание", "s"], ["period", "Период работ", "s"], ["ks2", "Выполнение КС-2, руб", "n"], ["ks3", "К оплате КС-3, руб", "n"], ["agr", "Согласование", "s"], ["sign", "Подписание", "s"], ["pay", "Оплата", "s"], ["payDate", "Дата оплаты", "d"], ["signDate", "Дата подписания", "d"], ["note", "Примечание", "s"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "P": ["Поступления", [["id", "ID", "s"], ["date", "Дата поступления", "d"], ["obj", "Объект", "s"], ["sum", "Сумма, руб", "n"], ["kind", "Вид поступления", "s"], ["basis", "№ КС / основание", "s"], ["note", "Примечание", "s"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "L": ["План", [["id", "ID", "s"], ["obj", "Объект", "s"], ["month", "Месяц", "m"], ["sum", "План, руб", "n"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "G": ["Гарантии", [["id", "ID", "s"], ["obj", "Объект", "s"], ["due", "Срок возврата", "d"], ["sum", "Сумма ГУ, руб", "n"], ["status", "Статус оплаты", "s"], ["party", "Контрагент", "s"], ["note", "Примечание", "s"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "N": ["Главное за месяц", [["id", "ID", "s"], ["month", "Месяц", "m"], ["text", "Комментарий начальника ПТО", "s"], ["upd", "Изменено", "s"], ["by", "Кто изменил", "s"]]], "S": ["Списки", [["resp", "Ответственные (руководители объектов)", "s"], ["pto", "Специалисты ПТО", "s"]]]};
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

function log_(items) {
  if (!items.length) return;
  let sh = ss_().getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss_().insertSheet(LOG_SHEET);
    sh.getRange(1, 1, 1, 6).setValues([['Когда', 'Кто', 'Лист', 'ID', 'Действие', 'Данные']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const now = new Date();
  const vals = items.map(function (x) { return [now, x[0] || '', SCHEMA[x[1]][0], String(x[2]), x[3], x[4] ? JSON.stringify(x[4]) : '']; });
  sh.getRange(sh.getLastRow() + 1, 1, vals.length, 6).setValues(vals);
}

/** Записать строки (новые — в конец, существующие — по ID). base — «Изменено» на момент открытия: если строку уже поменяли, вернётся конфликт. */
function upsert_(t, rows, by, force) {
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
    s.cols.forEach(function (c) { if (c[0] in r) line[s.hdr.indexOf(c[1])] = toCell_(r[c[0]], c[2]); });
    if (at) s.sh.getRange(at, 1, 1, width).setValues([line]);
    else fresh.push(line);
    logs.push([by, t, r.id, at ? 'изменено' : 'добавлено', r]);
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

function remove_(t, ids, by) {
  const s = sheet_(t), map = rowIndex_(s);
  const rows = ids.map(function (id) { return map[String(id)]; }).filter(Boolean).sort(function (a, b) { return b - a; });
  rows.forEach(function (r) { s.sh.deleteRow(r); });
  log_(ids.map(function (id) { return [by, t, id, 'удалено', null]; }));
  return rows.length;
}

/** Переименование объекта: имя меняется во всех листах. */
function renameObject_(from, to, by) {
  if (!from || !to || from === to) return;
  ['K', 'P', 'G', 'L'].forEach(function (t) {
    const s = sheet_(t), n = s.sh.getLastRow();
    if (n < 2) return;
    const oc = s.hdr.indexOf('Объект') + 1, ic = s.hdr.indexOf('ID') + 1;
    const objs = s.sh.getRange(2, oc, n - 1, 1).getValues(), ids = s.sh.getRange(2, ic, n - 1, 1).getValues();
    let ch = false;
    objs.forEach(function (v, i) {
      if (String(v[0]) === from) { v[0] = to; ch = true; if (t === 'L') ids[i][0] = String(ids[i][0]).replace(from + '|', to + '|'); }
    });
    if (ch) { s.sh.getRange(2, oc, n - 1, 1).setValues(objs); if (t === 'L') s.sh.getRange(2, ic, n - 1, 1).setValues(ids); }
  });
  log_([[by, 'O', from, 'переименовано → ' + to, null]]);
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

function doGet(e) {
  const p = (e && e.parameter) || {};
  const role = p.k && p.k === PTO_KEY ? 'pto' : (p.k && p.k === DIRECTOR_KEY ? 'dir' : '');
  if (!role) return out_({ ok: false, error: 'key' });
  if (p.a === 'ping') return out_({ ok: true, role: role });
  if (p.a === 'all') return out_({ ok: true, role: role, data: readAll_(), time: new Date().toISOString() });
  return out_({ ok: false, error: 'action' });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'json' }); }
  if (!b.k || b.k !== PTO_KEY) return out_({ ok: false, error: 'key' });
  if (b.t && !SCHEMA[b.t]) return out_({ ok: false, error: 'table' });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    let res = {};
    if (b.a === 'save') {
      if (b.t === 'O' && b.rename) renameObject_(b.rename.from, b.rename.to, b.by);
      res = upsert_(b.t, b.rows || [], b.by, !!b.force);
      if (b.del && b.del.length) res.deleted = remove_(b.t, b.del, b.by);
    } else if (b.a === 'del') {
      res.deleted = remove_(b.t, b.ids || [], b.by);
    } else if (b.a === 'lists') {
      saveLists_({ resp: (b.lists && b.lists.resp) || [], pto: (b.lists && b.lists.pto) || [] });
    } else return out_({ ok: false, error: 'action' });
    return out_(Object.assign({ ok: true }, res));
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
