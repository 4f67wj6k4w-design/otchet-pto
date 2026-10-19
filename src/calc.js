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
