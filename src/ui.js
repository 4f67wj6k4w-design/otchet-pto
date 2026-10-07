/* Общие помощники страниц ПТО: форматы, связь с базой, простые SVG-графики с подсказками */
(function (root) {
"use strict";
const U = {};
U.$ = id => document.getElementById(id);
U.el = (tag, cls, txt) => { const e = document.createElement(tag); if (cls) e.className = cls; if (txt != null) e.textContent = txt; return e; };
U.sget = k => { try { return localStorage.getItem(k); } catch (e) { return null; } };
U.sset = (k, v) => { try { localStorage.setItem(k, v); } catch (e) {} };
U.todayISO = () => { const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 10); };
U.ru = s => s ? s.slice(8, 10) + "." + s.slice(5, 7) + "." + s.slice(0, 4) : "";
U.rub = v => (v === null || v === undefined || v === "" || isNaN(v)) ? "—" : Number(v).toLocaleString("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
U.rub0 = v => (v === null || v === undefined || v === "" || isNaN(v)) ? "—" : Math.round(Number(v)).toLocaleString("ru-RU");
U.mln = (v, d) => (v === null || v === undefined || isNaN(v)) ? "—" : (Number(v) / 1e6).toLocaleString("ru-RU", { minimumFractionDigits: d === undefined ? 1 : d, maximumFractionDigits: d === undefined ? 1 : d });
U.pct = (v, d) => (v === null || v === undefined || isNaN(v) || !isFinite(v)) ? "—" : (v * 100).toLocaleString("ru-RU", { maximumFractionDigits: d === undefined ? 0 : d }) + "%";
U.parseNum = s => { if (s === null || s === undefined) return null; s = String(s).replace(/\s| /g, "").replace(",", "."); if (s === "") return null; const n = Number(s); return isNaN(n) ? NaN : n; };
U.pctClass = v => v === null || v === undefined || isNaN(v) ? "" : (v < 0.8 ? "bad" : (v < 1 ? "warnc" : "ok"));
U.uid = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

// ---------- база ----------
U.db = function (storeKey) {
  const h = new URLSearchParams(location.hash.slice(1));
  let db = null;
  if (h.get("u") && h.get("k")) { db = { u: h.get("u"), k: h.get("k") }; U.sset(storeKey, JSON.stringify(db)); }
  else { try { db = JSON.parse(U.sget(storeKey) || "null"); } catch (e) {} }
  // адрес с ключом остаётся в строке: «На экран Домой» сохраняет текущий адрес
  if (db && !(h.get("u") && h.get("k"))) { try { history.replaceState(null, "", location.pathname + "#u=" + encodeURIComponent(db.u) + "&k=" + encodeURIComponent(db.k)); } catch (e) {} }
  return db;
};
const ERR = { key: "неверный ключ в ссылке", action: "обновите скрипт базы (новая версия развертывания)", json: "ошибка данных" };
U.get = async function (db, params) {
  const u = db.u + (db.u.includes("?") ? "&" : "?") + new URLSearchParams(Object.assign({ k: db.k, t: Date.now() }, params));
  const r = await fetch(u, { redirect: "follow" });
  const j = await r.json();
  if (!j.ok) throw new Error(ERR[j.error] || j.error || "ошибка базы");
  return j;
};
U.post = async function (db, body) {
  // text/plain — без предварительного CORS-запроса, Apps Script его не умеет
  const r = await fetch(db.u, { method: "POST", body: JSON.stringify(Object.assign({ k: db.k }, body)), redirect: "follow", headers: { "Content-Type": "text/plain;charset=utf-8" } });
  const j = await r.json();
  if (!j.ok) throw new Error(ERR[j.error] || j.error || "ошибка базы");
  return j;
};

// ---------- графики ----------
const NS = "http://www.w3.org/2000/svg";
const svgEl = (tag, attrs) => { const e = document.createElementNS(NS, tag); for (const k in attrs) e.setAttribute(k, attrs[k]); return e; };
function niceMax(v) { if (v <= 0) return 1; const p = Math.pow(10, Math.floor(Math.log10(v))); const n = v / p; return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * p; }
function tipBox(host) {
  let t = host.querySelector(".tip");
  if (!t) { t = U.el("div", "tip"); t.hidden = true; host.appendChild(t); }
  return t;
}
function showTip(host, x, title, rows) {
  const t = tipBox(host); t.innerHTML = ""; t.appendChild(U.el("b", null, title));
  rows.forEach(r => { const l = U.el("div", "tr"); const sw = U.el("i"); if (r.dash) { sw.className = "dash ln"; sw.style.color = r.color; } else sw.style.background = r.color; l.appendChild(sw); l.appendChild(U.el("span", null, r.name)); l.appendChild(U.el("em", null, r.val)); t.appendChild(l); });
  t.hidden = false;
  const w = host.clientWidth, tw = t.offsetWidth;
  t.style.left = Math.max(0, Math.min(w - tw, x - tw / 2)) + "px";
}
function legend(host, series) {
  const lg = U.el("div", "legend");
  series.forEach(s => { const it = U.el("span"); const sw = U.el("i"); if (s.dash) { sw.className = "dash"; sw.style.color = s.color; } else sw.style.background = s.color; if (s.line) sw.className = (sw.className + " ln").trim(); it.appendChild(sw); it.appendChild(document.createTextNode(s.name)); lg.appendChild(it); });
  host.appendChild(lg);
}
/** Столбцы по категориям. series: [{name, color, values[], kind:'bar'|'line', dash}] */
U.chart = function (host, cfg) {
  host.innerHTML = ""; host.classList.add("chart");
  const W = Math.max(280, host.clientWidth || 600), H = cfg.height || 220, padL = 44, padR = 8, padT = 10, padB = 24;
  const cats = cfg.cats, n = cats.length, bars = cfg.series.filter(s => s.kind !== "line"), lines = cfg.series.filter(s => s.kind === "line");
  let mx = 0; cfg.series.forEach(s => s.values.forEach(v => { if (v != null && v > mx) mx = v; }));
  mx = niceMax(mx * 1.05);
  const iw = W - padL - padR, ih = H - padT - padB, band = iw / n;
  const y = v => padT + ih - (v / mx) * ih;
  const svg = svgEl("svg", { viewBox: `0 0 ${W} ${H}`, width: W, height: H, role: "img", "aria-label": cfg.title || "" });
  for (let i = 0; i <= 4; i++) {
    const v = mx * i / 4, yy = y(v);
    svg.appendChild(svgEl("line", { x1: padL, x2: W - padR, y1: yy, y2: yy, class: i ? "grid" : "axis" }));
    const t = svgEl("text", { x: padL - 6, y: yy + 4, class: "tick", "text-anchor": "end" }); t.textContent = (cfg.fmtAxis || U.mln)(v, 0); svg.appendChild(t);
  }
  const every = band < 30 ? Math.ceil(30 / band) : 1;
  cats.forEach((c, i) => { if (i % every) return; const t = svgEl("text", { x: padL + band * i + band / 2, y: H - 6, class: "tick", "text-anchor": "middle" }); t.textContent = c; svg.appendChild(t); });
  const bw = Math.min(24, (band - 6) / Math.max(1, bars.length) - 2);
  bars.forEach((s, j) => s.values.forEach((v, i) => {
    if (v == null || v <= 0) return;
    const x = padL + band * i + band / 2 - (bars.length * (bw + 2) - 2) / 2 + j * (bw + 2), top = y(v), h = padT + ih - top, r = Math.min(4, h, bw / 2);
    const col = (s.colors && s.colors[i]) || s.color;
    svg.appendChild(svgEl("path", { d: `M${x},${top + h}V${top + r}Q${x},${top} ${x + r},${top}H${x + bw - r}Q${x + bw},${top} ${x + bw},${top + r}V${top + h}Z`, fill: col }));
  }));
  lines.forEach(s => {
    let d = "", pen = false;
    s.values.forEach((v, i) => { if (v == null) { pen = false; return; } const xx = padL + band * i + band / 2; d += (pen ? "L" : "M") + xx + "," + y(v); pen = true; });
    svg.appendChild(svgEl("path", { d, fill: "none", stroke: s.color, "stroke-width": 2, "stroke-linejoin": "round", "stroke-linecap": "round", "stroke-dasharray": s.dash ? "5 4" : "none" }));
    s.values.forEach((v, i) => { if (v == null) return; const last = i === s.values.length - 1 || s.values[i + 1] == null; if (last) svg.appendChild(svgEl("circle", { cx: padL + band * i + band / 2, cy: y(v), r: 4, fill: s.color, class: "dot" })); });
  });
  const hl = svgEl("rect", { x: 0, y: padT, width: band, height: ih, class: "hl", visibility: "hidden" }); svg.insertBefore(hl, svg.firstChild);
  host.appendChild(svg);
  if (cfg.series.length > 1) legend(host, cfg.series.map(s => ({ name: s.name, color: s.color, dash: s.dash, line: s.kind === "line" })));
  const fv = cfg.fmtTip || (v => U.mln(v) + " млн");
  const pick = ev => {
    const r = svg.getBoundingClientRect(); const px = (ev.clientX - r.left) * (W / r.width);
    const i = Math.floor((px - padL) / band); if (i < 0 || i >= n) { hide(); return; }
    hl.setAttribute("x", padL + band * i); hl.setAttribute("visibility", "visible");
    const rows = cfg.series.filter(s => s.values[i] != null).map(s => ({ name: s.name, color: (s.colors && s.colors[i]) || s.color, dash: s.dash, val: fv(s.values[i]) }));
    if (cfg.extraTip) rows.push(...cfg.extraTip(i));
    showTip(host, (padL + band * i + band / 2) * (r.width / W), cfg.tipTitle ? cfg.tipTitle(i) : cats[i], rows);
  };
  const hide = () => { hl.setAttribute("visibility", "hidden"); tipBox(host).hidden = true; };
  svg.addEventListener("pointermove", pick); svg.addEventListener("pointerdown", pick); svg.addEventListener("pointerleave", hide);
};
/** Горизонтальные полосы «выполнено / остаток» (стек), по строке на категорию */
U.hbars = function (host, cfg) {
  host.innerHTML = ""; host.classList.add("chart");
  const mx = Math.max(1, ...cfg.rows.map(r => cfg.series.reduce((a, s) => a + (r[s.key] || 0), 0)));
  const list = U.el("div", "hb");
  cfg.rows.forEach(r => {
    const row = U.el("div", "hb-r"); row.appendChild(U.el("div", "hb-n", r.name));
    const track = U.el("div", "hb-t");
    cfg.series.forEach(s => { const v = r[s.key] || 0; if (v <= 0) return; const seg = U.el("i"); seg.style.width = (v / mx * 100) + "%"; seg.style.background = s.color; track.appendChild(seg); });
    row.appendChild(track);
    row.appendChild(U.el("div", "hb-v", cfg.label(r)));
    row.title = cfg.series.map(s => s.name + ": " + U.mln(r[s.key]) + " млн").join("\n");
    list.appendChild(row);
  });
  host.appendChild(list);
  legend(host, cfg.series);
};
root.U = U;
})(window);
