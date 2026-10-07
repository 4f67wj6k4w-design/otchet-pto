// Тестовый сервер: исполняет apps-script/Code.gs с имитацией Google Таблицы и раздаёт страницы.
// node tools/mockserver.js data.json [порт]   → http://localhost:8080/index.html#u=http://localhost:8080/exec&k=pto
const http = require("http"), fs = require("fs"), path = require("path"), vm = require("vm");
const ROOT = path.join(__dirname, ".."), data = JSON.parse(fs.readFileSync(process.argv[2])), PORT = +process.argv[3] || 8080;
const schema = JSON.parse(fs.readFileSync(path.join(ROOT, "src/schema.json")));
const pad = n => String(n).padStart(2, "0");
const toD = s => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d || 1); };

function Sheet(name) { this.name = name; this.rows = []; this.fmt = {}; }
Sheet.prototype = {
  getLastRow() { let n = this.rows.length; while (n && this.rows[n - 1].every(v => v === "" || v == null)) n--; return n; },
  getLastColumn() { return this.rows.reduce((a, r) => { let n = r.length; while (n && (r[n - 1] === "" || r[n - 1] == null)) n--; return Math.max(a, n); }, 0); },
  getMaxRows() { return Math.max(1000, this.rows.length); },
  getRange(r, c, nr, nc) { return new Range(this, r, c, nr || 1, nc || 1); },
  deleteRow(r) { this.rows.splice(r - 1, 1); },
  appendRow(v) { const n = this.getLastRow(); this.rows.splice(n, 0, v.slice()); },
  insertRowsAfter() {}, setFrozenRows() {}
};
function Range(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
const chain = function () { return this; };
Range.prototype = {
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = this.sh.rows[this.r - 1 + i] || []; const o = []; for (let j = 0; j < this.nc; j++) { const v = row[this.c - 1 + j]; o.push(v === undefined || v === null ? "" : v); } out.push(o); } return out; },
  getValue() { return this.getValues()[0][0]; },
  setValues(v) { for (let i = 0; i < this.nr; i++) { const ri = this.r - 1 + i; while (this.sh.rows.length <= ri) this.sh.rows.push([]); for (let j = 0; j < this.nc; j++) this.sh.rows[ri][this.c - 1 + j] = v[i][j]; } return this; },
  setValue(v) { return this.setValues([[v]]); },
  clearContent() { for (let i = 0; i < this.nr; i++) { const row = this.sh.rows[this.r - 1 + i]; if (row) for (let j = 0; j < this.nc; j++) row[this.c - 1 + j] = ""; } return this; },
  setNumberFormat: chain, setFontWeight: chain, setBackground: chain, setFontColor: chain, setWrap: chain
};
const sheets = {};
const ss = {
  getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new Sheet(n)), getSpreadsheetTimeZone: () => "Europe/Moscow",
  getSheets: () => Object.values(sheets), deleteSheet() {}
};
// наполнение из data.json — как после загрузки База_ПТО.xlsx
const src = { O: data.objects, K: data.ks, P: data.pays, L: data.plan, G: data.gar, N: data.notes };
for (const t in schema) {
  const [name, cols] = schema[t]; const sh = ss.insertSheet(name); sh.rows.push(cols.map(c => c[1]));
  let rows = src[t];
  if (t === "S") { const n = Math.max(data.lists.resp.length, data.lists.pto.length); rows = []; for (let i = 0; i < n; i++) rows.push({ resp: data.lists.resp[i] || "", pto: data.lists.pto[i] || "" }); }
  (rows || []).forEach(r => sh.rows.push(cols.map(([f, , typ]) => { const v = r[f]; if (v == null || v === "") return ""; if (typ === "d") return toD(v); if (typ === "m") return toD(v + "-01"); return v; })));
}
const ctx = {
  SpreadsheetApp: { getActiveSpreadsheet: () => ss },
  ContentService: { createTextOutput: s => ({ s, setMimeType() { return this; } }), MimeType: { JSON: "json" } },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Utilities: {
    formatDate: (d, tz, f) => f === "yyyy-MM" ? d.getFullYear() + "-" + pad(d.getMonth() + 1) : d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()),
    parseDate: (s) => toD(s)
  },
  console
};
vm.createContext(ctx);
const code = fs.readFileSync(path.join(ROOT, "apps-script/Code.gs"), "utf8").replace("ВСТАВЬТЕ_КЛЮЧ_ПТО", "pto").replace("ВСТАВЬТЕ_КЛЮЧ_РУКОВОДИТЕЛЯ", "dir");
vm.runInContext(code, ctx);
const MIME = { ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".css": "text/css", ".json": "application/json" };
http.createServer((req, res) => {
  const u = new URL(req.url, "http://x");
  if (u.pathname === "/exec") {
    if (req.method === "GET") { const out = ctx.doGet({ parameter: Object.fromEntries(u.searchParams) }); res.writeHead(200, { "Content-Type": "application/json" }); res.end(out.s); return; }
    let body = ""; req.on("data", c => body += c); req.on("end", () => { const out = ctx.doPost({ postData: { contents: body } }); res.writeHead(200, { "Content-Type": "application/json" }); res.end(out.s); }); return;
  }
  if (u.pathname === "/dump") { res.writeHead(200, { "Content-Type": "application/json" }); res.end(JSON.stringify(Object.fromEntries(Object.entries(sheets).map(([k, s]) => [k, s.rows])))); return; }
  const f = path.join(ROOT, u.pathname === "/" ? "index.html" : u.pathname);
  if (!f.startsWith(ROOT) || !fs.existsSync(f)) { res.writeHead(404); res.end("404"); return; }
  res.writeHead(200, { "Content-Type": MIME[path.extname(f)] || "application/octet-stream" }); fs.createReadStream(f).pipe(res);
}).listen(PORT, () => console.log("http://localhost:" + PORT));
