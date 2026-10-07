// Сверка расчётов с пересчитанным Excel: node tools/check.js data.json ref.json
const P = require("../src/calc.js"), fs = require("fs");
const data = JSON.parse(fs.readFileSync(process.argv[2])), ref = JSON.parse(fs.readFileSync(process.argv[3]));
const cases = [["month","2026-09"],["month","2026-04"],["quarter","2026-Q2"],["year","2026"],["week","2026-08-31"],["month",null]];
let bad = 0;
cases.forEach(([kind, period], i) => {
  const m = P.compute(data, { kind, period, today: "2026-10-07" }), r = ref[i];
  const pairs = { B10: m.period.start, B11: m.period.end, E2: m.plan, E3: m.fact, E5: m.factPct, E6: m.pct, E8: m.prevFact, E9: m.prevChange, E10: m.planY, E11: m.factY,
    E14: m.ksCount, E17: m.money.remain, E18: m.objLive, E19: m.money.portfolioPct, E20: m.money.garAll, E21: m.st.review.sum, E22: m.st.review.n, E23: m.st.agreed.sum, E24: m.st.agreed.n,
    E25: m.st.signed.sum, E26: m.st.signed.n, E28: m.st.paid.sum, E29: m.st.paid.n, E30: m.money.garIn, H4: m.money.expectPeriod, H5: m.money.expectCur, H6: m.money.expectTile, H7: m.money.overdue,
    H10: m.planYear, H13: m.rate, H14: m.forecast };
  const errs = [];
  for (const k in pairs) {
    const a = pairs[k], b = r[k];
    const ok = (typeof b === "number") ? Math.abs((a || 0) - b) <= Math.max(0.02, Math.abs(b) * 1e-9) : ((a ?? null) === (b ?? null) || (b === null && (a === null || a === "")));
    if (!ok) errs.push(k + ": мой " + a + " / Excel " + b);
  }
  bad += errs.length;
  console.log(kind, m.period.label, errs.length ? "РАСХОЖДЕНИЯ:\n  " + errs.join("\n  ") : "совпадает (" + Object.keys(pairs).length + " показателей)");
});
process.exit(bad ? 1 : 0);
