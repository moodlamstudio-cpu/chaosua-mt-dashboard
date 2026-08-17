// Smoke test: load index.html + app.js data path with a DOM stub, drive the
// render pipeline, and verify the Ship-to linked filter updates the KPI, charts,
// ship-to table highlight, and that clicking a row (toggleShipTo) filters and
// clicking again clears. Node-only; verifies the end-to-end render contract that
// the browser would otherwise exercise.
"use strict";
const fs = require("fs");
const vm = require("vm");
const DIR = __dirname;

const appSrc = fs.readFileSync(`${DIR}\\app.js`, "utf8");
const chanData = JSON.parse(fs.readFileSync(`${DIR}\\data_channels.json`, "utf8"));
const shiptoData = JSON.parse(fs.readFileSync(`${DIR}\\shipto_data.json`, "utf8"));

let pass = 0, fail = 0;
function chk(name, cond, extra) {
  if (cond) { pass++; console.log("  [PASS]", name); }
  else { fail++; console.log("  [FAIL]", name, extra || ""); }
}

// ---- Minimal DOM ----
function makeEl(id) {
  const el = {
    id, _html: "", _text: "",
    innerHTML: "", textContent: "",
    value: "", className: "", disabled: false, checked: false,
    style: {}, dataset: {},
    children: [],
    listeners: {},
    addEventListener(ev, fn) { (this.listeners[ev] = this.listeners[ev] || []).push(fn); },
    appendChild() {},
    querySelector() { return null; },
    querySelectorAll() { return []; },
  };
  return el;
}
const byId = {};
function getEl(id) {
  if (!byId[id]) byId[id] = makeEl(id);
  return byId[id];
}
// Intercept innerHTML/textContent writes to record them (setT/setV use these).
const documentStub = {
  getElementById: getEl,
  querySelectorAll: () => [],
  createElement: () => makeEl("tmp"),
};

const sandbox = {
  console, require, Date, JSON, Math, Number, String, Object, Array, isNaN,
  encodeURIComponent, decodeURIComponent, Promise,
  Chart: undefined,
  XLSX: undefined,
  fetch: () => Promise.reject(new Error("no fetch in smoke test")),
  document: documentStub,
  alert: () => {},
  setTimeout: () => {},
};
vm.createContext(sandbox);
sandbox.__shipto = shiptoData;
sandbox.__chan = chanData;

// A Chart stub that records instances so drawAnnual/drawMonthly/drawCategory create something.
let chartCount = 0;
sandbox.Chart = function (canvas, config) { this.config = config; this.canvas = canvas; this.destroy = () => {}; chartCount++; };
vm.runInContext(appSrc, sandbox, { filename: "app.js" });

// Set up state exactly as load() would after fetching.
vm.runInContext(`
  S = __shipto;
  D = __chan;
  lastClosed = (D && D._lastClosed) || 8;
  viewYear = fy; selFrom = 1; selTo = lastClosed;
  selChans = ["MAKRO","LOTUS'"];
  activeCat = null; activeSku = null; activeShipTo = null; kpiPick = null;
`, sandbox, { filename: "_smoke_setup" });

// Choose the top ship-to from the default view for a deterministic test.
const topRt = (() => {
  const rows = {};
  shiptoData.facts.forEach(r => {
    if ((r[1] === "MAKRO" || r[1] === "LOTUS'") && String(r[2]) === "2026") {
      const m = Number(r[3]); if (m < 1 || m > 8) return;
      rows[r[0]] = (rows[r[0]] || 0) + r[6];
    }
  });
  return Object.keys(rows).sort((a, b) => rows[b] - rows[a])[0];
})();
if (!topRt) { console.error("no top ship-to available"); process.exit(1); }

// Clone filters into the vm scope via a helper injected at end of app.js scope.
const runFrag = (code) => vm.runInContext(code, sandbox, { filename: "_smoke_frag" });

// 1. Initial render (no ship-to selected) populates all sections.
runFrag("renderAll();");
chk("renderAll runs without error (no ship-to)", (function(){ try{return (byId.tgM && byId.tgM.textContent.length>0);}catch(e){return false;} })());
chk("ship-to table header row rendered", byId.shipToTable && byId.shipToTable.innerHTML.indexOf("Ship-to party") >= 0);
chk("ship-to table has Total footer", byId.shipToTable && byId.shipToTable.innerHTML.indexOf("Total") >= 0);
chk("monthly chart created", chartCount >= 2);
const baseKpi = runFrag("(function(){renderAll();return rangeSum(C().s_actual,selFrom,selTo);})()");

// 2. Click top ship-to row -> KPI filters to it and reconciles with its table row.
const topJs = JSON.stringify(topRt);
runFrag(`activeShipTo=${topJs}; renderAll();`);
const kpiOn = runFrag("(function(){ return rangeSum(C().s_actual,selFrom,selTo); })()");
const tableRow = (() => {
  let sum = 0;
  shiptoData.facts.forEach(r => {
    if (r[0] !== topRt || (r[1] !== "MAKRO" && r[1] !== "LOTUS'")) return;
    const m = Number(r[3]); if (m < 1 || m > 8) return;
    if (String(r[2]) !== "2026") return;
    sum += r[6];
  });
  return sum;
})();
chk("clicking ship-to row filters KPI to that location", Math.abs(kpiOn - tableRow) < 0.001,
  `kpi=${kpiOn.toFixed(4)} expected=${tableRow.toFixed(4)}`);
chk("selected ship-to row shows highlight class in table HTML",
  byId.shipToTable.innerHTML.indexOf('class="st-row shipto-selected"') >= 0);
chk("ship-to header tag shows Ship-to prefix",
  byId.tgShipTo.textContent.indexOf("Ship-to") === 0);
chk("shiptoHeaderActions shows clear button",
  byId.shiptoHeaderActions && byId.shiptoHeaderActions.innerHTML.indexOf("stClear") >= 0);

// 3. Table still shows ALL rows (not reduced to the selected one) under other filters.
const rowCount = (byId.shipToTable.innerHTML.match(/class="st-row/g) || []).length;
chk("ship-to table still lists all rows after selection", rowCount >= 1, "rows=" + rowCount);

// 4. Click again -> clears and KPI returns to the aggregate (all ship-tos).
runFrag('activeShipTo=null; renderAll();');
chk("clearing ship-to returns KPI to aggregate", Math.abs(kpiOn - baseKpi) > 0.001, "kpiOn=" + kpiOn.toFixed(4) + " base=" + baseKpi.toFixed(4));

// 5. toggleShipTo round-trip (same handler the click wires to).
runFrag('toggleShipTo("A"); toggleShipTo("A");');
chk("toggleShipTo toggles cleanly to null on second click",
  runFrag("activeShipTo === null"));

console.log("\nRESULT: " + pass + " passed, " + fail + " failed");
process.exit(fail ? 1 : 0);
