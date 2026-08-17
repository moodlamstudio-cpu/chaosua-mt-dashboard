// Node harness to validate the Ship-to linked-filter logic that lives inside
// app.js. It loads app.js with a DOM stub into a vm context, then runs the
// assertions inside that SAME context so the test shares the module-scope
// `S`, `viewYear`, `selChans`, `activeShipTo`, ... bindings with the code it
// verifies. All figures come from app.js pure aggregation helpers, and are
// cross-checked against an independent re-computation of the drawShipTo rows /
// KPI from shipto_data.json, so the "selected row == Actual KPI" contract that
// the linked filter depends on is verified mechanically.
"use strict";
const fs = require("fs");
const vm = require("vm");
const DIR = __dirname;

const src = fs.readFileSync(`${DIR}\\app.js`, "utf8");
const shiptoData = JSON.parse(fs.readFileSync(`${DIR}\\shipto_data.json`, "utf8"));

// DOM stub sufficient for app.js's top-level declarations and guarded calls.
const noEl = {
  querySelectorAll: () => [],
  addEventListener: () => {},
  appendChild: () => {},
  setAttribute: () => {},
};
const sandbox = {
  console, fs, require, Date, JSON, Math, Number, String, Object, Array, isNaN,
  encodeURIComponent, decodeURIComponent,
  Chart: undefined,
  XLSX: undefined,
  fetch: () => Promise.reject(new Error("no fetch in test")),
  document: {
    getElementById: () => noEl,
    querySelectorAll: () => [],
    createElement: () => noEl,
    body: {},
  },
  setTimeout: (fn) => { /* swallow */ },
};
// Assertion reporter shared with the vm context.
sandbox.__test = { checks: 0, failures: 0 };
vm.createContext(sandbox);
vm.runInContext("var __t=function(name,pass,detail){__test.checks++;if(!pass){__test.failures++;console.error('  FAIL  '+name+(detail?'  ('+detail+')':''));}else{console.log('  ok    '+name);}};", sandbox, { filename: "_helpers" });
vm.runInContext(src, sandbox, { filename: "app.js" });
sandbox.__shipto = shiptoData;
vm.runInContext("S=__shipto; viewYear=fy; selFrom=1; selTo=lastClosed; selChans=['MAKRO',\"LOTUS'\"]; activeCat=null; activeSku=null; activeShipTo=null; kpiPick=null;", sandbox, { filename: "_setup" });

const testScript = `
(function(){
  "use strict";
  // Recompute the ship-to table (mirror of drawShipTo aggregation, BEFORE any
  // ship-to is selected) so we have an independent reference of row totals.
  function tableRows(from,to,year,chans,cat,sku){
    from=Number(from);to=Number(to);year=String(year);cat=cat||null;sku=sku||null;
    var set={};chans.forEach(function(id){set[id]=1;});
    var rows={},grand=0;
    (S.facts||[]).forEach(function(r){
      if(!set[r[1]])return;
      if(String(r[2])!==year)return;
      var m=Number(r[3]);if(m<from||m>to)return;
      if(cat&&r[4]!==cat)return;
      if(sku&&r[5]!==sku)return;
      rows[r[0]]=((rows[r[0]])||0)+num(r[6]);
      grand+=(num(r[6]));
    });
    return {rows:rows,grand:grand};
  }
  function kpiActualFor(name,from,to,year,cat,sku){
    var prev={activeShipTo:activeShipTo,selFrom:selFrom,selTo:selTo,viewYear:viewYear,activeCat:activeCat,activeSku:activeSku};
    activeShipTo=name;selFrom=from;selTo=to;viewYear=year;activeCat=cat||null;activeSku=sku||null;
    var ch=shipToChannel(name);
    var a=0;for(var m=selFrom;m<=selTo;m++)a+=num(ch.monthly[String(m)]);
    activeShipTo=prev.activeShipTo;selFrom=prev.selFrom;selTo=prev.selTo;viewYear=prev.viewYear;activeCat=prev.activeCat;activeSku=prev.activeSku;
    return {a:a,monthly:ch.monthly};
  }

  function reconcile1(clazz,label,name,from,to,year,chans,cat,sku){
    var ref=tableRows(from,to,year,chans,cat,sku);
    var top=Object.keys(ref.rows).sort(function(a,b){return ref.rows[b]-ref.rows[a];})[0];
    if(!top){__t(clazz+" has a top row",false,"no rows");return;}
    var kpi=kpiActualFor(top,from,to,year/*,cat,sku*/);
    var refRow=ref.rows[top];
    __t(clazz+": "+label+" ["+top.slice(0,20)+"...]",
      Math.abs(kpi.a-refRow)<0.001,
      "tableRow="+refRow.toFixed(4)+" kpi="+kpi.a.toFixed(4));
  }

  // Reconcile the Actual KPI with the ship-to table row when BOTH a ship-to
  // and a category (and optionally a SKU) are active. The table rows narrow to
  // the same cat/sku, so the selected row must equal the KPI.
  function reconcileCatSku(clazz,label,_,from,to,year,chans,cat,sku){
    var ref=tableRows(from,to,year,chans,cat,sku||null);
    var top=Object.keys(ref.rows).sort(function(a,b){return ref.rows[b]-ref.rows[a];})[0];
    if(!top){__t(clazz+": "+label+" no rows for cat/sku",false,"");return;}
    var prev={activeShipTo:activeShipTo,selFrom:selFrom,selTo:selTo,viewYear:viewYear,activeCat:activeCat,activeSku:activeSku};
    activeShipTo=top;selFrom=from;selTo=to;viewYear=year;activeCat=cat||null;activeSku=sku||null;
    var ch=shipToChannel(top);
    var kpi=0;for(var m=selFrom;m<=selTo;m++)kpi+=num(ch.monthly[String(m)]);
    activeShipTo=prev.activeShipTo;selFrom=prev.selFrom;selTo=prev.selTo;viewYear=prev.viewYear;activeCat=prev.activeCat;activeSku=prev.activeSku;
    __t(clazz+": "+label+" cat/sku drill reconciles ["+top.slice(0,20)+"...]",
      Math.abs(kpi-ref.rows[top])<0.001,
      "ref="+ref.rows[top].toFixed(4)+" kpi="+kpi.toFixed(4));
  }

  var C1=["MAKRO","LOTUS'"], C2=["MAKRO"], C3=["MAKRO","LOTUS'","7Eleven","BigC"];

  console.log("Ship-to linked filter reconciliation");
  reconcile1("default","2026 MAKRO+LOTUS' Jan-Aug",null,1,8,2026,C1,null,null);
  reconcile1("single-channel","2026 MAKRO only",null,1,8,2026,C2,null,null);
  reconcile1("multi-channel","2026 4 channels",null,1,8,2026,C3,null,null);
  reconcile1("range Jan-Jun","2026 MAKRO+LOTUS' Jan-Jun",null,1,6,2026,C1,null,null);
  reconcile1("historical-year","2024 view",null,1,8,2024,C1,null,null);
  reconcile1("historical-year Q3","2023 Jul-Dec",null,7,12,2023,C1,null,null);

  // Category + SKU drill filters with a ship-to selected.
  reconcileCatSku("drill","2026 cat filter",null,1,8,2026,C1,"หมูหยอง",null);
  reconcileCatSku("drill","2026 cat+sku filter",null,1,8,2026,C1,"หมูหยอง","หมูหยอง 70 ก.x30");

  // Channel-card split: sum of byChan must equal the ship-to KPI for the year.
  (function(){
    var prev={activeShipTo:activeShipTo,selFrom:selFrom,selTo:selTo,viewYear:viewYear,activeCat:activeCat,activeSku:activeSku,selChans:selChans};
    selChans=C1;selFrom=1;selTo=8;viewYear=2026;activeCat=null;activeSku=null;
    var ref=tableRows(1,8,2026,C1,null,null);
    var top=Object.keys(ref.rows).sort(function(a,b){return ref.rows[b]-ref.rows[a];})[0];
    activeShipTo=top;
    var ch=shipToChannel(top);
    var kpi=0;for(var m=selFrom;m<=selTo;m++)kpi+=num(ch.monthly[String(m)]);
    var byC={},tot=0;
    (S.facts||[]).forEach(function(r){
      if(r[0]!==top||!shipToBasisSet()[r[1]])return;
      var m=Number(r[3]);if(m<selFrom||m>selTo)return;
      if(String(r[2])!==String(viewYear))return;
      byC[r[1]]=num(byC[r[1]]||0)+num(r[6]);tot+=num(r[6]);
    });
    __t("cards: channel-card split sums to KPI", Math.abs(tot-kpi)<0.001, "cards="+tot.toFixed(4)+" kpi="+kpi.toFixed(4));
    activeShipTo=prev.activeShipTo;selFrom=prev.selFrom;selTo=prev.selTo;viewYear=prev.viewYear;activeCat=prev.activeCat;activeSku=prev.activeSku;selChans=prev.selChans;
  })();

  // LY reconciliation: LY KPI must match the ship-to's LY month sum.
  (function(){
    var prev={activeShipTo:activeShipTo,selFrom:selFrom,selTo:selTo,viewYear:viewYear,activeCat:activeCat,activeSku:activeSku};
    selFrom=1;selTo=8;viewYear=2026;activeCat=null;activeSku=null;
    var ref2026=tableRows(1,8,2026,C1,null,null);
    var top=Object.keys(ref2026.rows).sort(function(a,b){return ref2026.rows[b]-ref2026.rows[a];})[0];
    activeShipTo=top;
    var lyAgg=shipToAggFor(top,2025,true);
    var lySum=0;for(var m=selFrom;m<=selTo;m++)lySum+=num(lyAgg.months[String(m)]);
    var ch=shipToChannel(top);
    var kpiLy=0;for(var m=selFrom;m<=selTo;m++)kpiLy+=num(ch.s_ly[String(m)]);
    __t("ly: LY KPI equals ship-to LY month sum", Math.abs(lySum-kpiLy)<0.001, "lyAgg="+lySum.toFixed(4)+" kpiLy="+kpiLy.toFixed(4));
    activeShipTo=prev.activeShipTo;selFrom=prev.selFrom;selTo=prev.selTo;viewYear=prev.viewYear;activeCat=prev.activeCat;activeSku=prev.activeSku;
  })();

  // toggleShipTo click / clear semantics.
  (function(){
    var start=activeShipTo;
    toggleShipTo("TESTOP"),__t("toggle: selects on first click", activeShipTo==="TESTOP");
    toggleShipTo("TESTOP"),__t("toggle: clears on second click", activeShipTo===null);
    toggleShipTo("TESTOP"),toggleShipTo("OTHER"); // switching ship-tos
    __t("toggle: switching selects new one", activeShipTo==="OTHER");
    toggleShipTo("OTHER"); // clear
    activeShipTo=start;
  })();

  // shipToLabel reflects active ship-to for section headers.
  (function(){
    var start=activeShipTo;
    activeShipTo="X"; var lb=shipToLabel();
    __t("label: header shows Ship-to", lb.indexOf("Ship-to")===0 && lb.indexOf("X")>=0, lb);
    activeShipTo=start;
  })();

  console.log("\\n"+(__test.failures?"  "+__test.failures+" FAILURES / "+__test.checks+" checks":"  ALL "+__test.checks+" CHECKS PASS"));
  return __test.failures;
})();
`;

const result = vm.runInContext(testScript, sandbox, { filename: "_test_shipto_run" });
process.exit(result && result > 0 ? 1 : 0);
