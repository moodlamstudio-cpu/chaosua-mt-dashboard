const MONTHS=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
let D=null,charts={},cur="MT",selChans=["MAKRO","LOTUS'"],selFrom=1,selTo=8,lastClosed=8,fy=2026;
function num(v){return (v==null||isNaN(v))?0:Number(v);}
function fmt(v){return (v==null||isNaN(v))?"-":Number(v).toLocaleString("en-US",{minimumFractionDigits:1,maximumFractionDigits:1});}
function pct(a,b){a=num(a);b=num(b);if(!b)return null;return ((a-b)/b*100);}
function destroy(id){if(charts[id]){charts[id].destroy();charts[id]=null;}}
function setT(id,h){var e=document.getElementById(id);if(e)e.innerHTML=h;}
function setV(id,t){var e=document.getElementById(id);if(e)e.textContent=t;}
var CATCOLR=["#93c5fd","#f9a8d4","#86efac","#fcd34d","#c4b5fd","#67e8f9","#fdba74","#a7f3d0","#fca5a5","#bfdbfe","#d8b4fe","#bef264"];
function dark(h){var n=parseInt(h.slice(1),16);return "rgb("+Math.max(0,((n>>16)&255)-35)+","+Math.max(0,((n>>8)&255)-35)+","+Math.max(0,(n&255)-35)+")";}
var activeCat=null,activeSku=null,kpiPick=null;
var valueLabelPlugin={id:"valueLabelPlugin",afterDatasetsDraw:function(chart){var ctx=chart.ctx;ctx.save();ctx.textAlign="center";ctx.textBaseline="bottom";ctx.fillStyle="#475569";ctx.font="700 12px sans-serif";chart.data.datasets.forEach(function(ds,di){var datasetType=ds.type||chart.config.type;if(datasetType!=="bar")return;var meta=chart.getDatasetMeta(di);if(!meta||!meta.data)return;meta.data.forEach(function(bar,i){var v=ds.data[i];if(v==null||isNaN(v)||v===0)return;var x=bar.x,y=bar.y;if(x==null||y==null)return;ctx.fillText(num(v).toFixed(0),x,y-6);});});ctx.restore();}};
if(typeof Chart!=="undefined"){try{Chart.register(valueLabelPlugin);}catch(e){}}
var piePlugin={id:"piePlugin",afterDraw:function(chart){var ctx=chart.ctx;var o=(chart.options.plugins||{}).pieText;if(!o||o.hide)return;var a=chart.chartArea,cx=(a.left+a.right)/2,cy=(a.top+a.bottom)/2;
 if(o.label!==undefined){ctx.save();ctx.textAlign="center";ctx.textBaseline="middle";ctx.fillStyle="#0f172a";ctx.font="700 22px Arial,sans-serif";ctx.fillText(String(o.label),cx,cy);ctx.restore();}}};
function sel(){return D[cur];}
function selAll(){
  var out={label:selChans.map(function(i){return D[i]&&D[i].label||i;}).join(" + "),monthly:{},history:{},category:{items:[]},annual:{},s_actual:{},s_le:{},s_ly:{},s_aop:{},monthly_by_cat:{},history_by_cat:{},annual_by_cat:{},monthly_by_sku:{},history_by_sku:{},annual_by_sku:{},_planningOK:true};
  selChans.forEach(function(id){var c=D[id];if(!c)return;
    if(!hasS(c))out._planningOK=false;
    for(var m=1;m<=12;m++){out.monthly[String(m)]=num(out.monthly[String(m)])+num(c.monthly&&c.monthly[String(m)]||0);
      if(c.s_actual)out.s_actual[String(m)]=num(out.s_actual[String(m)])+num(c.s_actual[String(m)]||0);
      if(c.s_le)out.s_le[String(m)]=num(out.s_le[String(m)])+num(c.s_le[String(m)]||0);
      if(c.s_ly)out.s_ly[String(m)]=num(out.s_ly[String(m)])+num(c.s_ly[String(m)]||0);
      if(c.s_aop)out.s_aop[String(m)]=num(out.s_aop[String(m)])+num(c.s_aop[String(m)]||0);}
    ["2025","2026"].forEach(function(y){out.history[y]=out.history[y]||{};var h=c.history&&c.history[y]||{};for(var m=1;m<=12;m++)out.history[y][String(m)]=num(out.history[y][String(m)])+num(h[String(m)]||0);});
    var an=c.annual||{};Object.keys(an).forEach(function(y){out.annual[y]=num(out.annual[y]||0)+num(an[y]);});
    var ci=c.category&&c.category.items||[];ci.forEach(function(it){var ex=null;for(var k=0;k<out.category.items.length;k++)if(out.category.items[k].name===it.name){ex=out.category.items[k];break;}if(ex)ex.mb=ex.mb+it.mb;else out.category.items.push({name:it.name,mb:it.mb,pct:it.pct});});
    // Merge category series for linked category filtering.
    var mbc=c.monthly_by_cat||{},hbc=c.history_by_cat||{},abc=c.annual_by_cat||{};mergeFocusMaps(out.monthly_by_cat,out.history_by_cat,out.annual_by_cat,mbc,hbc,abc);
    var mbs=c.monthly_by_sku||{},hbs=c.history_by_sku||{},abs=c.annual_by_sku||{};mergeFocusMaps(out.monthly_by_sku,out.history_by_sku,out.annual_by_sku,mbs,hbs,abs);
  });
  out.total_mb=0;for(var m=1;m<=12;m++)out.total_mb+=num(out.monthly[String(m)]||0);
  var tsa=[],tsbc={};selChans.forEach(function(id){var cs=D[id];if(!cs)return;if(cs.top_sku_all)tsa=tsa.concat(cs.top_sku_all);var t=cs.top_sku_by_cat||{};Object.keys(t).forEach(function(cat){tsbc[cat]=(tsbc[cat]||[]).concat(t[cat]);});});out.top_sku_all=mergeSkuList(tsa);Object.keys(tsbc).forEach(function(cat){tsbc[cat]=mergeSkuList(tsbc[cat]);});out.top_sku_by_cat=tsbc;
  return out;
}
function mergeFocusMaps(outM,outH,outA,srcM,srcH,srcA){Object.keys(srcM).forEach(function(key){outM[key]=outM[key]||{actual:{},ly:{}};for(var m=1;m<=12;m++){outM[key].actual[String(m)]=num(outM[key].actual[String(m)])+num((srcM[key].actual||{})[String(m)]);outM[key].ly[String(m)]=num(outM[key].ly[String(m)])+num((srcM[key].ly||{})[String(m)]);}});Object.keys(srcH).forEach(function(key){outH[key]=outH[key]||{};Object.keys(srcH[key]||{}).forEach(function(y){outH[key][y]=outH[key][y]||{};for(var m=1;m<=12;m++)outH[key][y][String(m)]=num(outH[key][y][String(m)])+num(srcH[key][y][String(m)]);});});Object.keys(srcA).forEach(function(key){outA[key]=outA[key]||{};Object.keys(srcA[key]||{}).forEach(function(y){outA[key][y]=num(outA[key][y])+num(srcA[key][y]);});});}
function mergeSkuList(list){var map={};(list||[]).forEach(function(s){var k=s.name;if(!map[k])map[k]={name:s.name,mb:0,cat:s.cat||""};map[k].mb+=num(s.mb);});return Object.keys(map).map(function(k){return map[k];}).sort(function(a,b){return b.mb-a.mb;});}
function label(){var ids=(D&&D._channels||[]).filter(function(ch){return ch.id!=="MT";}).map(function(ch){return ch.id;});if(ids.length&&selChans&&selChans.length===ids.length)return "All Channel";if(selChans&&selChans.length>1)return selChans.map(function(i){return D[i]&&D[i].label||i;}).join(" + ");return D[cur]?D[cur].label:cur;}
function C(){
  if(selChans&&selChans.length>1)return selAll();
  return sel();
}
function load(){
 fetch("data_channels.json?v="+Date.now()).then(r=>r.json()).then(function(data){
   D=data;lastClosed=8;fy=2026;
   // channels
   var sc=document.getElementById("selChannel");sc.innerHTML="";
   var allChannelIds=(D._channels||[]).filter(function(ch){return ch.id!=="MT";}).map(function(ch){return ch.id;});
   var allLb=document.createElement("label");allLb.className="chk";allLb.style.cssText="display:inline-flex;align-items:center;gap:5px;margin:3px 12px 3px 0;cursor:pointer;font-size:12px;font-weight:700;";
   var allCb=document.createElement("input");allCb.type="checkbox";allCb.value="__ALL__";allLb.appendChild(allCb);allLb.appendChild(document.createTextNode("All Channel"));sc.appendChild(allLb);
   function scanChk(){var cbs=sc.querySelectorAll("input[type=checkbox]");for(var i=0;i<cbs.length;i++)cbs[i].checked=(cbs[i].value==="__ALL__")?(selChans.length===allChannelIds.length):(selChans.indexOf(cbs[i].value)>=0);}
   allCb.onchange=function(){selChans=allCb.checked?allChannelIds.slice():["MAKRO","LOTUS'"];cur=selChans[0];activeCat=null;activeSku=null;kpiPick=null;scanChk();renderAll();};
   (D._channels||[]).forEach(function(ch){
     if(ch.id==="MT")return;
     var lb=document.createElement("label");lb.className="chk";lb.style.cssText="display:inline-flex;align-items:center;gap:5px;margin:3px 12px 3px 0;cursor:pointer;font-size:12px;font-weight:600;";
     var cb=document.createElement("input");cb.type="checkbox";cb.value=ch.id;
     var nm=(ch.id==="LOTUS'"?"Lotus's":(D[ch.id]&&D[ch.id].label||ch.label));
     lb.appendChild(cb);lb.appendChild(document.createTextNode(nm));sc.appendChild(lb);
     cb.onchange=function(){var x=selChans.indexOf(ch.id);if(cb.checked&&x<0)selChans.push(ch.id);else if(!cb.checked&&x>=0)selChans.splice(x,1);if(!selChans.length)selChans=[ch.id];scanChk();cur=selChans[0];activeCat=null;activeSku=null;kpiPick=null;renderAll();};
   });
   scanChk();cur=selChans[0];
   var sy=document.getElementById("selYear");sy.innerHTML='<option value="2026">2026</option>';sy.value="2026";sy.disabled=true;
   // from/to
   var sf=document.getElementById("selFrom"),st=document.getElementById("selTo");
   sf.innerHTML="";st.innerHTML="";
   for(var i=1;i<=12;i++){var o=document.createElement("option");o.value=i;o.textContent=MONTHS[i-1];sf.appendChild(o);}
   for(var i=1;i<=12;i++){var o=document.createElement("option");o.value=i;o.textContent=MONTHS[i-1];st.appendChild(o);}
   sf.value=1;st.value=lastClosed;selFrom=1;selTo=lastClosed;
   sf.onchange=function(){selFrom=+sf.value;if(selTo<selFrom)selTo=selFrom;st.value=selTo;kpiPick=null;renderAll();};
   st.onchange=function(){selTo=+st.value;if(selFrom>selTo)selFrom=selTo;sf.value=selFrom;kpiPick=null;renderAll();};
   renderAll();
 }).catch(function(e){setT("mTable","Load error: "+e.message);});
}
function rangeSum(map,f,t){if(!map)return 0;var s=0;for(var m=f;m<=t;m++)s+=num(map[String(m)]);return s;}
function hasSeriesData(o){if(!o)return false;for(var k in o){if(num(o[k])!==0)return true;}return false;}
function actSeries(c){if(c&&hasS(c))return c.s_actual;return c&&c.monthly;}
function hasS(ch){ // genuine planning series required; empty aggregate objects must NOT count
  return ch && ch._planningOK!==false && hasSeriesData(ch.s_actual) && hasSeriesData(ch.s_le) && hasSeriesData(ch.s_ly) && hasSeriesData(ch.s_aop);
}
function comparisonLySeries(c){if(c&&c._planningOK!==false&&hasSeriesData(c.s_ly))return c.s_ly;return c&&c.history&&c.history[String(fy-1)]||{};}
function renderAll(){
  try{
    var c=C();var __foc__=currentFocus(c);var __fn=__foc__&&__foc__.name?(" "+String.fromCharCode(8226)+" "+__foc__.name):"";setV("tgY",label()+__fn);setV("tgM",label()+__fn);setV("tgTbl",label()+__fn);setV("tgCat",activeCat?("Category "+String.fromCharCode(8226)+" "+activeCat):label());setV("tgSku",label()+__fn);
    var ok=hasS(c);
    var rangeTxt=MONTHS[selFrom-1]+"-"+MONTHS[selTo-1];
    setV("lblAct",rangeTxt);setV("lblLE",rangeTxt);setV("lblLY",rangeTxt);setV("lblAOP",rangeTxt);
    if(!ok){
      // Channels without LE/AOP still show LY when prior-year history exists.
      var focNoPlan=currentFocus(c),a=rangeSum(focNoPlan?focNoPlan.series.actual:actSeries(c),selFrom,selTo),lyNoPlanSeries=focNoPlan?focNoPlan.series.ly:comparisonLySeries(c),lyNoPlan=rangeSum(lyNoPlanSeries,selFrom,selTo),hasLyNoPlan=hasSeriesData(lyNoPlanSeries);
      setT("kACT",fmt(a)+' <small>MB</small>');setT("kACTuntil",MONTHS[lastClosed-1]+" "+fy);setT("kACTdate","12 Aug 2026");
      setT("kLE","\u2014");setT("kGap","\u2014");
      setT("kLY",hasLyNoPlan?fmt(lyNoPlan)+' <small>MB</small>':"\u2014");setT("kLYgap",hasLyNoPlan?((a-lyNoPlan>=0?"+":"-")+Math.abs(a-lyNoPlan).toFixed(1)+" MB ("+(a-lyNoPlan>=0?"+":"-")+Math.abs(lyNoPlan?(a-lyNoPlan)/lyNoPlan*100:0).toFixed(1)+"%)"):"");setT("kLYcmp",hasLyNoPlan?"Prior-year history":"No LY data");
      setT("kAOP","\u2014");setT("kAOPgap","");setT("kAOPpc","");setT("kAOPcmp","No AOP data");
    } else {
      var focKPI=currentFocus(c),focusNameKPI=focKPI&&focKPI.name;
      var a=0,leK=0,lyK=0,aopK=0;
      if(focusNameKPI){ // Focus mode: use the selected category/SKU Actual and LY only.
        var sr=focKPI.series;for(var mm=selFrom;mm<=selTo;mm++){a+=num(sr.actual&&sr.actual[String(mm)]||0);lyK+=num(sr.ly&&sr.ly[String(mm)]||0);}
      } else { a=rangeSum(c.s_actual,selFrom,selTo);leK=0;for(var mm2=selFrom;mm2<=selTo;mm2++){if(mm2<lastClosed)leK+=num(c.s_actual[String(mm2)]);else leK+=num(c.s_le[String(mm2)]);}lyK=rangeSum(c.s_ly,selFrom,selTo);aopK=rangeSum(c.s_aop,selFrom,selTo);}
      setT("kACT",fmt(a)+' <small>MB</small>');setT("kACTuntil",MONTHS[lastClosed-1]+" "+fy);setT("kACTdate","12 Aug 2026");
      if(focusNameKPI){
        setT("kLE","\u2014");setT("kGap","\u2014");
        setT("kLY",lyK>0?fmt(lyK)+' <small>MB</small>':"\u2014");setT("kLYgap","");
        setT("kAOP","\u2014");setT("kAOPgap","");
      } else {
        var le=leK,ly=lyK,aop=aopK;
        var leDiff=a-le,leP=le>0?(leDiff/le*100):0,leCls=leDiff>=0?"up":"down";
        setT("kLE",fmt(le)+' <small>MB</small>');setT("kGap",(leDiff>=0?"+":"-")+Math.abs(leDiff).toFixed(1)+" MB ("+(leDiff>=0?"+":"-")+Math.abs(leP).toFixed(1)+"%)");document.getElementById("kGap").className=leCls;
        var lyDiff=a-ly,lyP=ly>0?(lyDiff/ly*100):0,lyCls=lyDiff>=0?"up":"down";
        setT("kLY",fmt(ly)+' <small>MB</small>');setT("kLYgap",(lyDiff>=0?"+":"-")+Math.abs(lyDiff).toFixed(1)+" MB ("+(lyDiff>=0?"+":"-")+Math.abs(lyP).toFixed(1)+"%)");document.getElementById("kLYgap").className=lyCls;setT("kLYpc","");
        var aopDiff=le-aop,aopP=aop>0?(aopDiff/aop*100):0,aopCls=aopDiff>=0?"up":"down";
        setT("kAOP",fmt(aop)+' <small>MB</small>');setT("kAOPgap",(aopDiff>=0?"+":"-")+Math.abs(aopDiff).toFixed(1)+" MB ("+(aopDiff>=0?"+":"-")+Math.abs(aopP).toFixed(1)+"%)");document.getElementById("kAOPgap").className=aopCls;setT("kAOPpc","");
      }
    }
    setV("hDate",MONTHS[lastClosed-1]+" "+fy);
    drawAnnual();drawMonthly();drawTable();
    drawCategory(c);drawSku(c);drawCategoryPerformance(c);drawChannelSummary();
    if(kpiPick)applyKpiPick();
  }catch(e){setT("tgM","ERR: "+e.message);}
}
function chgBg(m,sel){return m===sel?"#e11d48":"#f87171";}
function sv(c,key,m){ if(c&&c[key]&&hasSeriesData(c[key])) return num(c[key][String(m)]); if(c&&c.monthly) return num(c.monthly[String(m)]); return 0; }
function currentFocus(c){if(activeSku&&c.monthly_by_sku&&c.monthly_by_sku[activeSku])return {name:activeSku,series:c.monthly_by_sku[activeSku],annual:c.annual_by_sku&&c.annual_by_sku[activeSku],history:c.history_by_sku&&c.history_by_sku[activeSku]};if(activeCat&&c.monthly_by_cat&&c.monthly_by_cat[activeCat])return {name:activeCat,series:c.monthly_by_cat[activeCat],annual:c.annual_by_cat&&c.annual_by_cat[activeCat],history:c.history_by_cat&&c.history_by_cat[activeCat]};return null;}
function pickKpi(type,value){if(type==="year"){if(kpiPick&&kpiPick.type==="month")kpiPick={type:"yearMonth",year:String(value),month:+kpiPick.value};else if(kpiPick&&kpiPick.type==="yearMonth")kpiPick=(String(kpiPick.year)===String(value))?{type:"month",value:kpiPick.month}:{type:"yearMonth",year:String(value),month:kpiPick.month};else{kpiPick=(kpiPick&&kpiPick.type==="year"&&String(kpiPick.value)===String(value))?null:{type:"year",value:String(value)};}}else{if(kpiPick&&kpiPick.type==="year")kpiPick={type:"yearMonth",year:String(kpiPick.value),month:+value};else if(kpiPick&&kpiPick.type==="yearMonth")kpiPick=(+kpiPick.month===+value)?{type:"year",value:String(kpiPick.year)}:{type:"yearMonth",year:String(kpiPick.year),month:+value};else kpiPick=(kpiPick&&kpiPick.type==="month"&&+kpiPick.value===+value)?null:{type:"month",value:+value};}setTimeout(renderAll,0);}
function applyKpiPick(){var c=C(),focus=currentFocus(c);if(kpiPick.type==="year"){var y=String(kpiPick.value),annual=(focus&&focus.annual)||c.annual||{},prev=String(+y-1);setV("lblAct",y);setT("kACT",fmt(annual[y])+' <small>MB</small>');setT("kACTuntil","Full year / YTD");setV("lblLE",y);setT("kLE","\u2014");setT("kGap","\u2014");setV("lblLY",prev);setT("kLY",annual[prev]!=null?fmt(annual[prev])+' <small>MB</small>':"\u2014");setT("kLYcmp","Previous year");setV("lblAOP",y);setT("kAOP","\u2014");setT("kAOPcmp","No annual AOP");return;}
 var m=kpiPick.type==="yearMonth"?+kpiPick.month:+kpiPick.value,y=kpiPick.type==="yearMonth"?+kpiPick.year:fy,history=(focus&&focus.history)||c.history||{},actual=num(history[String(y)]&&history[String(y)][String(m)]),ly=history[String(y-1)]?num(history[String(y-1)][String(m)]):null,current=(y===fy),le=(!focus&&current&&hasS(c))?(m<lastClosed?actual:sv(c,"s_le",m)):null,aop=(!focus&&current&&hasS(c))?sv(c,"s_aop",m):null;["lblAct","lblLE","lblLY","lblAOP"].forEach(function(id){setV(id,MONTHS[m-1]+" "+y);});setT("kACT",fmt(actual)+' <small>MB</small>');setT("kACTuntil",MONTHS[m-1]+" "+y);setT("kLE",le==null?"\u2014":fmt(le)+' <small>MB</small>');setT("kGap",le==null?"\u2014":fmt(Math.max(0,le-actual))+" MB");setT("kLY",ly==null?"\u2014":fmt(ly)+' <small>MB</small>');setT("kLYcmp",ly==null?"No LY data":"Monthly LY "+(y-1));if(aop){var aopDiff=le-aop,aopP=aopDiff/aop*100;setT("kAOP",fmt(aop)+' <small>MB</small>');setT("kAOPgap",(aopDiff>=0?"+":"-")+Math.abs(aopDiff).toFixed(1)+" MB ("+(aopDiff>=0?"+":"-")+Math.abs(aopP).toFixed(1)+"%)");document.getElementById("kAOPgap").className=aopDiff>=0?"up":"down";setT("kAOPpc","");setT("kAOPcmp","Monthly LE vs AOP");}else{setT("kAOP","\u2014");setT("kAOPgap","");setT("kAOPpc","");setT("kAOPcmp","No AOP data");}}

function drawAnnual(){destroy("cAnnual");var c=C(),focus=currentFocus(c),annual=(focus&&focus.annual)||c.annual||{},history=(focus&&focus.history)||c.history||{},years=Object.keys(c.annual||annual).sort(),month=kpiPick&&(kpiPick.type==="month"?+kpiPick.value:(kpiPick.type==="yearMonth"?+kpiPick.month:null)),vals=years.map(function(y){return month?num(history[y]&&history[y][String(month)]):num(annual[y]);}),selectedYear=kpiPick&&(kpiPick.type==="year"?String(kpiPick.value):(kpiPick.type==="yearMonth"?String(kpiPick.year):null)),colors=years.map(function(y){return selectedYear===y?"#000000":"#525252";}),base=focus?(label()+" \u2022 "+focus.name):label();setV("tgY",base+(month?" \u2022 "+MONTHS[month-1]:""));charts["cAnnual"]=new Chart(document.getElementById("cAnnual"),{type:"bar",data:{labels:years,datasets:[{label:"Actual",data:vals,backgroundColor:colors,hoverBackgroundColor:"#000000",borderRadius:6,barPercentage:.55}]},
  plugins:[valueLabelPlugin],options:{responsive:true,maintainAspectRatio:false,onClick:function(e,els){if(els&&els.length)pickKpi("year",years[els[0].index]);},plugins:{legend:{display:false},tooltip:{callbacks:{label:function(x){return (month?MONTHS[month-1]:"Actual")+": "+fmt(x.parsed.y)+" MB";}}}},scales:{y:{beginAtZero:true,grace:"15%"},x:{grid:{display:false}}}}});}

function drawMonthly(){destroy("cMonthly");var c=C();var lb=[],act=[],le=[],ly=[],aop=[];
 var focusInfo=currentFocus(c),focus=focusInfo&&focusInfo.series,focusName=focusInfo&&focusInfo.name||"",viewYear=kpiPick&&(kpiPick.type==="year"?+kpiPick.value:(kpiPick.type==="yearMonth"?+kpiPick.year:fy))||fy,history=(focusInfo&&focusInfo.history)||c.history||{},historical=(viewYear!==fy),seriesOk=hasS(c)&&!focus&&!historical;
 setV("tgM",(focusName?(label()+" \u2022 "+focusName):label())+(historical?" \u2022 "+viewYear:""));
 ["lgLE","lgAOP"].forEach(function(id){var e=document.getElementById(id);if(e)e.style.display=(focus||historical)?"none":"inline-flex";});
 var actualSeries=historical?(history[String(viewYear)]||{}):(focus?focus.actual:null),lySeries=historical?(history[String(viewYear-1)]||{}):(focus?focus.ly:comparisonLySeries(c)),lyAvailable=hasSeriesData(lySeries),maxMonth=historical?12:lastClosed,latestVisible=Math.min(selTo,maxMonth);
 for(var i=1;i<=12;i++){var m=i;lb.push((m>=selFrom&&m<=selTo)?MONTHS[m-1]:"");act.push((m<=maxMonth&&m>=selFrom&&m<=selTo)?((focus||historical)?num(actualSeries[String(m)]):num(actSeries(c)[String(m)])):null);le.push((seriesOk&&m>=lastClosed&&m>=selFrom&&m<=selTo)?sv(c,"s_le",m):null);ly.push((lyAvailable&&m>=selFrom&&m<=selTo)?num(lySeries[String(m)]):null);aop.push((seriesOk&&m>=selFrom&&m<=selTo)?sv(c,"s_aop",m):null);}
 charts["cMonthly"]=new Chart(document.getElementById("cMonthly"),{data:{labels:lb,datasets:[
  {type:"bar",label:"Actual",data:act,hoverBackgroundColor:"#000000",backgroundColor:act.map(function(v,i){var picked=kpiPick&&(kpiPick.type==="month"?+kpiPick.value:(kpiPick.type==="yearMonth"?+kpiPick.month:null));return picked===i+1?"#000000":"#525252";}),borderRadius:5,barPercentage:.65},
  {type:"bar",label:"LE",data:le,backgroundColor:"#cbd5e1",borderRadius:5,barPercentage:.65},
  {type:"line",label:"LY",data:ly,borderColor:"#64748b",backgroundColor:"#64748b",borderDash:[5,3],fill:false,tension:.3,pointRadius:2,borderWidth:2},
  {type:"line",label:"AOP",data:aop,borderColor:"#2563eb",backgroundColor:"#2563eb",fill:false,tension:.3,pointRadius:2,borderWidth:2}]},plugins:[valueLabelPlugin],options:{responsive:true,maintainAspectRatio:false,interaction:{mode:"index",intersect:false},onClick:function(e,els){if(els&&els.length&&els[0].datasetIndex<=1)pickKpi("month",els[0].index+1);},
   plugins:{legend:{display:false},tooltip:{callbacks:{label:function(x){return x.dataset.label+": "+fmt(x.parsed.y)+" MB";}}}},
   scales:{y:{beginAtZero:true,ticks:{callback:function(v){return fmt(v);}}},x:{grid:{display:false}}}}});
}
function signCls(v){return Math.abs(v)<0.05?'up':(v<0?'down':'up');}
function signTxt(v,p,pmax){var vs=fmt(v);if(Math.abs(v)<0.05)vs='0.0';else if(v>0)vs='+'+fmt(v);if(p==null)return {v:vs,p:''};var pt=p.toFixed(1)+'%';if(Math.abs(p)<0.05)pt='0.0%';else if(p>0)pt='+'+pt;return {v:vs,p:pt};}
function cmpDoubleCell(title,vs,p){if(vs==null)return '<td>\u2014</td><td>\u2014</td>';var s=signCls(vs),t=signTxt(vs,p);return '<td class="'+s+'">'+t.v+'</td><td class="'+s+'">'+t.p+'</td>';}
function drawTable(){var c=C(),focusInfo=currentFocus(c),focus=focusInfo&&focusInfo.series,isTotal=!focus;
 var h='<thead><tr><th>Month</th><th>Actual</th><th>LE</th><th>LY 2025</th><th>AOP</th><th>VS LY</th><th>%VS LY</th><th>VS AOP</th><th>%VS AOP</th></tr></thead><tbody>';
 var seriesOk=hasS(c),lySeries=focus?focus.ly:comparisonLySeries(c),lyAvailable=hasSeriesData(lySeries),ta=0,tl=0,ty=0,ta2=0,tvly=0,tvaop=0;
 for(var m=1;m<=12;m++){if(m<selFrom||m>selTo)continue;
  var act,le,ly,aop;
  if(focus){act=(m<=lastClosed?num(focus.actual[String(m)]):null);le=null;ly=lyAvailable?num(lySeries[String(m)]):null;aop=null;}
  else{act=(m<=lastClosed)?num(actSeries(c)[String(m)]):null;le=seriesOk?((m<lastClosed)?act:sv(c,"s_le",m)):null;ly=lyAvailable?num(lySeries[String(m)]):null;aop=seriesOk?sv(c,"s_aop",m):null;}
  var vly=(act!=null&&ly!=null)?(act-ly):null,ply=(vly!=null&&ly>0)?(vly/ly*100):null,vaop=(act!=null&&aop!=null)?(act-aop):null,paop=(vaop!=null&&aop>0)?(vaop/aop*100):null;
  ta+=act||0;ty+=ly||0;tvly+=vly||0;if(isTotal){tl+=le||0;ta2+=aop||0;tvaop+=vaop||0;}
 h+="<tr><td>"+MONTHS[m-1]+"</td><td>"+(act==null?"\u2014":fmt(act))+"</td><td>"+(le==null?"\u2014":fmt(le))+"</td><td>"+(ly==null?"\u2014":fmt(ly))+"</td><td>"+(aop==null?"\u2014":fmt(aop))+"</td>"+cmpDoubleCell("VS LY",vly,ply)+cmpDoubleCell("VS AOP",vaop,paop)+"</tr>";}
 var ptvly=(ty>0)?(tvly/ty*100):null,ptvaop=(isTotal&&ta2>0)?(tvaop/ta2*100):null;
 h+='<tr class="total"><td>Total</td><td>'+fmt(ta)+'</td><td>'+(isTotal&&seriesOk?fmt(tl):'\u2014')+'</td><td>'+(lyAvailable?fmt(ty):'\u2014')+'</td><td>'+(isTotal&&seriesOk?fmt(ta2):'\u2014')+'</td>'+cmpDoubleCell("Total VS LY",lyAvailable?tvly:null,lyAvailable?ptvly:null)+(isTotal&&seriesOk?cmpDoubleCell("Total VS AOP",tvaop,ptvaop):'<td>\u2014</td><td>\u2014</td>')+'</tr></tbody>';
 setT("mTable",h);}
function drawCategory(c){var el=document.getElementById("cCat");if(!el||!Chart)return;destroy("cCat");var cT=document.getElementById("tgCat");if(cT)cT.textContent=(activeCat?("Category \u2022 "+activeCat):label());
 var allCats=(c.category&&c.category.items)?c.category.items:[];if(!allCats.length){setT("catLegend","No data");return;}
 var selectedIndex=-1;for(var ci=0;ci<allCats.length;ci++)if(allCats[ci].name===activeCat){selectedIndex=ci;break;}
 var C2=selectedIndex>=0?[{name:allCats[selectedIndex].name,mb:allCats[selectedIndex].mb,pct:100}]:allCats;
 var tot=c.total_mb,focus=currentFocus(c);if(focus&&focus.series)tot=rangeSum(focus.series.actual,1,lastClosed);if(selectedIndex>=0)C2[0].mb=tot;
 var vals=C2.map(function(x){return x.mb;});var colors=C2.map(function(x,i){var colorIndex=selectedIndex>=0?selectedIndex:i;return CATCOLR[colorIndex%CATCOLR.length];});
 var _sd=C2.map(function(x){return {pct:selectedIndex>=0?100:(x.pct||0)};});
 charts["cCat"]=new Chart(el,{type:"doughnut",data:{labels:C2.map(function(x){return x.name;}),datasets:[{data:vals,backgroundColor:colors,borderWidth:selectedIndex>=0?0:1,borderColor:"#fff",_sd:_sd}]},plugins:[piePlugin],
  options:{responsive:true,maintainAspectRatio:false,cutout:"60%",onClick:function(ev,els){if(els&&els.length){var idx=els[0].index;if(C2[idx])toggleCat(C2[idx].name);}},
   plugins:{title:{display:false},pieText:{label:fmt(tot)+' MB',sub:""},legend:{display:false},tooltip:{callbacks:{label:function(x){var it=C2[x.dataIndex];return it.name+": "+fmt(it.mb)+" MB ("+it.pct+"%)";}}}}}});
 var lg="";C2.forEach(function(x,i){var colorIndex=selectedIndex>=0?selectedIndex:i;var bg=(activeCat===x.name)?"background:#fef9c3;":"";var shownPct=selectedIndex>=0?"100.0":x.pct;lg+='<div class="catrow" data-cat="'+x.name+'" style="'+bg+'display:flex;justify-content:space-between;align-items:center;padding:5px 0;border-bottom:1px solid #f1f5f9;cursor:pointer;border-left:4px solid '+CATCOLR[colorIndex%CATCOLR.length]+';padding-left:8px"><span>'+x.name+'</span><b>'+fmt(x.mb)+' MB &middot; '+shownPct+'%</b></div>';});
 setT("catLegend",lg);
 var rows=document.querySelectorAll(".catrow");for(var i=0;i<rows.length;i++){(function(r){r.onclick=function(){toggleCat(r.dataset.cat);};})(rows[i]);}}
function toggleCat(cat){activeCat=(activeCat===cat)?null:cat;activeSku=null;kpiPick=null;renderAll();}
function toggleSku(name,cat){if(activeSku===name){activeSku=null;activeCat=null;}else{activeSku=name;activeCat=cat||null;}kpiPick=null;renderAll();}
function perfValues(series){var actual=series&&series.actual||{},ly=series&&series.ly||{},hasLy=hasSeriesData(ly),a=rangeSum(actual,selFrom,selTo),l=hasLy?rangeSum(ly,selFrom,selTo):null,g=hasLy?a-l:null,p=(hasLy&&l!==0)?g/l*100:null;return {actual:a,ly:l,growth:g,pct:p};}
function perfCell(v,suffix){return v==null?'\u2014':fmt(v)+(suffix||'');}
function signedPerf(v,suffix){if(v==null)return '\u2014';return '<span class="'+signCls(v)+'">'+(v>0?'+':'')+fmt(v)+(suffix||'')+'</span>';}
function drawCategoryPerformance(c){
 var map=c.monthly_by_cat||{},names=Object.keys(map),total=rangeSum(actSeries(c),selFrom,selTo);
 names.sort(function(a,b){return perfValues(map[b]).actual-perfValues(map[a]).actual;});
 var h='<thead><tr><th style="text-align:left">Category</th><th>Actual 2025</th><th>Actual 2026</th><th>Growth vs LY (MB)</th><th>% Growth vs LY</th><th>% Contribution</th></tr></thead><tbody>';
 if(!names.length)h+='<tr><td>No Category</td><td colspan="5">\u2014</td></tr>';
 names.forEach(function(name){var v=perfValues(map[name]),con=total?v.actual/total*100:0,cls='perf-row'+(v.growth!=null&&v.growth>0?' growing':'')+(activeCat===name?' selected':'');h+='<tr class="'+cls+'" data-cat="'+encodeURIComponent(name)+'"><td style="text-align:left;white-space:normal">'+name+'</td><td>'+perfCell(v.ly)+'</td><td>'+fmt(v.actual)+'</td><td>'+signedPerf(v.growth)+'</td><td>'+signedPerf(v.pct,'%')+'</td><td>'+fmt(con)+'%</td></tr>';});
 h+='</tbody>';setT('catPerfTable',h);setV('tgCatPerf',label()+(activeCat?' \u2022 '+activeCat:''));
 Array.prototype.forEach.call(document.querySelectorAll('#catPerfTable .perf-row'),function(r){r.onclick=function(){toggleCat(decodeURIComponent(r.getAttribute('data-cat')));};});
}
function drawSku(c){var total=rangeSum(actSeries(c),selFrom,selTo);var list;
 if(activeSku){var base=(c.top_sku_by_cat&&c.top_sku_by_cat[activeCat])?c.top_sku_by_cat[activeCat]:[];list=base.filter(function(s){return s.name===activeSku;});}
 else if(activeCat){list=(c.top_sku_by_cat&&c.top_sku_by_cat[activeCat])?c.top_sku_by_cat[activeCat]:[];}
 else{list=c.top_sku_all||[];}
 if(!list.length){setT("skuTable",'<thead><tr><th>SKU</th><th>Category</th><th>Actual 2025</th><th>Actual 2026</th><th>Growth vs LY (MB)</th><th>% Growth vs LY</th><th>% Contribution</th></tr></thead><tbody><tr><td>No SKU</td><td colspan="6">\u2014</td></tr></tbody>');return;}
 var h='<thead><tr><th>#</th><th style="text-align:left">SKU</th><th>Category</th><th>Actual 2025</th><th>Actual 2026</th><th>Growth vs LY (MB)</th><th>% Growth vs LY</th><th>% Contribution</th></tr></thead><tbody>';
 var shown=list.slice();if(activeSku){var sf=c.monthly_by_sku&&c.monthly_by_sku[activeSku];if(sf)total=rangeSum(sf.actual,selFrom,selTo);}else if(activeCat){var cf=c.monthly_by_cat&&c.monthly_by_cat[activeCat];if(cf)total=rangeSum(cf.actual,selFrom,selTo);}shown.sort(function(a,b){return perfValues(c.monthly_by_sku&&c.monthly_by_sku[b.name]).actual-perfValues(c.monthly_by_sku&&c.monthly_by_sku[a.name]).actual;});shown.forEach(function(s,i){var cat=s.cat||activeCat||"",v=perfValues(c.monthly_by_sku&&c.monthly_by_sku[s.name]),p=total?v.actual/total*100:0,cls='skuro perf-row'+(v.growth!=null&&v.growth>0?' growing':'')+(activeSku===s.name?' selected':'');h+='<tr class="'+cls+'" data-i="'+i+'"><td>'+(i+1)+'</td><td style="max-width:260px;white-space:normal;text-align:left">'+s.name+'</td><td>'+cat+'</td><td>'+perfCell(v.ly)+'</td><td>'+fmt(v.actual)+'</td><td>'+signedPerf(v.growth)+'</td><td>'+signedPerf(v.pct,'%')+'</td><td>'+fmt(p)+'%</td></tr>';});
 h+='</tbody>';setT("skuTable",h);
 var rows=document.querySelectorAll(".skuro");for(var i=0;i<rows.length;i++){(function(r){r.onclick=function(){var s=shown[+r.dataset.i];if(s)toggleSku(s.name,s.cat||activeCat||"");};})(rows[i]);}}
function chGrowthYTD(d,x){
  // Compare current YTD with the same closed-month range in the prior year.
  var ly=0,h=(d&&d.history&&d.history[x]);if(h){for(var m=1;m<=lastClosed;m++)ly+=num(h[String(m)]);}
  var a=num(d&&d.total_mb);
  var g=(ly>0)?((a-ly)/ly*100):null;
  return {a:a,ly:ly,g:g};
}
var CHLOGO={'7Eleven':'7-Eleven','MAKRO':'MAKRO',"LOTUS'":"LOTUS'",'BigC':'Big C','Tops':'Tops','CJ Express':'CJ','Jiffy':'Jiffy','Foodland':'Foodland','FamilyMart':'FamilyMart','Golden Place':'GP','Central food wholesales':'CF','Maxmart':'Maxmart','MaxValue':'MaxValue','Villa market':'Villa','Lawson':'Lawson'};
function selectChan(id){cur=id;selChans=[id];activeCat=null;activeSku=null;kpiPick=null;if(document.getElementById("selChannel")){var c=document.getElementById("selChannel").querySelectorAll("input[type=checkbox]");for(var i=0;i<c.length;i++)c[i].checked=(c[i].value===id);}renderAll();}

function catMaps(cat){ // {byChan:{chid:{monthly,history,annual}}, tot, byMonth{...}}
  var byChan={}, totMap={};
  (selChans.length?selChans:["MAKRO"]).forEach(function(chid){
    var cc=D[chid];if(!cc)return;
    var mb=cc.monthly_by_cat&&cc.monthly_by_cat[cat], hb=cc.history_by_cat&&cc.history_by_cat[cat], ab=cc.annual_by_cat&&cc.annual_by_cat[cat];
    if(!mb&&!hb&&!ab)return;
    var mo=mb&&mb.actual||{}, ly=mb&&mb.ly||{};
    byChan[chid]={monthly:mo,ly:ly,history:hb||{},annual:ab||{}};
    for(var m=1;m<=12;m++){totMap[String(m)]=num(totMap[String(m)]||0)+num(mo[String(m)]||0);}
  });
  return {byChan:byChan,totMap:totMap};
}
function chanGrowth(d){
  var f=num(selFrom),t=num(selTo);
  var str=String(fy-1);
  var ly=0,h=(d&&d.history&&d.history[str]);if(h){for(var m=f;m<=t;m++)ly+=num(h[String(m)]);}
  var a=0,hc=(d&&d.history&&d.history[String(fy)]);if(hc){for(var m=f;m<=t;m++)a+=num(hc[String(m)]);}
  if(!a)a=num(d&&d.total_mb);
  var g=(ly>0)?((a-ly)/ly*100):null;
  return {a:a,ly:ly,g:g};
}
function drawChannelSummary(){
  var el=document.getElementById("chTable");if(!el)return;
  var rows=[],tot=0;
  var focInfo=currentFocus(C()),focusName=focInfo&&focInfo.name;
  // Cards follow the active channel filter so their Actual total reconciles to the KPI.
  var channelDefs={};(D._channels||[]).forEach(function(ch){channelDefs[ch.id]=ch;});
  (selChans||[]).forEach(function(channelId){
    var ch=channelDefs[channelId]||{id:channelId,label:channelId};
    if(ch.id==="MT")return;
    var d=D[ch.id];if(!d)return;
    var g;
    if(focusName){
      var cm=activeSku?(D[ch.id].monthly_by_sku&&D[ch.id].monthly_by_sku[activeSku]):(D[ch.id].monthly_by_cat&&D[ch.id].monthly_by_cat[activeCat]);
      var a=0,ly=0;if(cm){for(var m=num(selFrom);m<=num(selTo);m++){a+=num(cm.actual&&cm.actual[String(m)]||0);ly+=num(cm.ly&&cm.ly[String(m)]||0);}}
      g={a:a,ly:ly,g:ly>0?(a-ly)/ly*100:(a>0?null:null)};
    } else { g=chanGrowth(d); }
    rows.push({id:ch.id,label:(ch.id==="LOTUS'"?"Lotus's":(d.label||ch.label)),a:g.a,ly:g.ly,g:g.g});
    tot+=g.a;
  });
  var actCur=cur,actSel=(selChans&&selChans.length===1)?selChans[0]:null;rows.sort(function(x,y){return y.a-x.a;});
  var h='<div class="chgrid">';
  rows.forEach(function(r){
    var sh=tot?(r.a/tot*100):0;
    var grow=r.g;
    var cls=(grow!=null&&grow>=0)?"up":"down";
    var gap=r.a-r.ly;
    var gtxt=(grow==null)?"\u2014":((gap>=0?"+":"")+gap.toFixed(1)+" MB ("+(grow>=0?"+":"")+grow.toFixed(1)+"%)");
    var lg="logos/"+r.id+".png";
        var bg=grow!=null?(grow>=0?"#dcfce7":"#fff"):"#fff";
    var bstyle=bg!=="#fff"?" style=\"background:"+bg+"\"":"";
    h+='<div class="chcard" data-ch="'+r.id+'"'+bstyle+' style="cursor:pointer">';
    h+='  <div class="chhead"><span class="chname">'+r.label+'</span></div>';
    h+='  <div class="chcon">'+sh.toFixed(1)+'%<small>Con</small></div>';
    h+='  <div class="chval">'+fmt(r.a)+'<small>MB</small></div>';
    h+='  <div class="chly">LY '+fmt(r.ly)+' Gap <span class="'+cls+'">'+gtxt+'</span></div>';
    h+='</div>';
  });
  h+='</div>';
  el.innerHTML=h;
  Array.prototype.forEach.call(document.querySelectorAll(".chcard[data-ch]"),function(el2){el2.addEventListener("click",function(){selectChan(el2.getAttribute("data-ch"));});});
}
load();
