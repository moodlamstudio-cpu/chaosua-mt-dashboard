const MONTHS=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
let D=null,charts={},cur="MT",selFrom=1,selTo=8,lastClosed=8,fy=2026;
function num(v){return (v==null||isNaN(v))?0:Number(v);}
function fmt(v){return (v==null||isNaN(v))?"-":Number(v).toLocaleString("en-US",{minimumFractionDigits:1,maximumFractionDigits:1});}
function pct(a,b){a=num(a);b=num(b);if(!b)return null;return ((a-b)/b*100);}
function destroy(id){if(charts[id]){charts[id].destroy();charts[id]=null;}}
function setT(id,h){var e=document.getElementById(id);if(e)e.innerHTML=h;}
function setV(id,t){var e=document.getElementById(id);if(e)e.textContent=t;}
var CATCOLR=["#93c5fd","#f9a8d4","#86efac","#fcd34d","#c4b5fd","#67e8f9","#fdba74","#a7f3d0","#fca5a5","#bfdbfe","#d8b4fe","#bef264"];
function dark(h){var n=parseInt(h.slice(1),16);return "rgb("+Math.max(0,((n>>16)&255)-35)+","+Math.max(0,((n>>8)&255)-35)+","+Math.max(0,(n&255)-35)+")";}
var activeCat=null,activeSku=null,kpiPick=null;
var valueLabelPlugin={id:"valueLabelPlugin",afterDatasetsDraw:function(chart){var ctx=chart.ctx;ctx.save();ctx.textAlign="center";ctx.textBaseline="bottom";ctx.fillStyle="#475569";ctx.font="600 11px sans-serif";chart.data.datasets.forEach(function(ds,di){if(ds.type&&ds.type!=="bar")return;var meta=chart.getDatasetMeta(di);meta.data.forEach(function(bar,i){var v=ds.data[i];if(v==null)return;ctx.fillText(fmt(v),bar.x,bar.y-5);});});ctx.restore();}};
var piePlugin={id:"piePlugin",afterDraw:function(chart){var ctx=chart.ctx;var o=(chart.options.plugins||{}).pieText;if(!o||o.hide)return;var a=chart.chartArea,cx=(a.left+a.right)/2,cy=(a.top+a.bottom)/2;
 if(o.label!==undefined){ctx.save();ctx.textAlign="center";ctx.textBaseline="middle";ctx.fillStyle="#0f172a";ctx.font="800 20px sans-serif";ctx.fillText(o.label,cx,cy-6);ctx.font="11px sans-serif";ctx.fillStyle="#64748b";ctx.fillText(o.sub||"MB",cx,cy+14);ctx.restore();}
 chart.data.datasets.forEach(function(ds,di){var m=chart.getDatasetMeta(di);if(!ds._sd||!m||!m.data)return;m.data.forEach(function(arc,i){var d=ds._sd[i];if(!d||d.pct<5)return;var an=(arc.startAngle+arc.endAngle)/2,r=(arc.outerRadius+arc.innerRadius)/2,x=cx+Math.cos(an)*r,y=cy+Math.sin(an)*r;ctx.save();ctx.fillStyle="#fff";ctx.font="bold 12px sans-serif";ctx.textAlign="center";ctx.textBaseline="middle";ctx.fillText(d.pct.toFixed(0)+"%",x,y);ctx.restore();});});}};
function sel(){return D[cur];}
function label(){return D[cur]?D[cur].label:cur;}
function C(){
  // series ตาม channel: MT ใช้ s_* ; ช่องอื่น fallback monthly-only (LE/LY/AOP ว่าง)
  return sel();
}
function load(){
 fetch("data_channels.json?v="+Date.now()).then(r=>r.json()).then(function(data){
   D=data;lastClosed=8;fy=2026;
   // channels
   var sc=document.getElementById("selChannel");sc.innerHTML="";
   (D._channels||[]).forEach(function(ch){var o=document.createElement("option");o.value=ch.id;o.textContent=ch.label;sc.appendChild(o);});
   sc.value="MT";cur="MT";
   sc.onchange=function(){cur=sc.value;activeCat=null;activeSku=null;kpiPick=null;renderAll();};
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
function hasS(ch){ return ch && ch.s_actual && ch.s_le && ch.s_ly && ch.s_aop; }
function renderAll(){
  try{
    var c=C();setV("tgM",label());setV("tgTbl",label());setV("tgCat",label());setV("tgSku",label());
    var ok=hasS(c);
    var rangeTxt=MONTHS[selFrom-1]+"-"+MONTHS[selTo-1];
    setV("lblAct",rangeTxt);setV("lblLE",rangeTxt);setV("lblLY",rangeTxt);setV("lblAOP",rangeTxt);
    if(!ok){
      // ช่องอื่น: มีแค่ monthly Actual (ไม่มี LE/LY/AOP)
      var a=rangeSum(c.s_actual||c.monthly,selFrom,selTo);
      setT("kACT",fmt(a)+' <small>MB</small>');setT("kACTuntil",MONTHS[lastClosed-1]+" "+fy);
      setT("kLE","—");setT("kGap","—");
      setT("kLY","—");setT("kLYcmp","No LY data");
      setT("kAOP","—");setT("kAOPcmp","No AOP data");
    } else {
      var a=rangeSum(c.s_actual,selFrom,selTo);
      var le=0;for(var m=selFrom;m<=selTo;m++){ if(m<lastClosed) le+=num(c.s_actual[String(m)]); else le+=num(c.s_le[String(m)]); }
      var ly=rangeSum(c.s_ly,selFrom,selTo);
      var aop=rangeSum(c.s_aop,selFrom,selTo);
      setT("kACT",fmt(a)+' <small>MB</small>');setT("kACTuntil",MONTHS[lastClosed-1]+" "+fy);
      setT("kLE",fmt(le)+' <small>MB</small>');setT("kGap",fmt(Math.max(0,le-a))+" MB");
      var lyVs=pct(ly,aop);
      setT("kLY",fmt(ly)+' <small>MB</small>');setT("kLYcmp",lyVs==null?"No AOP data":"vs AOP "+lyVs.toFixed(1)+"%");
      var vs= aop? (le/aop*100) : 0;
      var leVs=pct(le,aop);
      setT("kAOP",fmt(vs)+' <small>%</small>');setT("kAOPcmp",leVs==null?"No AOP data":"LE "+leVs.toFixed(1)+"% vs AOP");
    }
    setV("hDate",MONTHS[lastClosed-1]+" "+fy);
    drawAnnual();drawMonthly();drawTable();
    drawCategory(c);drawSku(c);
    if(kpiPick)applyKpiPick();
  }catch(e){setT("tgM","ERR: "+e.message);}
}
function chgBg(m,sel){return m===sel?"#e11d48":"#f87171";}
function sv(c,key,m){ if(c&&c[key]) return num(c[key][String(m)]); if(c&&c.monthly) return num(c.monthly[String(m)]); return 0; }
function currentFocus(c){if(activeSku&&c.monthly_by_sku&&c.monthly_by_sku[activeSku])return {name:activeSku,series:c.monthly_by_sku[activeSku],annual:c.annual_by_sku&&c.annual_by_sku[activeSku],history:c.history_by_sku&&c.history_by_sku[activeSku]};if(activeCat&&c.monthly_by_cat&&c.monthly_by_cat[activeCat])return {name:activeCat,series:c.monthly_by_cat[activeCat],annual:c.annual_by_cat&&c.annual_by_cat[activeCat],history:c.history_by_cat&&c.history_by_cat[activeCat]};return null;}
function pickKpi(type,value){if(type==="year"){if(kpiPick&&kpiPick.type==="month")kpiPick={type:"yearMonth",year:String(value),month:+kpiPick.value};else if(kpiPick&&kpiPick.type==="yearMonth")kpiPick=(String(kpiPick.year)===String(value))?{type:"month",value:kpiPick.month}:{type:"yearMonth",year:String(value),month:kpiPick.month};else{kpiPick=(kpiPick&&kpiPick.type==="year"&&String(kpiPick.value)===String(value))?null:{type:"year",value:String(value)};}}else{if(kpiPick&&kpiPick.type==="year")kpiPick={type:"yearMonth",year:String(kpiPick.value),month:+value};else if(kpiPick&&kpiPick.type==="yearMonth")kpiPick=(+kpiPick.month===+value)?{type:"year",value:String(kpiPick.year)}:{type:"yearMonth",year:String(kpiPick.year),month:+value};else kpiPick=(kpiPick&&kpiPick.type==="month"&&+kpiPick.value===+value)?null:{type:"month",value:+value};}setTimeout(renderAll,0);}
function applyKpiPick(){var c=C(),focus=currentFocus(c);if(kpiPick.type==="year"){var y=String(kpiPick.value),annual=(focus&&focus.annual)||c.annual||{},prev=String(+y-1);setV("lblAct",y);setT("kACT",fmt(annual[y])+' <small>MB</small>');setT("kACTuntil","Full year / YTD");setV("lblLE",y);setT("kLE","—");setT("kGap","—");setV("lblLY",prev);setT("kLY",annual[prev]!=null?fmt(annual[prev])+' <small>MB</small>':"—");setT("kLYcmp","Previous year");setV("lblAOP",y);setT("kAOP","—");setT("kAOPcmp","No annual AOP");return;}
 var m=kpiPick.type==="yearMonth"?+kpiPick.month:+kpiPick.value,y=kpiPick.type==="yearMonth"?+kpiPick.year:fy,history=(focus&&focus.history)||c.history||{},actual=num(history[String(y)]&&history[String(y)][String(m)]),ly=history[String(y-1)]?num(history[String(y-1)][String(m)]):null,current=(y===fy),le=(!focus&&current&&hasS(c))?(m<lastClosed?actual:sv(c,"s_le",m)):null,aop=(!focus&&current&&hasS(c))?sv(c,"s_aop",m):null;["lblAct","lblLE","lblLY","lblAOP"].forEach(function(id){setV(id,MONTHS[m-1]+" "+y);});setT("kACT",fmt(actual)+' <small>MB</small>');setT("kACTuntil",MONTHS[m-1]+" "+y);setT("kLE",le==null?"\u2014":fmt(le)+' <small>MB</small>');setT("kGap",le==null?"\u2014":fmt(Math.max(0,le-actual))+" MB");setT("kLY",ly==null?"\u2014":fmt(ly)+' <small>MB</small>');setT("kLYcmp",ly==null?"No LY data":"Monthly LY "+(y-1));setT("kAOP",aop?fmt(le/aop*100)+' <small>%</small>':"\u2014");setT("kAOPcmp",aop?"Monthly LE vs AOP":"No AOP data");}

function drawAnnual(){destroy("cAnnual");var c=C(),focus=currentFocus(c),annual=(focus&&focus.annual)||c.annual||{},history=(focus&&focus.history)||c.history||{},years=Object.keys(annual).sort(),month=kpiPick&&(kpiPick.type==="month"?+kpiPick.value:(kpiPick.type==="yearMonth"?+kpiPick.month:null)),vals=years.map(function(y){return month?num(history[y]&&history[y][String(month)]):num(annual[y]);}),selectedYear=kpiPick&&(kpiPick.type==="year"?String(kpiPick.value):(kpiPick.type==="yearMonth"?String(kpiPick.year):null)),colors=years.map(function(y){return selectedYear===y?"#be123c":"#f9a8d4";}),base=focus?(label()+" • "+focus.name):label();setV("tgY",base+(month?" • "+MONTHS[month-1]:""));charts["cAnnual"]=new Chart(document.getElementById("cAnnual"),{type:"bar",data:{labels:years,datasets:[{label:"Actual",data:vals,backgroundColor:colors,borderRadius:6,barPercentage:.55}]},plugins:[valueLabelPlugin],options:{responsive:true,maintainAspectRatio:false,onClick:function(e,els){if(els&&els.length)pickKpi("year",years[els[0].index]);},plugins:{legend:{display:false},tooltip:{callbacks:{label:function(x){return (month?MONTHS[month-1]:"Actual")+": "+fmt(x.parsed.y)+" MB";}}}},scales:{y:{beginAtZero:true,grace:"15%"},x:{grid:{display:false}}}}});}

function drawMonthly(){destroy("cMonthly");var c=C();var lb=[],act=[],le=[],ly=[],aop=[];
 var focusInfo=currentFocus(c),focus=focusInfo&&focusInfo.series,focusName=focusInfo&&focusInfo.name||"",viewYear=kpiPick&&(kpiPick.type==="year"?+kpiPick.value:(kpiPick.type==="yearMonth"?+kpiPick.year:fy))||fy,history=(focusInfo&&focusInfo.history)||c.history||{},historical=(viewYear!==fy),seriesOk=hasS(c)&&!focus&&!historical;
 setV("tgM",(focusName?(label()+" • "+focusName):label())+(historical?" • "+viewYear:""));
 ["lgLE","lgAOP"].forEach(function(id){var e=document.getElementById(id);if(e)e.style.display=(focus||historical)?"none":"inline-flex";});
 var actualSeries=historical?(history[String(viewYear)]||{}):(focus?focus.actual:null),lySeries=historical?(history[String(viewYear-1)]||{}):(focus?focus.ly:null),maxMonth=historical?12:lastClosed,latestVisible=Math.min(selTo,maxMonth);
 for(var i=1;i<=12;i++){var m=i;lb.push((m>=selFrom&&m<=selTo)?MONTHS[m-1]:"");act.push((m<=maxMonth&&m>=selFrom&&m<=selTo)?((focus||historical)?num(actualSeries[String(m)]):sv(c,"s_actual",m)):null);le.push((seriesOk&&m>=lastClosed&&m>=selFrom&&m<=selTo)?sv(c,"s_le",m):null);ly.push(((seriesOk||focus||historical)&&m>=selFrom&&m<=selTo)?((focus||historical)?num(lySeries[String(m)]):sv(c,"s_ly",m)):null);aop.push((seriesOk&&m>=selFrom&&m<=selTo)?sv(c,"s_aop",m):null);}
 charts["cMonthly"]=new Chart(document.getElementById("cMonthly"),{data:{labels:lb,datasets:[
  {type:"bar",label:"Actual",data:act,backgroundColor:act.map(function(v,i){var picked=kpiPick&&(kpiPick.type==="month"?+kpiPick.value:(kpiPick.type==="yearMonth"?+kpiPick.month:null));return picked===i+1?"#7f1d1d":((i+1)===latestVisible?"#be123c":"#f87171");}),borderRadius:5,barPercentage:.65},
  {type:"bar",label:"LE",data:le,backgroundColor:"#cbd5e1",borderRadius:5,barPercentage:.65},
  {type:"line",label:"LY",data:ly,borderColor:"#64748b",backgroundColor:"#64748b",borderDash:[5,3],fill:false,tension:.3,pointRadius:2,borderWidth:2},
  {type:"line",label:"AOP",data:aop,borderColor:"#2563eb",backgroundColor:"#2563eb",fill:false,tension:.3,pointRadius:2,borderWidth:2}]},
  options:{responsive:true,maintainAspectRatio:false,interaction:{mode:"index",intersect:false},onClick:function(e,els){if(els&&els.length&&els[0].datasetIndex<=1)pickKpi("month",els[0].index+1);},
   plugins:{legend:{display:false},tooltip:{callbacks:{label:function(x){return x.dataset.label+": "+fmt(x.parsed.y)+" MB";}}}},
   scales:{y:{beginAtZero:true,ticks:{callback:function(v){return fmt(v);}}},x:{grid:{display:false}}}}});
}
function drawTable(){var c=C();var h='<thead><tr><th>Month</th><th>Actual</th><th>LE</th><th>LY 2025</th><th>AOP</th></tr></thead><tbody>';
 var seriesOk=hasS(c),ta=0,tl=0,ty=0,ta2=0;
 for(var m=1;m<=12;m++){if(m<selFrom||m>selTo)continue;
  var ok=seriesOk;
  var act=(m<=lastClosed)?sv(c,"s_actual",m):null;
  var le = ok ? ((m<lastClosed)?act:sv(c,"s_le",m)) : null;
  var ly=ok?sv(c,"s_ly",m):null;
  var aop=ok?sv(c,"s_aop",m):null;
  ta+=act||0;if(ok){tl+=le||0;ty+=ly||0;ta2+=aop||0;}
  h+="<tr class=\"r"+(m===selTo?" notlast":"")+"\"><td>"+MONTHS[m-1]+"</td><td>"+(act?fmt(act):"—")+"</td><td>"+(le==null?"—":fmt(le))+"</td><td>"+(ly==null?"—":fmt(ly))+"</td><td>"+(aop==null?"—":fmt(aop))+"</td></tr>";}
 h+='<tr class="total"><td>Total</td><td>'+fmt(ta)+'</td><td>'+(seriesOk?fmt(tl):'—')+'</td><td>'+(seriesOk?fmt(ty):'—')+'</td><td>'+(seriesOk?fmt(ta2):'—')+'</td></tr></tbody>';
 setT("mTable",h);}
function drawCategory(c){var el=document.getElementById("cCat");if(!el||!Chart)return;destroy("cCat");
 var C2=(c.category&&c.category.items)?c.category.items:[];if(!C2.length){setT("catLegend","No data");return;}
 var tot=c.total_mb;var vals=C2.map(function(x){return x.mb;});var colors=C2.map(function(x,i){return (activeCat===x.name)?dark(CATCOLR[i%CATCOLR.length]):CATCOLR[i%CATCOLR.length];});
 var _sd=C2.map(function(x){return {pct:x.pct||0};});
 charts["cCat"]=new Chart(el,{type:"doughnut",data:{labels:C2.map(function(x){return x.name;}),datasets:[{data:vals,backgroundColor:colors,borderWidth:1,borderColor:"#fff",_sd:_sd}]},plugins:[piePlugin],
  options:{responsive:true,maintainAspectRatio:false,cutout:"60%",onClick:function(ev,els){if(els&&els.length){var idx=els[0].index;if(C2[idx])toggleCat(C2[idx].name);}},
   plugins:{title:{display:true,text:"YTD by Category - "+label()+" (MB)",font:{size:14,weight:"bold"}},pieText:{label:fmt(tot),sub:"MB"},legend:{display:false},tooltip:{callbacks:{label:function(x){var it=C2[x.dataIndex];return it.name+": "+fmt(it.mb)+" MB ("+it.pct+"%)";}}}}}});
 var lg="";C2.forEach(function(x,i){var bg=(activeCat===x.name)?"background:#fef9c3;":"";lg+='<div class="catrow" data-cat="'+x.name+'" style="'+bg+'display:flex;justify-content:space-between;align-items:center;padding:5px 0;border-bottom:1px solid #f1f5f9;cursor:pointer;border-left:4px solid '+CATCOLR[i%CATCOLR.length]+';padding-left:8px"><span>'+x.name+'</span><b>'+fmt(x.mb)+' MB · '+x.pct+'%</b></div>';});
 setT("catLegend",lg);
 var rows=document.querySelectorAll(".catrow");for(var i=0;i<rows.length;i++){(function(r){r.onclick=function(){toggleCat(r.dataset.cat);};})(rows[i]);}}
function toggleCat(cat){activeCat=(activeCat===cat)?null:cat;activeSku=null;kpiPick=null;renderAll();}
function toggleSku(name,cat){activeSku=(activeSku===name)?null:name;if(cat)activeCat=cat;kpiPick=null;renderAll();}
function drawSku(c){var tot=c.total_mb||0;var list;
 if(activeCat){list=(c.top_sku_by_cat&&c.top_sku_by_cat[activeCat])?c.top_sku_by_cat[activeCat]:[];}
 else{list=c.top_sku_all||[];}
 if(!list.length){setT("skuTable",'<thead><tr><th>SKU</th><th>MB</th></tr></thead><tbody><tr><td>No SKU</td><td>—</td></tr></tbody>');return;}
 var h='<thead><tr><th>#</th><th style="text-align:left">SKU</th><th>Category</th><th style="text-align:right">MB</th><th style="text-align:right">%</th></tr></thead><tbody>';
 var shown=list.slice(0,10);shown.forEach(function(s,i){var cat=s.cat||activeCat||"";var p=tot?(s.mb/tot*100):0;var bg=activeSku===s.name?'background:#fef3c7;':'';h+='<tr class="skuro" data-i="'+i+'" style="cursor:pointer;'+bg+'"><td>'+(i+1)+'</td><td style="max-width:220px;white-space:normal;text-align:left">'+s.name+'</td><td>'+cat+'</td><td style="text-align:right">'+fmt(s.mb)+'</td><td style="text-align:right">'+p.toFixed(1)+'%</td></tr>';});
 h+='</tbody>';setT("skuTable",h);
 var rows=document.querySelectorAll(".skuro");for(var i=0;i<rows.length;i++){(function(r){r.onclick=function(){var s=shown[+r.dataset.i];if(s)toggleSku(s.name,s.cat||activeCat||"");};})(rows[i]);}}
load();
