const MONTHS=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
let D=null,S=null,SW=null,charts={},cur="MT",selChans=["MAKRO","LOTUS'"],selFrom=1,selTo=8,lastClosed=8,fy=2026;
var viewYear=fy; // Year filter state; defaults to fy (2026) so the first-open view is unchanged.
function num(v){return (v==null||isNaN(v))?0:Number(v);}
function fmt(v){return (v==null||isNaN(v))?"-":Number(v).toLocaleString("en-US",{minimumFractionDigits:1,maximumFractionDigits:1});}
// MB with 2 decimal places, used by the daily Ship-to party table.
function fmt2(v){return (v==null||isNaN(v))?"-":Number(v).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});}
function pct(a,b){a=num(a);b=num(b);if(!b)return null;return ((a-b)/b*100);}
function destroy(id){if(charts[id]){charts[id].destroy();charts[id]=null;}}
function setT(id,h){var e=document.getElementById(id);if(e)e.innerHTML=h;}
function setV(id,t){var e=document.getElementById(id);if(e)e.textContent=t;}
// "Latest sales as of" date is auto-derived from the source workbook
// (data_channels.json => _lastSalesDate, format YYYY-MM-DD). Falls back to the
// closed-month label if the field is missing, so it is never hardcoded.
function salesDateTxt(){
  var v=D&&D._lastSalesDate;if(typeof v==="string"&&v.length>=10){var p=v.split("-");var y=+p[0],m=+p[1],dd=+p[2];if(y&&m>=1&&m<=12&&dd>=1&&dd<=31)return dd+" "+MONTHS[m-1]+" "+y;}
  return MONTHS[lastClosed-1]+" "+fy;
}
var CATCOLR=["#F2E9D8","#59718A","#2F6F6D","#E07A47","#D4A72C","#8FA68F","#C47C8A","#5B3A62"];
function dark(h){var n=parseInt(h.slice(1),16);return "rgb("+Math.max(0,((n>>16)&255)-35)+","+Math.max(0,((n>>8)&255)-35)+","+Math.max(0,(n&255)-35)+")";}
var activeCat=null,activeSku=null,kpiPick=null;
var activeShipTo=null; // clicked Ship-to party linked filter: drives KPI/charts until re-clicked to clear
var shipToUnit="mb"; // table-local display unit; does not alter any other dashboard visual
var skuWeeklyUnit="mb",skuWeeklySearch=""; // local display/search state for the weekly SKU table
var catSort=null,skuSort=null; // {col:'ly'|'actual'|'growth'|'pct'|'contrib', dir:'desc'|'asc'} or null
var channelExportRows=[];
function exportContext(){return {Channel:label(),From:MONTHS[selFrom-1],To:MONTHS[selTo-1],Year:viewYear,Category:activeCat||"All",SKU:activeSku||"All",ShipTo:activeShipTo||"All",LatestSales:salesDateTxt()};}
function exportCell(v,raw){if(raw!==undefined&&raw!==null&&raw!==""){var rv=Number(raw);if(!isNaN(rv))return rv;}var s=String(v==null?"":v).trim().replace(/\s*[▼▲]\s*/g,"").replace(/, sorted.*$/i,"");if(s==="—"||s==="-")return "";var p=s.match(/^([+-]?[\d,]+(?:\.\d+)?)%$/);if(p)return Number(p[1].replace(/,/g,""))/100;var n=s.match(/^([+-]?[\d,]+(?:\.\d+)?)$/);if(n)return Number(n[1].replace(/,/g,""));return s;}
function exportRawAttr(v){return v==null||isNaN(v)?"":" data-export-value=\""+String(Number(v))+"\"";}
function tableExportRows(id){var t=document.getElementById(id),out=[];if(!t)return out;Array.prototype.forEach.call(t.querySelectorAll("tr"),function(tr){var row=[];Array.prototype.forEach.call(tr.querySelectorAll("th,td"),function(td){row.push(exportCell(td.textContent,td.getAttribute("data-export-value")));});out.push(row);});return out;}
function chartExportRows(id){var ch=charts[id];if(!ch)return [];var ds=ch.data.datasets||[],head=["Period"].concat(ds.map(function(x){return x.label||"Value";})),out=[head];(ch.data.labels||[]).forEach(function(lb,i){if(lb==="")return;var row=[lb];ds.forEach(function(x){var v=x.data&&x.data[i];row.push(v==null?"":num(v));});out.push(row);});return out;}
function safeFilePart(s){return String(s||"").replace(/[^A-Za-z0-9_-]+/g,"_").replace(/^_+|_+$/g,"");}
function exportSection(section){
  if(typeof XLSX==="undefined"){alert("Excel export library is not available. Please check your connection and refresh.");return;}
  var names={annual:"Annual_Sales",monthly:"Monthly_Performance",monthly_table:"Monthly_Table",category:"Sales_by_Category",category_performance:"Category_Performance",sku:"SKU_Performance",channel:"Sales_by_Channel",ship_to:"Ship_to_Weekly",sku_weekly:"SKU_Weekly"};
  var rows=[];
  if(section==="annual")rows=chartExportRows("cAnnual");
  else if(section==="monthly")rows=chartExportRows("cMonthly");
  else if(section==="monthly_table")rows=tableExportRows("mTable");
  else if(section==="category")rows=chartExportRows("cCat");
  else if(section==="category_performance")rows=tableExportRows("catPerfTable");
  else if(section==="sku")rows=tableExportRows("skuTable");
  else if(section==="channel"){rows=[["Channel","Actual (MB)","Contribution","LY 2025 (MB)","Gap (MB)","Growth vs LY"]];channelExportRows.forEach(function(r){rows.push([r.label,r.actual,r.contribution/100,r.ly,r.gap,r.growth==null?"":r.growth/100]);});}
  else if(section==="ship_to"){rows=tableExportRows("shipToTable");
    var colHdr=rows.length?rows[0]:[];if(colHdr.length)colHdr[0]="Ship-to party";
  }
  else if(section==="sku_weekly"){buildSkuWeeklyExportRows();rows=skuWeeklyExportRows;
  }
  if(!rows.length){alert("No data available for export.");return;}
  var wb=XLSX.utils.book_new(),ws=XLSX.utils.aoa_to_sheet(rows),ctx=exportContext(),meta=[["Filter","Value"]];Object.keys(ctx).forEach(function(k){meta.push([k,ctx[k]]);});
  ws["!cols"]=rows[0].map(function(_,i){var w=12;rows.forEach(function(r){w=Math.max(w,String(r[i]==null?"":r[i]).length+2);});return {wch:Math.min(w,42)};});
  XLSX.utils.book_append_sheet(wb,ws,"Data");XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(meta),"Filters");
  var file="CHAOSUA_"+(names[section]||"Dashboard")+"_"+MONTHS[selFrom-1]+"-"+MONTHS[selTo-1]+"_"+fy+"_"+safeFilePart(label())+".xlsx";XLSX.writeFile(wb,file);
}
function pptChartSeries(chart,datasetFilter){var labels=chart&&chart.data&&chart.data.labels||[],keep=[];labels.forEach(function(x,i){if(x!=="")keep.push(i);});return (chart&&chart.data&&chart.data.datasets||[]).filter(datasetFilter||function(){return true;}).map(function(ds){return {name:ds.label||"Sales",labels:keep.map(function(i){return labels[i];}),values:keep.map(function(i){var v=ds.data&&ds.data[i];return v==null?0:num(v);})};});}
function exportEditablePpt(section){
  if(typeof PptxGenJS==="undefined"){alert("Editable PowerPoint library is not available. Please check your connection and refresh.");return;}
  var chartId=section==="annual"?"cAnnual":section==="monthly"?"cMonthly":"cCat",chart=charts[chartId];if(!chart){alert("No chart data available for export.");return;}
  var pptx=new PptxGenJS();pptx.layout="LAYOUT_WIDE";pptx.author="CHAOSUA";pptx.company="CHAOSUA Foods Industry Public Company Limited";pptx.subject="MT Sales Dashboard";
  var slide=pptx.addSlide();slide.background={color:"FFFFFF"};
  var titles={annual:"Annual Actual Sales",monthly:"Monthly Sales Performance",category:"Sales by Category"},title=titles[section]||"Sales Chart",sub=label()+" | "+MONTHS[selFrom-1]+"–"+MONTHS[selTo-1]+" "+fy+(activeCat?" | "+activeCat:"")+(activeSku?" | "+activeSku:"");
  slide.addText(title,{x:.55,y:.28,w:12.2,h:.38,fontFace:"Arial",fontSize:22,bold:true,color:"0F172A",margin:0});slide.addText(sub,{x:.55,y:.72,w:12.2,h:.25,fontFace:"Arial",fontSize:10,color:"64748B",margin:0});
  var common={x:.55,y:1.08,w:12.2,h:5.65,showTitle:false,showLegend:true,legendPos:"b",showValue:true,showCatName:false,showPercent:false,fontFace:"Arial",chartColors:["525252","CBD5E1","64748B","2563EB","E07A47","D4A72C","8FA68F"],showBorder:false,catAxisLabelFontSize:10,valAxisLabelFontSize:10,valAxisTitle:"MB",showValue:true};
  if(section==="annual"){common.showLegend=false;slide.addChart(pptx.ChartType.bar,pptChartSeries(chart),common);}
  else if(section==="category"){common.showLegend=true;common.legendPos="r";common.holeSize=58;common.showPercent=true;common.showValue=false;slide.addChart(pptx.ChartType.doughnut,pptChartSeries(chart),common);}
  else {var bars=pptChartSeries(chart,function(ds){return (ds.type||"")==="bar";}),lines=pptChartSeries(chart,function(ds){return (ds.type||"")==="line";});common.showValue=false;slide.addChart([{type:pptx.ChartType.bar,data:bars},{type:pptx.ChartType.line,data:lines}],common);}
  slide.addText("Source: CHAOSUA MT Sales Dashboard | Editable chart: right-click chart and choose Edit Data",{x:.55,y:7.18,w:12.2,h:.18,fontFace:"Arial",fontSize:8,color:"94A3B8",margin:0});
  var file="CHAOSUA_"+title.replace(/\s+/g,"_")+"_"+MONTHS[selFrom-1]+"-"+MONTHS[selTo-1]+"_"+fy+"_"+safeFilePart(label())+".pptx";pptx.writeFile({fileName:file}).catch(function(e){alert("PowerPoint export failed: "+e.message);});
}
var valueLabelPlugin={id:"valueLabelPlugin",afterDatasetsDraw:function(chart){var ctx=chart.ctx;ctx.save();ctx.textAlign="center";ctx.textBaseline="bottom";ctx.fillStyle="#475569";ctx.font="700 12px sans-serif";chart.data.datasets.forEach(function(ds,di){var datasetType=ds.type||chart.config.type;if(datasetType!=="bar")return;var meta=chart.getDatasetMeta(di);if(!meta||!meta.data)return;meta.data.forEach(function(bar,i){var v=ds.data[i];if(v==null||isNaN(v)||v===0)return;var x=bar.x,y=bar.y;if(x==null||y==null)return;ctx.fillText(num(v).toFixed(0),x,y-6);});});ctx.restore();}};
if(typeof Chart!=="undefined"){try{Chart.register(valueLabelPlugin);}catch(e){}}
var piePlugin={id:"piePlugin",afterDraw:function(chart){var ctx=chart.ctx;var o=(chart.options.plugins||{}).pieText;if(!o||o.hide)return;var a=chart.chartArea,cx=(a.left+a.right)/2,cy=(a.top+a.bottom)/2;
 if(o.label!==undefined){ctx.save();ctx.textAlign="center";ctx.textBaseline="middle";ctx.fillStyle="#0f172a";ctx.font="700 22px Arial,sans-serif";ctx.fillText(String(o.label),cx,cy);ctx.restore();}}};
function sourceData(){return D;}
function sel(){return sourceData()[cur];}
function selAll(){
  var SD=sourceData(),out={label:selChans.map(function(i){return SD[i]&&SD[i].label||i;}).join(" + "),monthly:{},history:{},category:{items:[]},annual:{},s_actual:{},s_le:{},s_ly:{},s_aop:{},monthly_by_cat:{},history_by_cat:{},annual_by_cat:{},monthly_by_sku:{},history_by_sku:{},annual_by_sku:{},_planningOK:true};
  selChans.forEach(function(id){var c=SD[id];if(!c)return;
    if(!hasS(c))out._planningOK=false;
    for(var m=1;m<=12;m++){out.monthly[String(m)]=num(out.monthly[String(m)])+num(c.monthly&&c.monthly[String(m)]||0);
      if(c.s_actual)out.s_actual[String(m)]=num(out.s_actual[String(m)])+num(c.s_actual[String(m)]||0);
      if(c.s_le)out.s_le[String(m)]=num(out.s_le[String(m)])+num(c.s_le[String(m)]||0);
      if(c.s_ly)out.s_ly[String(m)]=num(out.s_ly[String(m)])+num(c.s_ly[String(m)]||0);
      if(c.s_aop)out.s_aop[String(m)]=num(out.s_aop[String(m)])+num(c.s_aop[String(m)]||0);}
    // Merge prior-year history for ALL available years so the annual chart can show 2023/2024 etc. too.
    var hys=c.history||{};Object.keys(hys).forEach(function(y){out.history[y]=out.history[y]||{};var h=hys[y]||{};for(var m=1;m<=12;m++)out.history[y][String(m)]=num(out.history[y][String(m)])+num(h[String(m)]||0);});
    var an=c.annual||{};Object.keys(an).forEach(function(y){out.annual[y]=num(out.annual[y]||0)+num(an[y]);});
    var ci=c.category&&c.category.items||[];ci.forEach(function(it){var ex=null;for(var k=0;k<out.category.items.length;k++)if(out.category.items[k].name===it.name){ex=out.category.items[k];break;}if(ex)ex.mb=ex.mb+it.mb;else out.category.items.push({name:it.name,mb:it.mb,pct:it.pct});});
    // Merge category series for linked category filtering.
    var mbc=c.monthly_by_cat||{},hbc=c.history_by_cat||{},abc=c.annual_by_cat||{};mergeFocusMaps(out.monthly_by_cat,out.history_by_cat,out.annual_by_cat,mbc,hbc,abc);
    var mbs=c.monthly_by_sku||{},hbs=c.history_by_sku||{},abs=c.annual_by_sku||{};mergeFocusMaps(out.monthly_by_sku,out.history_by_sku,out.annual_by_sku,mbs,hbs,abs);
  });
  out.total_mb=0;for(var m=1;m<=12;m++)out.total_mb+=num(out.monthly[String(m)]||0);
  var tsa=[],tsbc={};selChans.forEach(function(id){var cs=SD[id];if(!cs)return;if(cs.top_sku_all)tsa=tsa.concat(cs.top_sku_all);var t=cs.top_sku_by_cat||{};Object.keys(t).forEach(function(cat){tsbc[cat]=(tsbc[cat]||[]).concat(t[cat]);});});out.top_sku_all=mergeSkuList(tsa);Object.keys(tsbc).forEach(function(cat){tsbc[cat]=mergeSkuList(tsbc[cat]);});out.top_sku_by_cat=tsbc;
  return out;
}
function mergeFocusMaps(outM,outH,outA,srcM,srcH,srcA){Object.keys(srcM).forEach(function(key){outM[key]=outM[key]||{actual:{},ly:{}};if(srcM[key].cat&&!outM[key].cat)outM[key].cat=srcM[key].cat;for(var m=1;m<=12;m++){outM[key].actual[String(m)]=num(outM[key].actual[String(m)])+num((srcM[key].actual||{})[String(m)]);outM[key].ly[String(m)]=num(outM[key].ly[String(m)])+num((srcM[key].ly||{})[String(m)]);}});Object.keys(srcH).forEach(function(key){outH[key]=outH[key]||{};Object.keys(srcH[key]||{}).forEach(function(y){outH[key][y]=outH[key][y]||{};for(var m=1;m<=12;m++)outH[key][y][String(m)]=num(outH[key][y][String(m)])+num(srcH[key][y][String(m)]);});});Object.keys(srcA).forEach(function(key){outA[key]=outA[key]||{};Object.keys(srcA[key]||{}).forEach(function(y){outA[key][y]=num(outA[key][y])+num(srcA[key][y]);});});}
function mergeSkuList(list){var map={};(list||[]).forEach(function(s){var k=s.name;if(!map[k])map[k]={name:s.name,mb:0,cat:s.cat||""};map[k].mb+=num(s.mb);});return Object.keys(map).map(function(k){return map[k];}).sort(function(a,b){return b.mb-a.mb;});}
function label(){var SD=sourceData(),ids=(D&&D._channels||[]).filter(function(ch){return ch.id!=="MT";}).map(function(ch){return ch.id;}),base;if(ids.length&&selChans&&selChans.length===ids.length)base="All Channel";else if(selChans&&selChans.length>1)base=selChans.map(function(i){return SD[i]&&SD[i].label||i;}).join(" + ");else base=SD[cur]?SD[cur].label:cur;return base;}
function C(){
  if(activeShipTo)return shipToChannel(activeShipTo);
  if(selChans&&selChans.length>1)return selAll();
  return sel();
}
// ---- Ship-to party linked filter -------------------------------
// When a Ship-to row is clicked, KPI, charts, Category, SKU and the Sales-by-
// Channel cards all focus on that single location. The Ship-to TABLE itself is
// deliberately NOT reduced to the selected row (it is the filter control, and
// would otherwise collapse into a meaningless self-filter loop); only the
// clicked row is highlighted, and clicking it again clears the filter. Every
// figure is derived from the SAME facts and filters used to build the table, so
// the selected row Total reconciles with the Actual KPI by construction.
function shipToBasisSet(){ // channels currently in scope as a lookup object
  var s={};(selChans||[]).forEach(function(id){s[id]=1;});return s;
}
function shipToFactsPass(r){ // fact [shipTo, chan, year, month, cat, sku, val] under the CURRENT table filters
  if(!shipToBasisSet()[r[1]])return false;
  if(String(r[2])!==String(viewYear))return false;
  var m=Number(r[3]);if(m<num(selFrom)||m>num(selTo))return false;
  if(activeCat&&r[4]!==activeCat)return false;
  if(activeSku&&r[5]!==activeSku)return false;
  return true;
}
function shipToAggFor(name,year,respectDrill){ // aggregate one Ship-to into a channel-like series object
  year=(year==null?viewYear:Number(year));
  var byChan={},selSet=shipToBasisSet(),from=num(selFrom),to=num(selTo);
  var months={},perCat={},perSku={},total=0;
  (S.facts||[]).forEach(function(r){
    if(r[0]!==name)return;
    if(!selSet[r[1]])return;
    if(String(r[2])!==String(year))return;
    var m=Number(r[3]);if(m<from||m>to)return;
    if(respectDrill){if(activeCat&&r[4]!==activeCat)return;if(activeSku&&r[5]!==activeSku)return;}
    var v=num(r[6]);months[m]=num(months[m])+v;total+=v;
    perCat[r[4]]=num(perCat[r[4]]||0)+v;
    perSku[r[5]]=num(perSku[r[5]]||0)+v;
    byChan[r[1]]=num(byChan[r[1]]||0)+v;
  });
  return {months:months,total:total,perCat:perCat,perSku:perSku,byChan:byChan};
}
function shipToChannel(name){
  var c={label:name,_planningOK:false,shipTo:name};
  var act=shipToAggFor(name,viewYear,true);   // respects active cat/sku => reconciles with the table row
  var ly=shipToAggFor(name,viewYear-1,true);  // same basis, prior year
  var hist={},annual={};
  // Actual series for the current view year (capped at the closed month like the KPI).
  var actualSeries={},effClosed=isHistoricalYear()?12:lastClosed;
  for(var m=1;m<=12;m++){if(m<=effClosed)actualSeries[String(m)]=num(act.months[String(m)]);}
  c.monthly=actualSeries;c.s_actual=actualSeries;c.s_le={};c.s_ly=ly.months||{};c.s_aop={};
  c.total_mb=act.total;
  // Per-year annual (sum over the visible month range) for the annual chart.
  (S.facts||[]).forEach(function(r){
    if(r[0]!==name||!shipToBasisSet()[r[1]])return;
    var yy=String(r[2]),mm=Number(r[3]);if(mm<num(selFrom)||mm>num(selTo))return;
    annual[yy]=num(annual[yy]||0)+num(r[6]);
  });
  c.annual=annual;
  hist[String(viewYear)]=actualSeries;
  if(ly.total>0||Object.keys(ly.months).length)hist[String(viewYear-1)]=ly.months;
  c.history=hist;
  // Category / SKU splits (its own composition), used by Category & SKU tables.
  c.category={items:objectToItems(act.perCat)};
  var cm=breakdownMaps(name);
  c.monthly_by_cat=cm.cat.cur;c.history_by_cat=cm.cat.hist;c.annual_by_cat=cm.cat.ann;
  c.monthly_by_sku=cm.sku.cur;c.history_by_sku=cm.sku.hist;c.annual_by_sku=cm.sku.ann;
  return c;
}
function objectToItems(o){var it=[];Object.keys(o||{}).forEach(function(k){it.push({name:k,mb:num(o[k]),pct:0});});var t=0;it.forEach(function(x){t+=x.mb;});it.forEach(function(x){x.pct=t?x.mb/t*100:0;});return it;}
// Build the Category and SKU breakdown maps for one Ship-to from facts in ONE pass.
// cur  = {key:{actual:{m:val},ly:{m:val},cat:catName}}  (viewYear / viewYear-1 per month)
// hist = {key:{year:{m:val}}}  (all selected-view years, capped by the visible month range)
// ann  = {key:{year:val}}
function breakdownMaps(name){
  var y1=String(viewYear),y0=String(viewYear-1);
  var catCur={},catHist={},catAnn={},skuCur={},skuHist={},skuAnn={};
  var from=num(selFrom),to=num(selTo),selSet=shipToBasisSet();
  function bucket(c,histc,annc,key,y,m,v){var yo=histc[key];if(!yo)yo=histc[key]={};if(!yo[y])yo[y]={};yo[y][String(m)]=num(yo[y][String(m)]||0)+v;var ac=annc[key];if(!ac)ac=annc[key]={};ac[y]=num(ac[y]||0)+v;var cc=c[key];if(!cc)cc=c[key]={actual:{},ly:{}};var kk=(y===y1)?"actual":(y===y0?"ly":null);if(kk)cc[kk][String(m)]=num(cc[kk][String(m)]||0)+v;}
  (S.facts||[]).forEach(function(r){
    if(r[0]!==name||!selSet[r[1]])return;
    var y=String(r[2]);if(y!==y1&&y!==y0)return;
    var m=Number(r[3]);if(m<from||m>to)return;
    if(activeCat&&r[4]!==activeCat)return;
    if(activeSku&&r[5]!==activeSku)return;
    var v=num(r[6]);
    if(!catCur[r[4]])catCur[r[4]]={actual:{},ly:{},cat:r[4]};bucket(catCur,catHist,catAnn,r[4],y,m,v);
    if(!skuCur[r[5]])skuCur[r[5]]={actual:{},ly:{},cat:r[4]};bucket(skuCur,skuHist,skuAnn,r[5],y,m,v);
  });
  return {cat:{cur:catCur,hist:catHist,ann:catAnn},sku:{cur:skuCur,hist:skuHist,ann:skuAnn}};
}
function shipToLabel(){return activeShipTo?("Ship-to \u2022 "+(activeShipTo||"")):label();}

function load(){
 Promise.all([fetch("data_channels.json?v="+Date.now()).then(r=>r.json()),fetch("shipto_data.json?v="+Date.now()).then(r=>r.json()),fetch("sku_weekly_data.json?v="+Date.now()).then(r=>r.json()).catch(function(){return null;})]).then(function(res){
   D=res[0];S=res[1];SW=res[2];lastClosed=D._lastClosed||8;fy=2026;
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
   // Year filter: list every year present in the data (actual/LY history and annual),
   // default to the current fiscal year (fy) so the first-open view is unchanged.
   var sy=document.getElementById("selYear");sy.innerHTML="";
   var yearSet={};(D._channels||[]).forEach(function(chc){var cc=D[chc.id];if(!cc)return;(Object.keys(cc.history||{}).concat(Object.keys(cc.annual||{}))).forEach(function(y){if(/^\d{4}$/.test(String(y)))yearSet[y]=1;});});
   var years=Object.keys(yearSet).map(Number).sort(function(a,b){return a-b;});
   if(!years.length)years=[fy];
   years.forEach(function(y){var o=document.createElement("option");o.value=y;o.textContent=y;sy.appendChild(o);});
   sy.value=String(fy);viewYear=fy;sy.disabled=false;
   sy.onchange=function(){viewYear=+sy.value;kpiPick=null;renderAll();};
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
// ---- Year filter resolution (viewYear). These fall back to the current-year
// series (monthly / s_actual / s_ly) so the default 2026 view is unchanged. ----
function isHistoricalYear(){return viewYear!==fy;}
function actualSeriesFor(c,y){y=String(y==null?viewYear:y);if(c&&c.history&&c.history[y])return c.history[y];return (c&&c.monthly)||{};}
function lySeriesFor(c,y){y=Number(y==null?viewYear:y);if(c&&c.history&&c.history[String(y-1)])return c.history[String(y-1)];return (y===fy&&c&&c.s_ly)||{};}
function actualFocusSeries(c,name,isSku,y){y=Number(y==null?viewYear:y);var s=String(y),map=isSku?(c.history_by_sku||c.monthly_by_sku):(c.history_by_cat||c.monthly_by_cat);if(map&&map[name]&&map[name][s])return map[name][s];if(map&&map[name]&&map[name].actual)return map[name].actual;return {};}
function lyFocusSeries(c,name,isSku,y){y=Number(y==null?viewYear:y);var s=String(y-1),map=isSku?(c.history_by_sku||c.monthly_by_sku):(c.history_by_cat||c.monthly_by_cat);if(map&&map[name]&&map[name][s])return map[name][s];if(map&&map[name]&&map[name].ly)return map[name].ly;return {};}
function renderAll(){
  try{
    var c=C();var __foc__=currentFocus(c);var __fn=__foc__&&__foc__.name?(" "+String.fromCharCode(8226)+" "+__foc__.name):"";setV("tgY",shipToLabel()+__fn);setV("tgM",shipToLabel()+__fn);setV("tgTbl",shipToLabel()+__fn);setV("tgCat",activeCat?("Category "+String.fromCharCode(8226)+" "+activeCat):shipToLabel());setV("tgSku",shipToLabel()+__fn);
    var ok=hasS(c);
    var rangeTxt=MONTHS[selFrom-1]+"-"+MONTHS[selTo-1];
    var histYear2=isHistoricalYear(),effClosed=histYear2?12:lastClosed,showPlan=false;
    var focKPI=currentFocus(c),focusNameKPI=focKPI&&focKPI.name;
    showPlan=ok&&!focusNameKPI&&!histYear2;
    var viewLbl=histYear2?(rangeTxt+" "+viewYear):rangeTxt;
    setV("lblAct",viewLbl);setV("lblLE",viewLbl);setV("lblLY",viewLbl);setV("lblAOP",viewLbl);
    // KPI actual/ly/le/aop for the active year + channel + range.
    var a=0,leK=null,lyK=0,aopK=null;
    if(focusNameKPI){
      var sr=focKPI.series;for(var mm=selFrom;mm<=selTo;mm++){a+=num(sr.actual&&sr.actual[String(mm)]||0);lyK+=num(sr.ly&&sr.ly[String(mm)]||0);}
    } else if(histYear2){
      var amap=actualSeriesFor(c,viewYear);for(var mm=selFrom;mm<=selTo;mm++)a+=num(amap[String(mm)]);
      lyK=rangeSum(lySeriesFor(c,viewYear),selFrom,selTo);
    } else {
      a=rangeSum(c.s_actual,selFrom,selTo);leK=0;for(var mm2=selFrom;mm2<=selTo;mm2++){if(mm2<lastClosed)leK+=num(c.s_actual[String(mm2)]);else leK+=num(c.s_le[String(mm2)]);}
      lyK=rangeSum(c.s_ly,selFrom,selTo);aopK=rangeSum(c.s_aop,selFrom,selTo);
    }
    setT("kACT",fmt(a)+' <small>MB</small>');setT("kACTuntil",MONTHS[effClosed-1]+" "+viewYear);setT("kACTdate",salesDateTxt());
    // LY card
    var lySeriesKPI=focusNameKPI?focKPI.series.ly:(histYear2?lySeriesFor(c,viewYear):c.s_ly);var lyAvail=hasSeriesData(lySeriesKPI);
    setT("kLY",(lyAvail&&lyK>0)?fmt(lyK)+' <small>MB</small>':"\u2014");
    if(lyAvail&&lyK>0){var lyDiff=a-lyK,lyP=lyK>0?(lyDiff/lyK*100):0,lyCls=lyDiff>=0?"up":"down";setT("kLYgap",(lyDiff>=0?"+":"-")+Math.abs(lyDiff).toFixed(1)+" MB ("+(lyDiff>=0?"+":"-")+Math.abs(lyP).toFixed(1)+"%)");document.getElementById("kLYgap").className=lyCls;}else setT("kLYgap","");
    setT("kLYpc","");setT("kLYcmp",lyAvail?("LY "+String(viewYear-1)):"No LY data");
    // LE card
    if(showPlan&&leK!=null){setT("kLE",fmt(leK)+' <small>MB</small>');var leDiff=a-leK,leP=leK>0?(leDiff/leK*100):0,leCls=leDiff>=0?"up":"down";setT("kGap",(leDiff>=0?"+":"-")+Math.abs(leDiff).toFixed(1)+" MB ("+(leDiff>=0?"+":"-")+Math.abs(leP).toFixed(1)+"%)");document.getElementById("kGap").className=leCls;}
    else{setT("kLE","\u2014");setT("kGap","\u2014");}
    // AOP card
    if(showPlan&&aopK!=null){setT("kAOP",fmt(aopK)+' <small>MB</small>');var aopDiff=leK-aopK,aopP=aopK>0?(aopDiff/aopK*100):0,aopCls=aopDiff>=0?"up":"down";setT("kAOPgap",(aopDiff>=0?"+":"-")+Math.abs(aopDiff).toFixed(1)+" MB ("+(aopDiff>=0?"+":"-")+Math.abs(aopP).toFixed(1)+"%)");document.getElementById("kAOPgap").className=aopCls;setT("kAOPpc","");setT("kAOPcmp","LE vs AOP");}
    else{setT("kAOP","\u2014");setT("kAOPgap","");setT("kAOPpc","");setT("kAOPcmp",histYear2?"No planning for selected year":"No AOP data");}
    setV("hDate",salesDateTxt());
    drawAnnual();drawMonthly();drawTable();
    drawCategory(c);drawSku(c);drawCategoryPerformance(c);drawChannelSummary();drawShipTo();drawSkuWeekly();
    if(kpiPick)applyKpiPick();
  }catch(e){setT("tgM","ERR: "+e.message);}
}
function chgBg(m,sel){return m===sel?"#e11d48":"#f87171";}
function sv(c,key,m){ if(c&&c[key]&&hasSeriesData(c[key])) return num(c[key][String(m)]); if(c&&c.monthly) return num(c.monthly[String(m)]); return 0; }
function currentFocus(c){var foc=null,isSku=!!activeSku;if(activeSku&&(c.monthly_by_sku||c.history_by_sku))foc={name:activeSku,annual:c.annual_by_sku&&c.annual_by_sku[activeSku],history:c.history_by_sku&&c.history_by_sku[activeSku]};else if(activeCat&&(c.monthly_by_cat||c.history_by_cat)){foc={name:activeCat,annual:c.annual_by_cat&&c.annual_by_cat[activeCat],history:c.history_by_cat&&c.history_by_cat[activeCat]};isSku=false;}if(!foc)return null;foc.series={actual:actualFocusSeries(c,foc.name,isSku,viewYear),ly:lyFocusSeries(c,foc.name,isSku,viewYear)};return foc;}
function pickKpi(type,value){if(type==="year"){viewYear=+value;var syD=document.getElementById("selYear");if(syD)syD.value=String(viewYear);if(kpiPick&&kpiPick.type==="month")kpiPick={type:"yearMonth",year:String(value),month:+kpiPick.value};else if(kpiPick&&kpiPick.type==="yearMonth")kpiPick=(String(kpiPick.year)===String(value))?{type:"month",value:kpiPick.month}:{type:"yearMonth",year:String(value),month:kpiPick.month};else{kpiPick=(kpiPick&&kpiPick.type==="year"&&String(kpiPick.value)===String(value))?null:{type:"year",value:String(value)};}}else{if(kpiPick&&kpiPick.type==="year")kpiPick={type:"yearMonth",year:String(kpiPick.value),month:+value};else if(kpiPick&&kpiPick.type==="yearMonth")kpiPick=(+kpiPick.month===+value)?{type:"year",value:String(kpiPick.year)}:{type:"yearMonth",year:String(kpiPick.year),month:+value};else kpiPick=(kpiPick&&kpiPick.type==="month"&&+kpiPick.value===+value)?null:{type:"month",value:+value};}setTimeout(renderAll,0);}
function applyKpiPick(){var c=C(),focus=currentFocus(c);if(kpiPick.type==="year"){var y=String(kpiPick.value),annual=(focus&&focus.annual)||c.annual||{},prev=String(+y-1);setV("lblAct",y);setT("kACT",fmt(annual[y])+' <small>MB</small>');setT("kACTuntil","Full year / YTD");setV("lblLE",y);setT("kLE","\u2014");setT("kGap","\u2014");setV("lblLY",prev);setT("kLY",annual[prev]!=null?fmt(annual[prev])+' <small>MB</small>':"\u2014");setT("kLYcmp","Previous year");setV("lblAOP",y);setT("kAOP","\u2014");setT("kAOPcmp","No annual AOP");return;}
 var m=kpiPick.type==="yearMonth"?+kpiPick.month:+kpiPick.value,y=kpiPick.type==="yearMonth"?+kpiPick.year:fy,history=(focus&&focus.history)||c.history||{},actual=num(history[String(y)]&&history[String(y)][String(m)]),ly=history[String(y-1)]?num(history[String(y-1)][String(m)]):null,current=(y===fy),le=(!focus&&current&&hasS(c))?(m<lastClosed?actual:sv(c,"s_le",m)):null,aop=(!focus&&current&&hasS(c))?sv(c,"s_aop",m):null;["lblAct","lblLE","lblLY","lblAOP"].forEach(function(id){setV(id,MONTHS[m-1]+" "+y);});setT("kACT",fmt(actual)+' <small>MB</small>');setT("kACTuntil",MONTHS[m-1]+" "+y);setT("kLE",le==null?"\u2014":fmt(le)+' <small>MB</small>');setT("kGap",le==null?"\u2014":fmt(Math.max(0,le-actual))+" MB");setT("kLY",ly==null?"\u2014":fmt(ly)+' <small>MB</small>');setT("kLYcmp",ly==null?"No LY data":"Monthly LY "+(y-1));if(aop){var aopDiff=le-aop,aopP=aopDiff/aop*100;setT("kAOP",fmt(aop)+' <small>MB</small>');setT("kAOPgap",(aopDiff>=0?"+":"-")+Math.abs(aopDiff).toFixed(1)+" MB ("+(aopDiff>=0?"+":"-")+Math.abs(aopP).toFixed(1)+"%)");document.getElementById("kAOPgap").className=aopDiff>=0?"up":"down";setT("kAOPpc","");setT("kAOPcmp","Monthly LE vs AOP");}else{setT("kAOP","\u2014");setT("kAOPgap","");setT("kAOPpc","");setT("kAOPcmp","No AOP data");}}

function drawAnnual(){destroy("cAnnual");var c=C(),focus=currentFocus(c),annual=(focus&&focus.annual)||c.annual||{},history=(focus&&focus.history)||c.history||{},years=Object.keys(c.annual||annual).sort(),month=kpiPick&&(kpiPick.type==="month"?+kpiPick.value:(kpiPick.type==="yearMonth"?+kpiPick.month:null)),rangeLabel=MONTHS[selFrom-1]+"\u2013"+MONTHS[selTo-1],vals=years.map(function(y){if(month)return num(history[y]&&history[y][String(month)]);var total=0,hy=history[y]||{};for(var m=selFrom;m<=selTo;m++)total+=num(hy[String(m)]);return total;}),selectedYear=kpiPick&&(kpiPick.type==="year"?String(kpiPick.value):(kpiPick.type==="yearMonth"?String(kpiPick.year):null)),colors=years.map(function(y){return selectedYear===y?"#000000":"#525252";}),base=focus?(shipToLabel()+" \u2022 "+focus.name):shipToLabel();setV("tgY",base+(month?" \u2022 "+MONTHS[month-1]:" \u2022 "+rangeLabel));charts["cAnnual"]=new Chart(document.getElementById("cAnnual"),{type:"bar",data:{labels:years,datasets:[{label:"Actual",data:vals,backgroundColor:colors,hoverBackgroundColor:"#000000",borderRadius:6,barPercentage:.55}]},
  plugins:[valueLabelPlugin],options:{responsive:true,maintainAspectRatio:false,onClick:function(e,els){if(els&&els.length)pickKpi("year",years[els[0].index]);},plugins:{legend:{display:false},tooltip:{callbacks:{label:function(x){return (month?MONTHS[month-1]:rangeLabel)+": "+fmt(x.parsed.y)+" MB";}}}},scales:{y:{beginAtZero:true,grace:"15%"},x:{grid:{display:false}}}}});}

function drawMonthly(){destroy("cMonthly");var c=C();var lb=[],act=[],le=[],ly=[],aop=[];
 var focusInfo=currentFocus(c),focus=focusInfo&&focusInfo.series,focusName=focusInfo&&focusInfo.name||"",vy=viewYear,history=(focusInfo&&focusInfo.history)||c.history||{},historical=(vy!==fy),seriesOk=hasS(c)&&!focus&&!historical;
 setV("tgM",(focusName?(shipToLabel()+" \u2022 "+focusName):shipToLabel())+(historical?" \u2022 "+vy:""));
 ["lgLE","lgAOP"].forEach(function(id){var e=document.getElementById(id);if(e)e.style.display=(focus||historical)?"none":"inline-flex";});
 var actualSeries=(focus?focus.actual:(historical?(history[String(vy)]||{}):actSeries(c))),lySeries=(focus?focus.ly:(historical?(history[String(vy-1)]||{}):comparisonLySeries(c))),lyAvailable=hasSeriesData(lySeries),maxMonth=historical?12:lastClosed,latestVisible=Math.min(selTo,maxMonth);
 for(var i=1;i<=12;i++){var m=i;lb.push((m>=selFrom&&m<=selTo)?MONTHS[m-1]:"");act.push((m<=maxMonth&&m>=selFrom&&m<=selTo)?num(actualSeries[String(m)]):null);le.push((seriesOk&&m>=lastClosed&&m>=selFrom&&m<=selTo)?sv(c,"s_le",m):null);ly.push((lyAvailable&&m>=selFrom&&m<=selTo)?num(lySeries[String(m)]):null);aop.push((seriesOk&&m>=selFrom&&m<=selTo)?sv(c,"s_aop",m):null);}
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
function cmpDoubleCell(title,vs,p){if(vs==null)return '<td>\u2014</td><td>\u2014</td>';var s=signCls(vs),t=signTxt(vs,p);return '<td class="'+s+'"'+exportRawAttr(vs)+'>'+t.v+'</td><td class="'+s+'"'+exportRawAttr(p)+'>'+t.p+'</td>';}
function drawTable(){var c=C(),focusInfo=currentFocus(c),focus=focusInfo&&focusInfo.series,isTotal=!focus,histYear=isHistoricalYear(),effClosed=histYear?12:lastClosed;
 var lyYear=viewYear-1,h='<thead><tr><th>Month</th><th>Actual</th><th>LE</th><th>LY '+lyYear+'</th><th>AOP</th><th>VS LY</th><th>%VS LY</th><th>VS AOP</th><th>%VS AOP</th></tr></thead><tbody>';
 var seriesOk=hasS(c)&&!histYear,lySeries=focus?focus.ly:(histYear?lySeriesFor(c,viewYear):comparisonLySeries(c)),lyAvailable=hasSeriesData(lySeries),ta=0,tl=0,ty=0,ta2=0,tvly=0,tvaop=0;
 for(var m=1;m<=12;m++){if(m<selFrom||m>selTo)continue;
  var act,le,ly,aop;
  if(focus){act=(m<=effClosed?num((focus.actual||{})[String(m)]):null);le=null;ly=lyAvailable?num(lySeries[String(m)]):null;aop=null;}
  else if(histYear){act=(m<=effClosed?num((actualSeriesFor(c,viewYear)||{})[String(m)]):null);le=null;ly=lyAvailable?num(lySeries[String(m)]):null;aop=null;}
  else{act=(m<=lastClosed)?num(actSeries(c)[String(m)]):null;le=seriesOk?((m<lastClosed)?act:sv(c,"s_le",m)):null;ly=lyAvailable?num(lySeries[String(m)]):null;aop=seriesOk?sv(c,"s_aop",m):null;}
  var vly=(act!=null&&ly!=null)?(act-ly):null,ply=(vly!=null&&ly>0)?(vly/ly*100):null,vaop=(act!=null&&aop!=null)?(act-aop):null,paop=(vaop!=null&&aop>0)?(vaop/aop*100):null;
  ta+=act||0;ty+=ly||0;tvly+=vly||0;if(isTotal&&!histYear){tl+=le||0;ta2+=aop||0;tvaop+=vaop||0;}
 h+="<tr><td>"+MONTHS[m-1]+"</td><td"+exportRawAttr(act)+">"+(act==null?"\u2014":fmt(act))+"</td><td"+exportRawAttr(le)+">"+(le==null?"\u2014":fmt(le))+"</td><td"+exportRawAttr(ly)+">"+(ly==null?"\u2014":fmt(ly))+"</td><td"+exportRawAttr(aop)+">"+(aop==null?"\u2014":fmt(aop))+"</td>"+cmpDoubleCell("VS LY",vly,ply)+cmpDoubleCell("VS AOP",vaop,paop)+"</tr>";}
 var ptvly=(ty>0)?(tvly/ty*100):null,ptvaop=(isTotal&&!histYear&&ta2>0)?(tvaop/ta2*100):null;
 h+='<tr class="total"><td>Total</td><td'+exportRawAttr(ta)+'>'+fmt(ta)+'</td><td'+exportRawAttr(isTotal&&!histYear&&seriesOk?tl:null)+'>'+(isTotal&&!histYear&&seriesOk?fmt(tl):'\u2014')+'</td><td'+exportRawAttr(lyAvailable?ty:null)+'>'+(lyAvailable?fmt(ty):'\u2014')+'</td><td'+exportRawAttr(isTotal&&!histYear&&seriesOk?ta2:null)+'>'+(isTotal&&!histYear&&seriesOk?fmt(ta2):'\u2014')+'</td>'+cmpDoubleCell("Total VS LY",lyAvailable?tvly:null,lyAvailable?ptvly:null)+((isTotal&&!histYear&&seriesOk)?cmpDoubleCell("Total VS AOP",tvaop,ptvaop):'<td>\u2014</td><td>\u2014</td>')+'</tr></tbody>';
 setT("mTable",h);}
function drawCategory(c){var el=document.getElementById("cCat");if(!el||!Chart)return;destroy("cCat");var cT=document.getElementById("tgCat");if(cT)cT.textContent=(activeCat?("Category \u2022 "+activeCat):shipToLabel());
 var histYear=isHistoricalYear(),allCats,tot;
 if(histYear){var hm=c.history_by_cat||{},hs=Object.keys(hm).map(function(n){var s=0,h=hm[n]&&hm[n][String(viewYear)]||{};for(var m=selFrom;m<=selTo;m++)s+=num(h[String(m)]);return {name:n,mb:s};}).sort(function(a,b){return b.mb-a.mb;}),grand=0;hs.forEach(function(x){grand+=x.mb;});allCats=hs.map(function(x){return {name:x.name,mb:x.mb,pct:grand?(x.mb/grand*100):0};});tot=grand;}
 else{allCats=(c.category&&c.category.items)?c.category.items:[];tot=c.total_mb;}
 if(!allCats.length){setT("catLegend","No data");return;}
 var selectedIndex=-1;for(var ci=0;ci<allCats.length;ci++)if(allCats[ci].name===activeCat){selectedIndex=ci;break;}
 var C2=selectedIndex>=0?[{name:allCats[selectedIndex].name,mb:allCats[selectedIndex].mb,pct:100}]:allCats;
 var focus=currentFocus(c);if(focus&&focus.series)tot=rangeSum(focus.series.actual,1,(isHistoricalYear()?12:lastClosed));if(selectedIndex>=0)C2[0].mb=tot;
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
// ---- Tri-state sortable headers for performance tables ----
function sortPerfRows(rows,key,dir){
  if(!dir)return rows.slice(); // clear -> default order (rows already default-sorted)
  var miss=function(x){return x==null||isNaN(x);};
  var seq=rows.slice();
  seq.sort(function(a,b){
    var va=a.vals[key],vb=b.vals[key];
    var am=miss(va),bm=miss(vb);
    if(am&&bm)return a.o-b.o;
    if(am)return 1;      // missing always last
    if(bm)return -1;
    if(va<vb)return dir==='desc'?1:-1;
    if(va>vb)return dir==='desc'?-1:1;
    return a.o-b.o;      // tie -> stable default order
  });
  return seq;
}
function perfHeaderCell(label,col,active){
  var d=active?active.dir:null;
  var aria=d==='desc'?'descending':d==='asc'?'ascending':'none';
  var ic=d==='desc'?' \u25bc':d==='asc'?' \u25b2':'';
  var cls='sort-btn'+(d?' active':'');
  var sr=(active?'<span class="sr-only">, sorted '+ (d==='desc'?'descending':'ascending')+'</span>':'');
  return '<th scope="col" aria-sort="'+aria+'"><button type="button" class="'+cls+'" data-colsort="'+col+'" aria-label="Sort by '+label+(d?', currently '+(d==='desc'?'descending':'ascending'):'')+'">'+label+'<span class="sort-ind">'+ic+'</span>'+sr+'</button></th>';
}
function nextSortState(cur,key){if(!cur||cur.col!==key)return {col:key,dir:'desc'};if(cur.dir==='desc')return {col:key,dir:'asc'};return null;}
function bindSortHeaders(scopeId){
  var tbl=document.getElementById(scopeId);if(!tbl)return;
  var isCat=scopeId==='catPerfTable';
  Array.prototype.forEach.call(tbl.querySelectorAll('button[data-colsort]'),function(btn){
    var col=btn.getAttribute('data-colsort');
    function doSort(){var ns=nextSortState(isCat?catSort:skuSort,col);if(isCat)catSort=ns;else skuSort=ns;renderAll();}
    btn.addEventListener('click',doSort);
    btn.addEventListener('keydown',function(e){if(e.key==='Enter'||e.key===' '){e.preventDefault();doSort();}});
  });
}
function drawCategoryPerformance(c){
 var histYear=isHistoricalYear(),map,names,total;
 if(histYear){map={};names=Object.keys(c.history_by_cat||{});total=rangeSum(actualSeriesFor(c,viewYear),selFrom,selTo);names.forEach(function(name){map[name]={actual:(c.history_by_cat[name]&&c.history_by_cat[name][String(viewYear)])||{},ly:(c.history_by_cat[name]&&c.history_by_cat[name][String(viewYear-1)])||{}};});}
 else{map=c.monthly_by_cat||{};names=Object.keys(map);total=rangeSum(actSeries(c),selFrom,selTo);}
 names=names.slice().sort(function(a,b){return perfValues(map[b]).actual-perfValues(map[a]).actual;});
 var rows=[];names.forEach(function(name){rows.push({name:name,v:perfValues(map[name])});});
 rows.forEach(function(r,i){var v=r.v;r.vals={ly:v.ly,actual:v.actual,growth:v.growth,pct:v.pct,contrib:total?v.actual/total*100:0};r.o=i;});
 var st=catSort,ordered=sortPerfRows(rows,st?st.col:null,st?st.dir:null);
 var h='<thead><tr><th style="text-align:left">Category</th>'+perfHeaderCell('Actual '+(viewYear-1),'ly',st)+perfHeaderCell('Actual '+viewYear,'actual',st)+perfHeaderCell('Growth vs LY (MB)','growth',st)+perfHeaderCell('% Growth vs LY','pct',st)+perfHeaderCell('% Contribution','contrib',st)+'</tr></thead><tbody>';
 if(!ordered.length)h+='<tr><td>No Category</td><td colspan="5">\u2014</td></tr>';
 var ta=0,tl=0,allLy=true;
 ordered.forEach(function(r){var v=r.v,con=r.vals.contrib,cls='perf-row'+(v.growth!=null&&v.growth>0?' growing':'')+(activeCat===r.name?' selected':'');ta+=v.actual;if(v.ly==null)allLy=false;else tl+=v.ly;h+='<tr class="'+cls+'" data-cat="'+encodeURIComponent(r.name)+'"><td style="text-align:left;white-space:normal">'+r.name+'</td><td'+exportRawAttr(v.ly)+'>'+perfCell(v.ly)+'</td><td'+exportRawAttr(v.actual)+'>'+fmt(v.actual)+'</td><td'+exportRawAttr(v.growth)+'>'+signedPerf(v.growth)+'</td><td'+exportRawAttr(v.pct)+'>'+signedPerf(v.pct,'%')+'</td><td'+exportRawAttr(con)+'>'+fmt(con)+'%</td></tr>';});
 var tg=allLy?ta-tl:null,tp=(allLy&&tl!==0)?tg/tl*100:null;
 h+='</tbody><tfoot><tr class="total"><td style="text-align:left">Total</td><td'+exportRawAttr(allLy?tl:null)+'>'+perfCell(allLy?tl:null)+'</td><td'+exportRawAttr(ta)+'>'+fmt(ta)+'</td><td'+exportRawAttr(tg)+'>'+signedPerf(tg)+'</td><td'+exportRawAttr(tp)+'>'+signedPerf(tp,'%')+'</td><td'+exportRawAttr(ordered.length?100:null)+'>'+(ordered.length?'100.0%':'\u2014')+'</td></tr></tfoot>';setT('catPerfTable',h);setV('tgCatPerf',shipToLabel()+(activeCat?' \u2022 '+activeCat:''));
 Array.prototype.forEach.call(document.querySelectorAll('#catPerfTable .perf-row'),function(r){r.onclick=function(){toggleCat(decodeURIComponent(r.getAttribute('data-cat')));};});
 bindSortHeaders('catPerfTable');
}
function skuAuthoritative(c){var histYear=isHistoricalYear(),series,actual,ly,hasLy;
 if(activeSku){var hsSku=c.history_by_sku&&c.history_by_sku[activeSku];if(histYear)series=hsSku?{actual:hsSku[String(viewYear)]||{},ly:hsSku[String(viewYear-1)]||{}}:null;else series=c.monthly_by_sku&&c.monthly_by_sku[activeSku];}
 else if(activeCat){var hsCat=c.history_by_cat&&c.history_by_cat[activeCat];if(histYear)series=hsCat?{actual:hsCat[String(viewYear)]||{},ly:hsCat[String(viewYear-1)]||{}}:null;else series=c.monthly_by_cat&&c.monthly_by_cat[activeCat];}
 if(series){actual=rangeSum(series.actual,selFrom,selTo);hasLy=hasSeriesData(series.ly);ly=hasLy?rangeSum(series.ly,selFrom,selTo):null;}
 else{actual=rangeSum(actualSeriesFor(c,viewYear),selFrom,selTo);var lys=lySeriesFor(c,viewYear);hasLy=hasSeriesData(lys);ly=hasLy?rangeSum(lys,selFrom,selTo):null;}
 return {actual:actual,ly:ly};}
function drawSku(c){var histYear=isHistoricalYear(),map;
 if(histYear){map={};Object.keys(c.history_by_sku||{}).forEach(function(n){var ys=c.history_by_sku[n]||{},cat=(c.monthly_by_sku&&c.monthly_by_sku[n]&&c.monthly_by_sku[n].cat)||"";map[n]={cat:cat,actual:ys[String(viewYear)]||{},ly:ys[String(viewYear-1)]||{}};});}
 else map=c.monthly_by_sku||{};
 var auth=skuAuthoritative(c),total=auth.actual;var list=Object.keys(map).filter(function(name){var s=map[name]||{};if(activeSku)return name===activeSku;if(activeCat)return s.cat===activeCat;return true;}).map(function(name){return {name:name,cat:(map[name]&&map[name].cat)||activeCat||""};});
 if(!list.length){setT("skuTable",'<thead><tr><th>SKU</th><th>Category</th><th>Actual '+(viewYear-1)+'</th><th>Actual '+viewYear+'</th><th>Growth vs LY (MB)</th><th>% Growth vs LY</th><th>% Contribution</th></tr></thead><tbody><tr><td>No SKU</td><td colspan="6">\u2014</td></tr></tbody>');return;}
 var h='<thead><tr><th>#</th><th style="text-align:left">SKU</th><th>Category</th>'+perfHeaderCell('Actual '+(viewYear-1),'ly',skuSort)+perfHeaderCell('Actual '+viewYear,'actual',skuSort)+perfHeaderCell('Growth vs LY (MB)','growth',skuSort)+perfHeaderCell('% Growth vs LY','pct',skuSort)+perfHeaderCell('% Contribution','contrib',skuSort)+'</tr></thead><tbody>';
 var shown=list.slice(),ta=0,tl=0;shown.sort(function(a,b){return perfValues(map[b.name]).actual-perfValues(map[a.name]).actual;});var rows=[];shown.forEach(function(s,i){rows.push({name:s.name,cat:s.cat,o:i,v:perfValues(map[s.name])});});rows.forEach(function(r){var v=r.v;r.vals={ly:v.ly,actual:v.actual,growth:v.growth,pct:v.pct,contrib:total?v.actual/total*100:0};});var st=skuSort,ordered=sortPerfRows(rows,st?st.col:null,st?st.dir:null);ordered.forEach(function(s,i){var cat=s.cat||activeCat||"",v=s.v,p=s.vals.contrib,cls='skuro perf-row'+(v.growth!=null&&v.growth>0?' growing':'')+(activeSku===s.name?' selected':'');ta+=v.actual;tl+=v.ly==null?0:v.ly;h+='<tr class="'+cls+'" data-sku="'+encodeURIComponent(s.name)+'" data-cat="'+encodeURIComponent(cat||'')+'"><td>'+(i+1)+'</td><td style="max-width:260px;white-space:normal;text-align:left">'+s.name+'</td><td>'+cat+'</td><td>'+perfCell(v.ly)+'</td><td>'+fmt(v.actual)+'</td><td>'+signedPerf(v.growth)+'</td><td>'+signedPerf(v.pct,'%')+'</td><td>'+fmt(p)+'%</td></tr>';});
 var residualA=auth.actual-ta,residualLy=auth.ly==null?null:auth.ly-tl;if(Math.abs(residualA)>.005||(residualLy!=null&&Math.abs(residualLy)>.005)){var rg=residualLy==null?null:residualA-residualLy,rp=(residualLy&&rg!=null)?rg/residualLy*100:null;h+='<tr class="recon-row"><td></td><td style="text-align:left;white-space:normal">Unallocated / No SKU detail</td><td></td><td>'+perfCell(residualLy)+'</td><td>'+fmt(residualA)+'</td><td>'+signedPerf(rg)+'</td><td>'+signedPerf(rp,'%')+'</td><td>'+fmt(total?residualA/total*100:0)+'%</td></tr>';}
 var tg=auth.ly==null?null:auth.actual-auth.ly,tp=(auth.ly&&tg!=null)?tg/auth.ly*100:null;
 h+='</tbody><tfoot><tr class="total"><td></td><td style="text-align:left">Total</td><td></td><td>'+perfCell(auth.ly)+'</td><td>'+fmt(auth.actual)+'</td><td>'+signedPerf(tg)+'</td><td>'+signedPerf(tp,'%')+'</td><td>100.0%</td></tr></tfoot>';setT("skuTable",h);
 var srows=document.querySelectorAll(".skuro");for(var i=0;i<srows.length;i++){(function(r){r.onclick=function(){var n=r.getAttribute('data-sku'),cat=r.getAttribute('data-cat');if(n)toggleSku(decodeURIComponent(n),decodeURIComponent(cat||'')||activeCat||"");};})(srows[i]);}
 bindSortHeaders('skuTable');}
function chGrowthYTD(d,x){
  // Compare current (or selected-year) actual with the same closed-month range in the prior year.
  var histYear=isHistoricalYear(),eff=histYear?12:lastClosed;
  var a=0,h=d&&d.history&&d.history[String(viewYear)];if(h){for(var m=1;m<=eff;m++)a+=num(h[String(m)]);}if(!a)a=num(d&&d.total_mb);
  var ly=0,hp=d&&d.history&&d.history[String(viewYear-1)];if(hp){for(var m=1;m<=eff;m++)ly+=num(hp[String(m)]);}
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
  var a=0,hc=(d&&d.history&&d.history[String(viewYear)]);if(hc){for(var m=f;m<=t;m++)a+=num(hc[String(m)]);}if(!a)a=num(d&&d.total_mb);
  var ly=0,h=(d&&d.history&&d.history[String(viewYear-1)]);if(h){for(var m=f;m<=t;m++)ly+=num(h[String(m)]);}
  var g=(ly>0)?((a-ly)/ly*100):null;
  return {a:a,ly:ly,g:g};
}
function drawChannelSummary(){
  var el=document.getElementById("chTable");if(!el)return;
  var SD=sourceData();
  setV("tgChRange",MONTHS[selFrom-1]+"\u2013"+MONTHS[selTo-1]+" "+viewYear);
  // [[SHIPTO-SUMMARY]] When a Ship-to is selected the cards show that location\'s
  // split across the selected channels, so "Sales by Channel" also follows the filter.
  if(activeShipTo){
    var byC={},byLy={},stTot=0;
    (S.facts||[]).forEach(function(r){
      if(r[0]!==activeShipTo||!shipToBasisSet()[r[1]])return;
      var m=Number(r[3]);if(m<num(selFrom)||m>num(selTo))return;
      if(activeCat&&r[4]!==activeCat)return;
      if(activeSku&&r[5]!==activeSku)return;
      if(String(r[2])===String(viewYear)){byC[r[1]]=num(byC[r[1]]||0)+num(r[6]);stTot+=num(r[6]);}
      else if(String(r[2])===String(viewYear-1))byLy[r[1]]=num(byLy[r[1]]||0)+num(r[6]);
    });
    var defs={};(D._channels||[]).forEach(function(ch){defs[ch.id]=ch;});
    var rows2=[],tot2=0;
    Object.keys(byC).forEach(function(cid){
      var def=defs[cid]||{label:cid},ly=num(byLy[cid]||0),a=num(byC[cid]||0),g=ly>0?(a-ly)/ly*100:(a>0?null:null);
      rows2.push({id:cid,label:(cid==="LOTUS'"?"Lotus's":(def.label||cid)),a:a,ly:ly,g:g});tot2+=a;
    });
    rows2.sort(function(x,y){return y.a-x.a;});
    channelExportRows=rows2.map(function(r){var contribution=tot2?r.a/tot2*100:0;return {label:r.label,actual:r.a,contribution:contribution,ly:r.ly,gap:r.a-r.ly,growth:r.g};});
    var h2='<div class="chgrid">';
    if(!rows2.length)h2+='<div class="chcard" style="cursor:default"><div class="chname">No channel data</div><div class="chval">\u2014</div></div>';
    rows2.forEach(function(r){
      var sh=tot2?(r.a/tot2*100):0,grow=r.g,cls=(grow!=null&&grow>=0)?"up":"down",gap=r.a-r.ly;
      var gtxt=(grow==null)?"\u2014":((gap>=0?"+":"")+gap.toFixed(1)+" MB ("+(grow>=0?"+":"")+grow.toFixed(1)+"%)");
      var bg=grow!=null?(grow>=0?"#dcfce7":"#fff"):"#fff";
      h2+='<div class="chcard" data-ch="'+r.id+'" style="background:'+bg+';cursor:pointer">';
      h2+='  <div class="chhead"><span class="chname">'+r.label+'</span></div>';
      h2+='  <div class="chcon">'+sh.toFixed(1)+'%<small>Con</small></div>';
      h2+='  <div class="chval">'+fmt(r.a)+'<small>MB</small></div>';
      h2+='  <div class="chly">LY '+fmt(r.ly)+' Gap <span class="'+cls+'">'+gtxt+'</span></div>';
      h2+='</div>';
    });
    h2+='</div>';
    el.innerHTML=h2;
    Array.prototype.forEach.call(document.querySelectorAll(".chcard[data-ch]"),function(el3){el3.addEventListener("click",function(){selectChan(el3.getAttribute("data-ch"));});});
    return;
  }
  var rows=[],tot=0;
  var focInfo=currentFocus(C()),focusName=focInfo&&focInfo.name;
  // Cards follow the active channel filter so their Actual total reconciles to the KPI.
  var channelDefs={};(D._channels||[]).forEach(function(ch){channelDefs[ch.id]=ch;});
  (selChans||[]).forEach(function(channelId){
    var ch=channelDefs[channelId]||{id:channelId,label:channelId};
    if(ch.id==="MT")return;
    var d=SD[ch.id];if(!d)return;
    var g;
    if(focusName){
      var cm=activeSku?(d.history_by_sku&&d.history_by_sku[activeSku]):(d.history_by_cat&&d.history_by_cat[activeCat]);
      var a=0,ly=0;if(cm){var ca=cm[String(viewYear)]||{},cl=cm[String(viewYear-1)]||{};for(var m=num(selFrom);m<=num(selTo);m++){a+=num(ca[String(m)]||0);ly+=num(cl[String(m)]||0);}}
      if(!a&&!cm){var m2=activeSku?(d.monthly_by_sku&&d.monthly_by_sku[activeSku]):(d.monthly_by_cat&&d.monthly_by_cat[activeCat]);if(m2){for(var m=num(selFrom);m<=num(selTo);m++){a+=num(m2.actual&&m2.actual[String(m)]||0);ly+=num(m2.ly&&m2.ly[String(m)]||0);}}}
      g={a:a,ly:ly,g:ly>0?(a-ly)/ly*100:(a>0?null:null)};
    } else { g=chanGrowth(d); }
    rows.push({id:ch.id,label:(ch.id==="LOTUS'"?"Lotus's":(d.label||ch.label)),a:g.a,ly:g.ly,g:g.g});
    tot+=g.a;
  });
  var actCur=cur,actSel=(selChans&&selChans.length===1)?selChans[0]:null;rows.sort(function(x,y){return y.a-x.a;});
  channelExportRows=rows.map(function(r){var contribution=tot?r.a/tot*100:0;return {label:r.label,actual:r.a,contribution:contribution,ly:r.ly,gap:r.a-r.ly,growth:r.g};});
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
// Sales by Ship-to party table (rows = ship-to, columns = source WK numbers
// inside the selected month range). Unit selector is local to this table.
// Driven by shipto_data.json facts and responds to the same Channel / Year /
// From-To / Category / SKU filters as the KPI Actual, so the row totals
// reconcile with the Actual KPI under identical filters.
function drawShipTo(){
  var el=document.getElementById("shipToTable");if(!el||!S)return;
  var selSet=shipToBasisSet(),from=num(selFrom),to=num(selTo),year=String(viewYear);
  var filterCat=activeCat,filterSku=activeSku;
  var unitMeta={ea:{label:"ชิ้น",idx:8},ton:{label:"ตัน",idx:9},baht:{label:"บาท",idx:10},mb:{label:"ล้านบาท",idx:6}};
  var unit=unitMeta[shipToUnit]||unitMeta.mb;
  var rows={},weekTotals={},weekSet={},grand=0;
  (S.facts||[]).forEach(function(r){
    if(!selSet[r[1]])return;             // channel filter
    if(String(r[2])!==year)return;       // year filter
    var m=Number(r[3]);if(m<from||m>to)return;
    if(filterCat&&r[4]!==filterCat)return; // category filter
    if(filterSku&&r[5]!==filterSku)return;  // SKU filter
    var week=Number(r[11]);if(week<1||week>53)return;
    var nm=r[0],v=num(r[unit.idx]);
    var row=rows[nm];if(!row){row=rows[nm]={name:nm,weeks:{},total:0};}
    row.weeks[week]=num(row.weeks[week])+v;
    row.total=num(row.total)+v;
    weekTotals[week]=num(weekTotals[week])+v;weekSet[week]=1;
    grand=num(grand)+v;
  });
  var weeks=Object.keys(weekSet).map(Number).sort(function(a,b){return a-b;});
  var names=Object.keys(rows).sort(function(a,b){return rows[b].total-rows[a].total;});
  var h='<thead><tr><th>Ship-to party ('+names.length+')</th>';
  weeks.forEach(function(w){h+='<th>WK '+w+'</th>';});
  h+='<th>Total</th></tr></thead><tbody>';
  var show=(names.length<=400);
  if(!show)names=names.slice(0,400); // cap rendered rows; footer Total still sums all
  names.forEach(function(nm){
    var row=rows[nm];
    var on=activeShipTo===nm;
    var enc=encodeURIComponent(nm);
    var cls=['st-row',on?'shipto-selected':''].join(' ');
    h+='<tr class="'+cls+'" tabindex="0" role="button" data-shipto="'+enc+'" aria-pressed="'+on+'" title="'+escAttr((on?"":"ดูเฉพาะ ")+(on?"คลิกอีกครั้งเพื่อแสดงยอดรวมทั้งหมด":nm))+'" style="cursor:pointer">';
    h+='<td style="text-align:left;white-space:normal">'+escHtml(row.name)+(on?' <span class="st-show">\u25bc</span>':'')+'</td>';
    weeks.forEach(function(w){h+='<td'+exportRawAttr(row.weeks[w]||0)+'>'+shipToFmt(row.weeks[w]||0)+'</td>';});
    h+='<td'+exportRawAttr(row.total)+'><b>'+shipToFmt(row.total)+'</b></td></tr>';
  });
  if(!show)h+='<tr><td style="text-align:left" class="recon-row">Showing top 400 by sales; remaining ship-to parties included in Total.</td><td colspan="'+(weeks.length+1)+'"></td></tr>';
  h+='</tbody><tfoot><tr class="total"><td style="text-align:left">Total</td>';
  weeks.forEach(function(w){h+='<td'+exportRawAttr(weekTotals[w]||0)+'>'+shipToFmt(weekTotals[w]||0)+'</td>';});
  h+='<td'+exportRawAttr(grand)+'>'+shipToFmt(grand)+'</td></tr></tfoot>';
  setT("shipToTable",h);
  // Click a row to filter everything to that Ship-to; click again to clear.
  Array.prototype.forEach.call(document.querySelectorAll("#shipToTable tbody tr.st-row"),function(tr){
    function pick(){toggleShipTo(decodeURIComponent(tr.getAttribute("data-shipto")));}
    tr.addEventListener("click",pick);
    tr.addEventListener("keydown",function(e){if(e.key==="Enter"||e.key===" "){e.preventDefault();pick();}});
  });
  var rangeLabel=MONTHS[from-1]+"–"+MONTHS[to-1]+" "+year;
  setV("tgShipTo",(activeShipTo?("Ship-to \u2022 "+activeShipTo):(rangeLabel+" | Total "+shipToFmt(grand)+" "+unit.label)));
  var header=document.getElementById("shiptoHeaderActions");
  if(header){header.innerHTML=activeShipTo?(''+
    '<button type="button" class="export-btn st-clear" id="stClear" style="border-color:#e11d48;background:#fff1f2;color:#be123c">\u2715 \u0e40\u0e04\u0e25\u0e35\u0e22\u0e23\u0e4c\u0e01\u0e23\u0e2d\u0e07\u0e2a\u0e16\u0e32\u0e19\u0e17\u0e35\u0e48</button>'):'';
    var clr=header.querySelector("#stClear");if(clr)clr.onclick=function(){toggleShipTo(activeShipTo);};}
}
function drawSkuWeekly(){
  var el=document.getElementById("skuWeeklyTable");if(!el)return;
  if(!SW||!SW.facts){setT("skuWeeklyTable","");setV("tgSkuWeekly","No data");return;}
  var selSet=shipToBasisSet(),from=num(selFrom),to=num(selTo),year=String(viewYear);
  var filterCat=activeCat,filterSku=activeSku,search=String(skuWeeklySearch||"").trim().toLowerCase();
  var unitMeta={ea:{label:"ชิ้น",idx:6},ton:{label:"ตัน",idx:7},baht:{label:"บาท",idx:8},mb:{label:"ล้านบาท",idx:9}};
  var unit=unitMeta[skuWeeklyUnit]||unitMeta.mb;
  var rows={},weekTotals={},weekSet={},grand=0;
  (SW.facts||[]).forEach(function(r){
    if(!selSet[r[2]])return;
    if(String(r[3])!==year)return;
    var m=Number(r[4]);if(m<from||m>to)return;
    if(filterCat&&r[5]!==filterCat)return;
    if(filterSku&&r[1]!==filterSku)return;
    var mat=String(r[0]||""),desc=String(r[1]||"");
    if(search&&mat.toLowerCase().indexOf(search)<0&&desc.toLowerCase().indexOf(search)<0)return;
    var week=Number(r[10]);if(week<1||week>53)return;
    var key=mat+"|"+desc,v=num(r[unit.idx]);
    var row=rows[key];if(!row){row=rows[key]={material:mat,name:desc,weeks:{},total:0};}
    row.weeks[week]=num(row.weeks[week])+v;
    row.total=num(row.total)+v;
    weekTotals[week]=num(weekTotals[week])+v;weekSet[week]=1;
    grand=num(grand)+v;
  });
  var weeks=Object.keys(weekSet).map(Number).sort(function(a,b){return a-b;});
  var keys=Object.keys(rows).sort(function(a,b){return rows[b].total-rows[a].total;});
  var h='<thead><tr><th>Material / SKU ('+keys.length+')</th>';
  weeks.forEach(function(w){h+='<th>WK '+w+'</th>';});
  h+='<th>Total</th></tr></thead><tbody>';
  var show=(keys.length<=400||search);
  if(!show)keys=keys.slice(0,400);
  keys.forEach(function(k){
    var row=rows[k],label=(row.material?row.material+" ":"")+row.name;
    h+='<tr data-material="'+escAttr(row.material)+'" data-name="'+escAttr(row.name)+'"><td style="text-align:left;white-space:normal">'+escHtml(label)+'</td>';
    weeks.forEach(function(w){h+='<td'+exportRawAttr(row.weeks[w]||0)+'>'+weeklyFmt(row.weeks[w]||0,skuWeeklyUnit)+'</td>';});
    h+='<td'+exportRawAttr(row.total)+'><b>'+weeklyFmt(row.total,skuWeeklyUnit)+'</b></td></tr>';
  });
  if(!show)h+='<tr><td style="text-align:left" class="recon-row">Showing top 400 by sales; remaining SKUs included in Total. Use Material search to narrow.</td><td colspan="'+(weeks.length+1)+'"></td></tr>';
  if(!keys.length)h+='<tr><td style="text-align:left" class="recon-row">No SKU found for the current filters.</td><td colspan="'+(weeks.length+1)+'"></td></tr>';
  h+='</tbody><tfoot><tr class="total"><td style="text-align:left">Total</td>';
  weeks.forEach(function(w){h+='<td'+exportRawAttr(weekTotals[w]||0)+'>'+weeklyFmt(weekTotals[w]||0,skuWeeklyUnit)+'</td>';});
  h+='<td'+exportRawAttr(grand)+'>'+weeklyFmt(grand,skuWeeklyUnit)+'</td></tr></tfoot>';
  setT("skuWeeklyTable",h);
  syncSkuWeeklyScroll();
  var rangeLabel=MONTHS[from-1]+"–"+MONTHS[to-1]+" "+year;
  setV("tgSkuWeekly",rangeLabel+" | Total "+weeklyFmt(grand,skuWeeklyUnit)+" "+unit.label+(search?" | Search "+skuWeeklySearch:""));
}
function shipToFmt(v){return Number(v||0).toLocaleString("en-US",{minimumFractionDigits:shipToUnit==="ea"?0:2,maximumFractionDigits:shipToUnit==="ea"?0:2});}
function weeklyFmt(v,unitKey){return Number(v||0).toLocaleString("en-US",{minimumFractionDigits:unitKey==="ea"?0:2,maximumFractionDigits:unitKey==="ea"?0:2});}
function setShipToUnit(v){if(!({ea:1,ton:1,baht:1,mb:1})[v])return;shipToUnit=v;drawShipTo();}
function setSkuWeeklyUnit(v){if(!({ea:1,ton:1,baht:1,mb:1})[v])return;skuWeeklyUnit=v;drawSkuWeekly();}
function setSkuWeeklySearch(v){skuWeeklySearch=v||"";drawSkuWeekly();}
function syncSkuWeeklyScroll(){
  var viewport=document.getElementById("skuWeeklyViewport"),bar=document.getElementById("skuWeeklyBottomScroll"),inner=document.getElementById("skuWeeklyBottomScrollInner"),table=document.getElementById("skuWeeklyTable");
  if(!viewport||!bar||!inner||!table)return;
  inner.style.width=Math.max(table.scrollWidth,viewport.clientWidth)+"px";
  bar.style.display=table.scrollWidth>viewport.clientWidth?"block":"none";
  bar.scrollLeft=viewport.scrollLeft;
  if(!bar._skuBound){
    bar.addEventListener("scroll",function(){if(!bar._syncing){viewport._syncing=1;viewport.scrollLeft=bar.scrollLeft;viewport._syncing=0;}});
    viewport.addEventListener("scroll",function(){if(!viewport._syncing){bar._syncing=1;bar.scrollLeft=viewport.scrollLeft;bar._syncing=0;}});
    bar._skuBound=1;
  }
}
window.addEventListener("resize",syncSkuWeeklyScroll);
var skuWeeklyExportRows=[];
// Build the Export rows for the SKU weekly card: split Material / SKU name into
// two separate columns (page display keeps them combined in one cell; only the
// exported .xlsx separates them). Re-reads the current on-screen rows via the
// per-row data-material / data-name attributes so it stays in sync with the
// visible filters/unit.
function buildSkuWeeklyExportRows(){
  skuWeeklyExportRows=[];
  var t=document.getElementById("skuWeeklyTable");
  if(!t)return;
  var body=t.querySelector("tbody");
  var weeks=[];
  var hdrRow=t.querySelector("thead tr");
  if(hdrRow){Array.prototype.forEach.call(hdrRow.querySelectorAll("th"),function(th,i){
    if(i===0)return; weeks.push((th.textContent||"").trim());
  });}
  var header=["Material","SKU"].concat(weeks);
  skuWeeklyExportRows.push(header);
  if(!body)return;
  Array.prototype.forEach.call(body.querySelectorAll("tr"),function(tr){
    var mat=tr.getAttribute("data-material")||"";
    var name=tr.getAttribute("data-name")||"";
    // rebuild numeric cells from the visible td values in order (skip the
    // combined Material / SKU first cell)
    var vals=[];
    Array.prototype.forEach.call(tr.querySelectorAll("td"),function(td,i){ if(i>0) vals.push(exportCell((td.textContent||"").trim(),td.getAttribute("data-export-value"))); });
    var row=[mat,name].concat(vals);
    skuWeeklyExportRows.push(row);
  });
}
function toggleShipTo(name){activeShipTo=(activeShipTo===name)?null:name;kpiPick=null;renderAll();}
function escHtml(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function escAttr(s){return escHtml(String(s==null?"":s).replace(/'/g,"&#39;"));}

// Dashboard bootstraps after the password gate unlocks it (see index.html).
