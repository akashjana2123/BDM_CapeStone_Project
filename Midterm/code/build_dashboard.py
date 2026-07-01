import json
data=open('dash_data.json').read()

HTML = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Ankita & Adrita Cashew — Mid-term Dashboard</title>
<style>
  :root{--navy:#1f4e79;--orange:#c55a11;--green:#548235;--bg:#f4f6fa;--card:#fff;--ink:#222;--muted:#6b7280;--line:#e3e8ef;}
  *{box-sizing:border-box}
  body{margin:0;font-family:Arial,Helvetica,sans-serif;background:var(--bg);color:var(--ink);font-size:14px}
  header{background:var(--navy);color:#fff;padding:16px 22px}
  header h1{margin:0;font-size:19px}
  header p{margin:4px 0 0;font-size:12.5px;opacity:.85}
  .wrap{max-width:1180px;margin:0 auto;padding:18px}
  .grid{display:grid;gap:14px}
  .kpis{grid-template-columns:repeat(6,1fr)}
  @media(max-width:900px){.kpis{grid-template-columns:repeat(3,1fr)}}
  @media(max-width:560px){.kpis{grid-template-columns:repeat(2,1fr)}}
  .card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px;box-shadow:0 1px 2px rgba(0,0,0,.03)}
  .kpi .label{font-size:11.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em}
  .kpi .val{font-size:22px;font-weight:700;color:var(--navy);margin-top:4px}
  .kpi .sub{font-size:11.5px;color:var(--muted);margin-top:2px}
  h2{font-size:15px;color:var(--navy);margin:22px 2px 6px;border-bottom:2px solid var(--line);padding-bottom:5px}
  h3{font-size:13px;color:var(--orange);margin:0 0 8px}
  .two{grid-template-columns:1fr 1fr}
  .three{grid-template-columns:1.3fr 1fr 1fr}
  @media(max-width:860px){.two,.three{grid-template-columns:1fr}}
  table{width:100%;border-collapse:collapse;font-size:12.5px}
  th,td{padding:6px 8px;border-bottom:1px solid var(--line);text-align:right}
  th:first-child,td:first-child{text-align:left}
  thead th{background:var(--navy);color:#fff;font-weight:600;border:none}
  tbody tr:nth-child(even){background:#f7f9fc}
  .num{font-variant-numeric:tabular-nums}
  svg{width:100%;height:auto;display:block}
  form{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;align-items:end}
  @media(max-width:860px){form{grid-template-columns:repeat(2,1fr)}}
  label{display:block;font-size:11.5px;color:var(--muted);margin-bottom:3px}
  input,select{width:100%;padding:7px 8px;border:1px solid #cdd5e0;border-radius:6px;font-size:13px;background:#fff}
  .btn{background:var(--navy);color:#fff;border:none;padding:9px 14px;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600}
  .btn:hover{background:#163a5a}
  .btn.alt{background:#fff;color:var(--navy);border:1px solid var(--navy)}
  .btn.alt:hover{background:#eef3f9}
  .btn.green{background:var(--green)}.btn.green:hover{background:#3f641f}
  .row-actions{display:flex;gap:8px;margin-top:10px;flex-wrap:wrap;align-items:center}
  .toolbar{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-top:14px}
  .pill{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600}
  .pill.A{background:#fde2cf;color:#a3430b}.pill.B{background:#d7e3f5;color:#1f4e79}.pill.C{background:#e7e9ee;color:#555}
  .pill.sale{background:#dcecd4;color:#3c5e22}.pill.pur{background:#fde2cf;color:#a3430b}
  .pill.Paid{background:#dcecd4;color:#3c5e22}.pill.Pending{background:#fde0db;color:#a3290b}.pill.Partial{background:#fdf1cf;color:#8a6300}
  .toast{position:fixed;bottom:18px;left:50%;transform:translateX(-50%);background:#16331f;color:#fff;padding:10px 16px;border-radius:8px;font-size:13px;opacity:0;transition:.3s;pointer-events:none;z-index:5}
  .toast.show{opacity:1}
  .added-tbl td{text-align:left}
  .x{color:#c0392b;cursor:pointer;font-weight:700}
  .legend{display:flex;gap:14px;flex-wrap:wrap;font-size:11.5px;color:var(--muted);margin-top:6px}
  .legend span{display:inline-flex;align-items:center;gap:5px}
  .dot{width:10px;height:10px;border-radius:2px;display:inline-block}
  .note{font-size:11.5px;color:var(--muted);margin-top:6px}
  .flagYes{color:#c0392b;font-weight:700}
</style>
</head>
<body>
<header>
  <h1>Ankita &amp; Adrita Cashew Processing — Mid-term Dashboard</h1>
  <p>BDM Capstone · Akash Jana · 23f2000990 · data 07-Apr-2022 to 10-Feb-2026 · values recompute when you add a record</p>
</header>
<div class="wrap">

  <!-- ADD RECORD -->
  <div class="card">
    <h3>Add a new sales / purchase record</h3>
    <form id="frm" autocomplete="off">
      <div><label>Date</label><input type="date" id="f_date" required></div>
      <div><label>Company</label><select id="f_co"><option>Ankita</option><option>Adrita</option></select></div>
      <div><label>Type</label><select id="f_flow"><option value="Sale">Sale</option><option value="Purchase">Purchase</option></select></div>
      <div><label>Goods</label><select id="f_goods"><option value="FCN">FCN (kernel)</option><option value="RCN">RCN (raw nut)</option></select></div>
      <div><label>Payment status</label><select id="f_pay"><option>Paid</option><option>Pending</option><option>Partial</option></select></div>
      <div><label>Quantity (KG)</label><input type="number" id="f_qty" min="0" step="1" placeholder="optional"></div>
      <div><label>Rate (₹/KG)</label><input type="number" id="f_rate" min="0" step="0.01" placeholder="optional"></div>
      <div><label>Amount (₹)</label><input type="number" id="f_amt" min="0" step="1" placeholder="qty × rate, or type"></div>
      <div><label>Invoice / Voucher No.</label><input type="text" id="f_inv" placeholder="optional"></div>
      <div><label>Party (buyer/supplier)</label><input type="text" id="f_party" placeholder="optional"></div>
    </form>
    <div class="row-actions">
      <button class="btn" id="addBtn">+ Add record</button>
      <button class="btn alt" id="resetBtn">Reset my added records</button>
      <span class="note" id="addedCount"></span>
    </div>
    <div class="note">Tip: enter Quantity and Rate and the Amount fills in automatically (you can still type the Amount directly).</div>
  </div>

  <div class="toolbar">
    <div class="note" id="dataNote"></div>
    <button class="btn green" id="exportBtn">⬇ Export report to Excel</button>
  </div>

  <!-- KPIs -->
  <div class="grid kpis" id="kpis" style="margin-top:10px"></div>

  <!-- CHARTS -->
  <h2>Sales &amp; Purchase (recomputed live)</h2>
  <div class="grid two">
    <div class="card"><h3>Annual sales vs purchase (₹ crore)</h3><div id="chAnnual"></div>
      <div class="legend"><span><i class="dot" style="background:var(--navy)"></i>Sales</span><span><i class="dot" style="background:var(--orange)"></i>Purchase</span></div></div>
    <div class="card"><h3>Monthly sales trend (₹ lakh)</h3><div id="chMonthly"></div></div>
  </div>
  <div class="grid three" style="margin-top:14px">
    <div class="card"><h3>By company</h3><div id="tblCompany"></div></div>
    <div class="card"><h3>Revenue mix (sales)</h3><div id="chMix"></div>
      <div class="legend"><span><i class="dot" style="background:var(--navy)"></i>FCN kernel</span><span><i class="dot" style="background:var(--orange)"></i>RCN raw</span></div></div>
    <div class="card"><h3>Top buyers</h3><div id="tblBuyers"></div></div>
  </div>

  <!-- DESCRIPTIVE -->
  <h2>Descriptive statistics (recomputed live)</h2>
  <div class="card"><div id="tblStats"></div>
    <div class="note">Mean = average · Median = middle value · SD = spread · values in ₹.</div></div>

  <!-- YIELD -->
  <h2>Inventory yield &amp; ABC <span style="font-size:11px;color:var(--muted);font-weight:400">(from stock records — reference)</span></h2>
  <div class="grid two">
    <div class="card"><h3>Yield per cycle vs 22% target</h3><div id="chYield"></div>
      <div class="note">Red label = flagged by Z-score (|Z| &gt; 1.5).</div></div>
    <div class="card"><h3>ABC classification of stock by value</h3><div id="chABC"></div></div>
  </div>
  <div class="grid two" style="margin-top:14px">
    <div class="card"><h3>Yield detail</h3><div id="tblYield"></div></div>
    <div class="card"><h3>Forecast &amp; price (reference)</h3><div id="tblFore"></div></div>
  </div>

  <!-- ADDED -->
  <h2>My added records</h2>
  <div class="card"><div id="tblAdded"></div></div>

  <p class="note" style="margin:18px 2px">Self-contained file · works offline · your added records are saved in this browser until you reset them.</p>
</div>
<div class="toast" id="toast"></div>

<script>
const DATA = __DATA__;
const BASE = DATA.tx;
const S = DATA.static;
const LS_KEY = "ankita_added_v2";
let added = [];
try{ added = JSON.parse(localStorage.getItem(LS_KEY)) || []; }catch(e){ added=[]; }

const isSale = t => t.f==="Sale" || t.f==="Export";
const allTx = () => BASE.concat(added);
const sum = a => a.reduce((s,t)=>s+t.a,0);
function fyLabel(d){ const y=+d.slice(0,4), m=+d.slice(5,7); const st=m>=4?y:y-1; return "FY"+String(st).slice(2)+"-"+String(st+1).slice(2); }
function inr(n){ return Math.round(n).toLocaleString("en-IN"); }
function cr(n){ return (n/1e7).toFixed(2); }
function stats(arr){
  const v=arr.map(t=>t.a).sort((a,b)=>a-b); const n=v.length;
  if(!n) return {n:0,total:0,mean:0,median:0,sd:0,min:0,max:0};
  const total=v.reduce((s,x)=>s+x,0), mean=total/n;
  const median=n%2?v[(n-1)/2]:(v[n/2-1]+v[n/2])/2;
  const sd=Math.sqrt(v.reduce((s,x)=>s+(x-mean)*(x-mean),0)/(n>1?n-1:1));
  return {n,total,mean,median,sd,min:v[0],max:v[n-1]};
}
function compute(){
  const tx=allTx(), sales=tx.filter(isSale), pur=tx.filter(t=>t.f==="Purchase");
  const totSales=sum(sales), totPur=sum(pur);
  const A={}; tx.forEach(t=>{const fy=fyLabel(t.d);A[fy]=A[fy]||{sales:0,pur:0};if(isSale(t))A[fy].sales+=t.a;else A[fy].pur+=t.a;});
  const fys=Object.keys(A).sort();
  const M={}; sales.forEach(t=>{const ym=t.d.slice(0,7);M[ym]=(M[ym]||0)+t.a;});
  const months=Object.keys(M).sort();
  const C={}; tx.forEach(t=>{C[t.c]=C[t.c]||{sales:0,pur:0};if(isSale(t))C[t.c].sales+=t.a;else C[t.c].pur+=t.a;});
  const mix={FCN:0,RCN:0}; sales.forEach(t=>{mix[t.g]=(mix[t.g]||0)+t.a;});
  const B={}; sales.forEach(t=>{const k=t.p||"(blank)";B[k]=(B[k]||0)+t.a;});
  const buyers=Object.entries(B).sort((a,b)=>b[1]-a[1]);
  const top5=buyers.slice(0,5).reduce((s,x)=>s+x[1],0);
  return {tx,sales,pur,totSales,totPur,A,fys,M,months,C,mix,buyers,top5,sStat:stats(sales),pStat:stats(pur)};
}

// ---- SVG charts ----
function groupedBars(cats,s1,s2,fmt){const W=520,H=240,pad=38,bw=(W-2*pad)/cats.length,gap=bw*0.18;const max=Math.max(1,...s1,...s2),sc=(H-pad-20)/max;let g="";for(let i=0;i<=4;i++){const y=pad+(H-pad-20)*i/4,val=max*(1-i/4);g+=`<line x1="${pad}" y1="${y}" x2="${W-8}" y2="${y}" stroke="#eef1f5"/><text x="4" y="${y+3}" font-size="9" fill="#9aa3af">${fmt(val)}</text>`;}cats.forEach((c,i)=>{const x=pad+i*bw,w=(bw-gap)/2,h1=s1[i]*sc,h2=s2[i]*sc;g+=`<rect x="${x+gap/2}" y="${H-20-h1}" width="${w}" height="${h1}" fill="#1f4e79"><title>${c} Sales ${fmt(s1[i])}</title></rect><rect x="${x+gap/2+w}" y="${H-20-h2}" width="${w}" height="${h2}" fill="#c55a11"><title>${c} Purchase ${fmt(s2[i])}</title></rect><text x="${x+bw/2}" y="${H-6}" font-size="10" fill="#555" text-anchor="middle">${c}</text>`;});return `<svg viewBox="0 0 ${W} ${H}">${g}</svg>`;}
function lineChart(labels,vals,color,fmt){const W=520,H=240,pad=40,max=Math.max(1,...vals),min=Math.min(...vals,0),sc=(H-pad-22)/(max-min||1),step=(W-pad-10)/Math.max(1,labels.length-1);let g="";for(let i=0;i<=4;i++){const y=pad+(H-pad-22)*i/4,val=max-(max-min)*i/4;g+=`<line x1="${pad}" y1="${y}" x2="${W-8}" y2="${y}" stroke="#eef1f5"/><text x="4" y="${y+3}" font-size="9" fill="#9aa3af">${fmt(val)}</text>`;}let pts=vals.map((v,i)=>`${pad+i*step},${H-22-(v-min)*sc}`).join(" ");g+=`<polyline points="${pts}" fill="none" stroke="${color}" stroke-width="2"/>`;vals.forEach((v,i)=>{g+=`<circle cx="${pad+i*step}" cy="${H-22-(v-min)*sc}" r="2.2" fill="${color}"><title>${labels[i]} ${fmt(v)}</title></circle>`;});const k=Math.ceil(labels.length/8);labels.forEach((l,i)=>{if(i%k===0)g+=`<text x="${pad+i*step}" y="${H-6}" font-size="8" fill="#777" text-anchor="middle">${l.slice(2)}</text>`;});return `<svg viewBox="0 0 ${W} ${H}">${g}</svg>`;}
function donut(a,b){const total=a+b||1,R=70,r=42,cx=90,cy=90,frac=a/total;function arc(s0,e0,col){const s=s0*2*Math.PI-Math.PI/2,e=e0*2*Math.PI-Math.PI/2,x1=cx+R*Math.cos(s),y1=cy+R*Math.sin(s),x2=cx+R*Math.cos(e),y2=cy+R*Math.sin(e),big=(e0-s0)>0.5?1:0;return `<path d="M ${x1} ${y1} A ${R} ${R} 0 ${big} 1 ${x2} ${y2} L ${cx} ${cy} Z" fill="${col}"/>`;}let g=arc(0,frac,"#1f4e79")+arc(frac,1,"#c55a11")+`<circle cx="${cx}" cy="${cy}" r="${r}" fill="#fff"/><text x="${cx}" y="${cy-2}" font-size="15" font-weight="700" fill="#1f4e79" text-anchor="middle">${(frac*100).toFixed(1)}%</text><text x="${cx}" y="${cy+14}" font-size="9" fill="#777" text-anchor="middle">FCN kernel</text>`;return `<svg viewBox="0 0 180 180" style="max-width:200px;margin:auto">${g}</svg>`;}
function hbars(items,fmt){const W=520,rowH=26,H=items.length*rowH+10,pad=130,max=Math.max(1,...items.map(d=>d.value));let g="";items.forEach((d,i)=>{const y=8+i*rowH,w=(W-pad-50)*d.value/max;g+=`<text x="0" y="${y+13}" font-size="10" fill="#444">${d.label}</text><rect x="${pad}" y="${y+3}" width="${w}" height="14" rx="2" fill="${d.color}"><title>${d.label} ${fmt(d.value)}</title></rect><text x="${pad+w+4}" y="${y+14}" font-size="9.5" fill="#555">${fmt(d.value)}</text>`;});return `<svg viewBox="0 0 ${W} ${H}">${g}</svg>`;}
function vbars(cats,vals,colors,target,flags){const W=520,H=240,pad=34,bw=(W-2*pad)/cats.length,max=Math.max(target||0,...vals)*1.1,sc=(H-pad-20)/max;let g="";if(target){const y=H-20-target*sc;g+=`<line x1="${pad}" y1="${y}" x2="${W-8}" y2="${y}" stroke="#000" stroke-dasharray="4 3"/><text x="${W-8}" y="${y-3}" font-size="9" fill="#000" text-anchor="end">target ${target}%</text>`;}cats.forEach((c,i)=>{const x=pad+i*bw,h=vals[i]*sc,w=bw*0.62;g+=`<rect x="${x+bw*0.19}" y="${H-20-h}" width="${w}" height="${h}" fill="${colors[i]}"><title>${c} ${vals[i].toFixed(1)}%</title></rect><text x="${x+bw/2}" y="${H-20-h-3}" font-size="9" text-anchor="middle" fill="${flags&&flags[i]?'#c0392b':'#444'}" font-weight="${flags&&flags[i]?'700':'400'}">${vals[i].toFixed(1)}</text><text x="${x+bw/2}" y="${H-6}" font-size="8.5" text-anchor="middle" fill="#666">${c}</text>`;});return `<svg viewBox="0 0 ${W} ${H}">${g}</svg>`;}

const el=id=>document.getElementById(id);
function render(){
  const m=compute();
  const net=m.totSales-m.totPur, fcShare=m.mix.FCN/((m.mix.FCN+m.mix.RCN)||1)*100, top5pct=m.top5/(m.totSales||1)*100;
  const kpis=[["Total Sales","₹"+cr(m.totSales)+" cr","₹"+inr(m.totSales)],["Total Purchase","₹"+cr(m.totPur)+" cr","₹"+inr(m.totPur)],["Net (Sales−Purchase)","₹"+cr(net)+" cr",net>=0?"surplus":"deficit"],["Transactions",inr(m.tx.length),m.sales.length+" sales · "+m.pur.length+" purchases"],["Kernel (FCN) share",fcShare.toFixed(1)+"%","of sales value"],["Top-5 buyer share",top5pct.toFixed(1)+"%","of sales value"]];
  el("kpis").innerHTML=kpis.map(k=>`<div class="card kpi"><div class="label">${k[0]}</div><div class="val">${k[1]}</div><div class="sub">${k[2]}</div></div>`).join("");
  el("chAnnual").innerHTML=groupedBars(m.fys,m.fys.map(f=>m.A[f].sales),m.fys.map(f=>m.A[f].pur),v=>(v/1e7).toFixed(1));
  el("chMonthly").innerHTML=lineChart(m.months,m.months.map(x=>m.M[x]),"#1f4e79",v=>(v/1e5).toFixed(0));
  let ct=`<table><thead><tr><th>Company</th><th>Sales</th><th>Purchase</th><th>Net</th></tr></thead><tbody>`;
  Object.keys(m.C).sort().forEach(c=>{const r=m.C[c];ct+=`<tr><td>${c}</td><td class="num">${inr(r.sales)}</td><td class="num">${inr(r.pur)}</td><td class="num">${inr(r.sales-r.pur)}</td></tr>`;});
  el("tblCompany").innerHTML=ct+`</tbody></table>`;
  el("chMix").innerHTML=donut(m.mix.FCN,m.mix.RCN);
  let bt=`<table><thead><tr><th>Buyer</th><th>Sales (₹)</th></tr></thead><tbody>`;
  m.buyers.slice(0,7).forEach(b=>{bt+=`<tr><td>${b[0].length>26?b[0].slice(0,26)+"…":b[0]}</td><td class="num">${inr(b[1])}</td></tr>`;});
  el("tblBuyers").innerHTML=bt+`</tbody></table>`;
  const ss=m.sStat,ps=m.pStat,rows=[["Count",ss.n,ps.n],["Total",inr(ss.total),inr(ps.total)],["Mean",inr(ss.mean),inr(ps.mean)],["Median",inr(ss.median),inr(ps.median)],["SD",inr(ss.sd),inr(ps.sd)],["Min",inr(ss.min),inr(ps.min)],["Max",inr(ss.max),inr(ps.max)]];
  let st=`<table><thead><tr><th>Measure</th><th>Sales (₹)</th><th>Purchase (₹)</th></tr></thead><tbody>`;
  rows.forEach(r=>{st+=`<tr><td>${r[0]}</td><td class="num">${r[1]}</td><td class="num">${r[2]}</td></tr>`;});
  el("tblStats").innerHTML=st+`</tbody></table>`;
  renderAdded();
  el("addedCount").textContent=added.length?added.length+" record(s) added by you (saved in this browser).":"No records added yet.";
  el("dataNote").textContent="Showing "+inr(m.tx.length)+" records ("+inr(BASE.length)+" original + "+added.length+" added). Export reflects everything below.";
}
function renderStatic(){
  const yb=S.yield_batches;
  el("chYield").innerHTML=vbars(yb.map(b=>b.Company.slice(0,2)+"-"+b.Year.slice(2)),yb.map(b=>b.Yield_pct),yb.map(b=>b.Variance>=0?"#548235":"#c55a11"),22,yb.map(b=>b.Outlier==="Yes"));
  el("chABC").innerHTML=hbars(S.abc.map(a=>({label:a.Group.length>20?a.Group.slice(0,20)+"…":a.Group,value:a.Value,color:a.Class==="A"?"#c55a11":a.Class==="B"?"#1f4e79":"#9aa3af"})),v=>(v/1e7).toFixed(2)+"cr");
  let yt=`<table><thead><tr><th>Cycle</th><th>Yield%</th><th>Var</th><th>Z</th><th>Flag</th></tr></thead><tbody>`;
  yb.forEach(b=>{yt+=`<tr><td>${b.Company} ${b.Year}</td><td class="num">${b.Yield_pct.toFixed(1)}</td><td class="num">${b.Variance.toFixed(1)}</td><td class="num">${b.Zscore.toFixed(2)}</td><td>${b.Outlier==="Yes"?'<span class="flagYes">●</span>':''}</td></tr>`;});
  el("tblYield").innerHTML=yt+`</tbody></table><div class="note">Full-period: `+S.yield_company.map(c=>c.Company+" "+c.Recovery_pct.toFixed(1)+"%").join(" · ")+`</div>`;
  const fa=S.forecast_acc;
  let ft=`<table><thead><tr><th>Model</th><th>RMSE (₹)</th><th>MAE (₹)</th></tr></thead><tbody><tr><td>ARIMA</td><td class="num">${inr(fa.sarima_rmse)}</td><td class="num">${inr(fa.sarima_mae)}</td></tr><tr><td>Linear (base)</td><td class="num">${inr(fa.lin_rmse)}</td><td class="num">${inr(fa.lin_mae)}</td></tr><tr><td>Naive</td><td class="num">${inr(fa.naive_rmse)}</td><td class="num">${inr(fa.naive_mae)}</td></tr></tbody></table>`;
  const fn=S.forecast_next;ft+=`<div class="note">Next-month forecast: `+Object.keys(fn).map(k=>k+" ≈ ₹"+(fn[k].mean/1e7).toFixed(2)+"cr").join(" · ")+`</div><div class="note">RCN price/kg: `+Object.entries(S.rcn_price_yearly).map(([y,p])=>y+": ₹"+p).join(" · ")+`</div>`;
  el("tblFore").innerHTML=ft;
}
function renderAdded(){
  if(!added.length){el("tblAdded").innerHTML=`<div class="note">You have not added any records yet. Use the form at the top.</div>`;return;}
  let t=`<table class="added-tbl"><thead><tr><th>#</th><th>Date</th><th>Company</th><th>Type</th><th>Goods</th><th>Party</th><th style="text-align:right">Qty (KG)</th><th style="text-align:right">Amount (₹)</th><th>Payment</th><th>Invoice</th><th></th></tr></thead><tbody>`;
  added.forEach((r,i)=>{t+=`<tr><td>${i+1}</td><td>${r.d}</td><td>${r.c}</td><td><span class="pill ${isSale(r)?'sale':'pur'}">${r.f}</span></td><td>${r.g}</td><td>${r.p||""}</td><td class="num" style="text-align:right">${r.q?inr(r.q):"—"}</td><td class="num" style="text-align:right">${inr(r.a)}</td><td><span class="pill ${r.pay||'Paid'}">${r.pay||'Paid'}</span></td><td>${r.inv||""}</td><td><span class="x" data-i="${i}" title="remove">✕</span></td></tr>`;});
  el("tblAdded").innerHTML=t+`</tbody></table>`;
  el("tblAdded").querySelectorAll(".x").forEach(b=>b.onclick=()=>{added.splice(+b.dataset.i,1);save();render();toast("Record removed");});
}
function save(){try{localStorage.setItem(LS_KEY,JSON.stringify(added));}catch(e){}}
function toast(msg){const t=el("toast");t.textContent=msg;t.classList.add("show");setTimeout(()=>t.classList.remove("show"),1900);}

// ---- auto amount ----
function autoAmt(){const q=parseFloat(el("f_qty").value),r=parseFloat(el("f_rate").value);if(q>0&&r>0)el("f_amt").value=Math.round(q*r);}
el("f_qty").oninput=autoAmt; el("f_rate").oninput=autoAmt;

el("addBtn").onclick=()=>{
  const d=el("f_date").value, amt=parseFloat(el("f_amt").value);
  if(!d){toast("Please pick a date");return;}
  if(!(amt>0)){toast("Please enter a valid amount (or quantity and rate)");return;}
  const q=parseFloat(el("f_qty").value)||0, rt=parseFloat(el("f_rate").value)||0;
  added.push({c:el("f_co").value,f:el("f_flow").value,g:el("f_goods").value,d:d,p:el("f_party").value.trim(),a:amt,q:q,rt:rt,inv:el("f_inv").value.trim(),pay:el("f_pay").value});
  save(); ["f_amt","f_qty","f_rate","f_inv","f_party"].forEach(id=>el(id).value=""); render(); toast("Record added — totals updated");
};
el("resetBtn").onclick=()=>{if(added.length&&confirm("Remove all "+added.length+" records you added?")){added=[];save();render();toast("Your added records were cleared");}};

// ---- EXPORT TO EXCEL (SpreadsheetML 2003, no libraries) ----
function xesc(s){return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function xcell(v,style){const isNum=typeof v==="number"&&isFinite(v);const t=isNum?"Number":"String";const sid=style?` ss:StyleID="${style}"`:"";return `<Cell${sid}><Data ss:Type="${t}">${isNum?v:xesc(v)}</Data></Cell>`;}
function xsheet(name,rows){let r=rows.map((row,i)=>`<Row>${row.map(c=>xcell(c,i===0?"hdr":(typeof c==="number"?"num":null))).join("")}</Row>`).join("");return `<Worksheet ss:Name="${xesc(name)}"><Table>${r}</Table></Worksheet>`;}
function download(text,name,mime){const blob=new Blob([text],{type:mime});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();document.body.removeChild(a);setTimeout(()=>URL.revokeObjectURL(a.href),1000);}
function exportExcel(){
  const m=compute(), net=m.totSales-m.totPur;
  const fcShare=m.mix.FCN/((m.mix.FCN+m.mix.RCN)||1)*100, top5pct=m.top5/(m.totSales||1)*100;
  const sheets=[];
  sheets.push(xsheet("Summary",[["Metric","Value"],["Report generated",new Date().toLocaleString()],["Records (total)",m.tx.length],["— original",BASE.length],["— added by user",added.length],["Total Sales (₹)",Math.round(m.totSales)],["Total Purchase (₹)",Math.round(m.totPur)],["Net (₹)",Math.round(net)],["Kernel (FCN) share of sales (%)",+fcShare.toFixed(1)],["Top-5 buyer share of sales (%)",+top5pct.toFixed(1)]]));
  sheets.push(xsheet("Annual",[["Financial Year","Sales (₹)","Purchase (₹)","Net (₹)"]].concat(m.fys.map(f=>[f,Math.round(m.A[f].sales),Math.round(m.A[f].pur),Math.round(m.A[f].sales-m.A[f].pur)]))));
  sheets.push(xsheet("By Company",[["Company","Sales (₹)","Purchase (₹)","Net (₹)"]].concat(Object.keys(m.C).sort().map(c=>[c,Math.round(m.C[c].sales),Math.round(m.C[c].pur),Math.round(m.C[c].sales-m.C[c].pur)]))));
  const ss=m.sStat,ps=m.pStat;
  sheets.push(xsheet("Descriptive Stats",[["Measure","Sales (₹)","Purchase (₹)"],["Count",ss.n,ps.n],["Total",Math.round(ss.total),Math.round(ps.total)],["Mean",Math.round(ss.mean),Math.round(ps.mean)],["Median",Math.round(ss.median),Math.round(ps.median)],["SD",Math.round(ss.sd),Math.round(ps.sd)],["Min",Math.round(ss.min),Math.round(ps.min)],["Max",Math.round(ss.max),Math.round(ps.max)]]));
  sheets.push(xsheet("Monthly Sales",[["Year-Month","Sales (₹)"]].concat(m.months.map(x=>[x,Math.round(m.M[x])]))));
  const txRows=[["Company","Type","Goods","Date","Party","Amount (₹)","Quantity (KG)","Rate (₹/KG)","Invoice","Payment","Source"]];
  m.tx.forEach(t=>txRows.push([t.c,t.f,t.g,t.d,t.p||"",Math.round(t.a),t.q||"",t.rt||"",t.inv||"",t.pay||"",t.q!==undefined||t.inv!==undefined||t.pay!==undefined?"added":"original"]));
  sheets.push(xsheet("Transactions",txRows));
  sheets.push(xsheet("Top Buyers",[["Buyer","Sales (₹)"]].concat(m.buyers.slice(0,15).map(b=>[b[0],Math.round(b[1])]))));
  sheets.push(xsheet("Yield",[["Cycle","RCN In (KG)","Kernel Out (KG)","Yield %","Target %","Variance","Z-score","Outlier"]].concat(S.yield_batches.map(b=>[b.Company+" "+b.Year,Math.round(b.RCN_Purchased),Math.round(b.Kernel_Out),+b.Yield_pct.toFixed(1),22,+b.Variance.toFixed(1),+b.Zscore.toFixed(2),b.Outlier]))));
  sheets.push(xsheet("ABC",[["Stock Group","Value (₹)","% of Total","Cumulative %","Class"]].concat(S.abc.map(a=>[a.Group,Math.round(a.Value),+a.Pct.toFixed(1),+a.CumPct.toFixed(1),a.Class]))));
  const xml='<?xml version="1.0" encoding="UTF-8"?>\n<?mso-application progid="Excel.Sheet"?>\n'+
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'+
    '<Styles><Style ss:ID="hdr"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#1F4E79" ss:Pattern="Solid"/></Style>'+
    '<Style ss:ID="num"><NumberFormat ss:Format="#,##0"/></Style></Styles>'+
    sheets.join("")+'</Workbook>';
  const stamp=new Date().toISOString().slice(0,10);
  download(xml,"Cashew_Dashboard_Report_"+stamp+".xls","application/vnd.ms-excel");
  toast("Excel report downloaded ("+m.tx.length+" records, 9 sheets)");
}
el("exportBtn").onclick=exportExcel;

renderStatic();
render();
</script>
</body>
</html>'''

HTML = HTML.replace("__DATA__", data)
open("Cashew_Dashboard.html","w",encoding="utf-8").write(HTML)
print("written, size KB:", round(len(HTML.encode('utf-8'))/1024,1))
