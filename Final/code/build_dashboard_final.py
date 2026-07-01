import json
data=open('dash_data_final.json').read()
HTML = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Ankita & Adrita Cashew — Final Dashboard</title>
<style>
  :root{--teal:#0f766e;--amber:#b45309;--slate:#334155;--rose:#be123c;--green:#15803d;--bg:#f5f7f9;--card:#fff;--ink:#1f2937;--muted:#6b7280;--line:#e3e8ef;}
  *{box-sizing:border-box}
  body{margin:0;font-family:Arial,Helvetica,sans-serif;background:var(--bg);color:var(--ink);font-size:14px}
  header{background:var(--teal);color:#fff;padding:16px 22px}
  header h1{margin:0;font-size:19px} header p{margin:4px 0 0;font-size:12.5px;opacity:.9}
  .wrap{max-width:1200px;margin:0 auto;padding:18px}
  .grid{display:grid;gap:14px}
  .kpis{grid-template-columns:repeat(6,1fr)}
  @media(max-width:900px){.kpis{grid-template-columns:repeat(3,1fr)}}
  @media(max-width:560px){.kpis{grid-template-columns:repeat(2,1fr)}}
  .card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px;box-shadow:0 1px 2px rgba(0,0,0,.03)}
  .kpi .label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em}
  .kpi .val{font-size:21px;font-weight:700;color:var(--teal);margin-top:4px}
  .kpi .sub{font-size:11px;color:var(--muted);margin-top:2px}
  h2{font-size:15px;color:var(--teal);margin:24px 2px 8px;border-bottom:2px solid var(--line);padding-bottom:5px}
  h2 .tag{font-size:10.5px;color:#fff;background:var(--amber);padding:1px 7px;border-radius:10px;vertical-align:middle;margin-left:8px;font-weight:600}
  h3{font-size:13px;color:var(--amber);margin:0 0 8px}
  .two{grid-template-columns:1fr 1fr} .three{grid-template-columns:1.2fr 1fr 1fr}
  @media(max-width:860px){.two,.three{grid-template-columns:1fr}}
  table{width:100%;border-collapse:collapse;font-size:12.5px}
  th,td{padding:6px 8px;border-bottom:1px solid var(--line);text-align:right}
  th:first-child,td:first-child{text-align:left}
  thead th{background:var(--teal);color:#fff;font-weight:600;border:none}
  tbody tr:nth-child(even){background:#eef6f4}
  .num{font-variant-numeric:tabular-nums}
  svg{width:100%;height:auto;display:block}
  form{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;align-items:end}
  @media(max-width:860px){form{grid-template-columns:repeat(2,1fr)}}
  label{display:block;font-size:11px;color:var(--muted);margin-bottom:3px}
  input,select{width:100%;padding:7px 8px;border:1px solid #cdd5e0;border-radius:6px;font-size:13px;background:#fff}
  .btn{background:var(--teal);color:#fff;border:none;padding:9px 14px;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600}
  .btn:hover{filter:brightness(.92)} .btn.alt{background:#fff;color:var(--teal);border:1px solid var(--teal)}
  .btn.amber{background:var(--amber)}
  .row-actions{display:flex;gap:8px;margin-top:10px;flex-wrap:wrap;align-items:center}
  .toolbar{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-top:14px}
  .pill{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600}
  .pill.sale{background:#dcfce7;color:#166534}.pill.pur{background:#fee2e2;color:#991b1b}
  .pill.Low{background:#dcfce7;color:#166534}.pill.Medium{background:#fef3c7;color:#92400e}.pill.High{background:#fee2e2;color:#991b1b}
  .pill.Paid{background:#dcfce7;color:#166534}.pill.Pending{background:#fee2e2;color:#991b1b}.pill.Partial{background:#fef3c7;color:#92400e}
  .toast{position:fixed;bottom:18px;left:50%;transform:translateX(-50%);background:#0b3b36;color:#fff;padding:10px 16px;border-radius:8px;font-size:13px;opacity:0;transition:.3s;pointer-events:none;z-index:9}
  .toast.show{opacity:1}
  .x{color:var(--rose);cursor:pointer;font-weight:700}
  .legend{display:flex;gap:14px;flex-wrap:wrap;font-size:11px;color:var(--muted);margin-top:6px}
  .legend span{display:inline-flex;align-items:center;gap:5px}.dot{width:10px;height:10px;border-radius:2px;display:inline-block}
  .note{font-size:11px;color:var(--muted);margin-top:6px}
  .assume{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}
  @media(max-width:700px){.assume{grid-template-columns:repeat(2,1fr)}}
  .big{font-size:20px;font-weight:700;color:var(--teal)}
  .sens td{text-align:center;font-size:11.5px}
  .mini{font-size:11px;color:var(--muted)}
</style></head>
<body>
<header>
  <h1>Ankita &amp; Adrita Cashew Processing — Final Dashboard</h1>
  <p>BDM Capstone · Akash Jana · 23f2000990 · live recompute · P&amp;L · buyer risk · anomaly detection · Excel export</p>
</header>
<div class="wrap">

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
      <div><label>Invoice No.</label><input type="text" id="f_inv" placeholder="optional"></div>
      <div><label>Party (buyer/supplier)</label><input type="text" id="f_party" placeholder="optional"></div>
    </form>
    <div class="row-actions">
      <button class="btn" id="addBtn">+ Add record</button>
      <button class="btn alt" id="resetBtn">Reset my added records</button>
      <span class="note" id="addedCount"></span>
    </div>
  </div>

  <div class="toolbar"><div class="note" id="dataNote"></div><button class="btn amber" id="exportBtn">⬇ Export full report to Excel</button></div>
  <div class="grid kpis" id="kpis" style="margin-top:10px"></div>

  <h2>Sales &amp; Purchase</h2>
  <div class="grid two">
    <div class="card"><h3>Annual sales vs purchase (₹ crore)</h3><div id="chAnnual"></div>
      <div class="legend"><span><i class="dot" style="background:var(--teal)"></i>Sales</span><span><i class="dot" style="background:var(--amber)"></i>Purchase</span></div></div>
    <div class="card"><h3>Monthly sales trend (₹ lakh)</h3><div id="chMonthly"></div></div>
  </div>
  <div class="grid three" style="margin-top:14px">
    <div class="card"><h3>By company</h3><div id="tblCompany"></div></div>
    <div class="card"><h3>Revenue mix (sales)</h3><div id="chMix"></div>
      <div class="legend"><span><i class="dot" style="background:var(--teal)"></i>FCN</span><span><i class="dot" style="background:var(--amber)"></i>RCN</span></div></div>
    <div class="card"><h3>Top buyers</h3><div id="tblBuyers"></div></div>
  </div>

  <h2>Descriptive statistics</h2>
  <div class="card"><div id="tblStats"></div></div>

  <h2>Profit &amp; Loss model <span class="tag">edit the assumptions</span></h2>
  <div class="card">
    <div class="assume">
      <div><label>RCN price (₹/kg)</label><input type="number" id="p_rcn" step="0.1"></div>
      <div><label>Outturn / yield (%)</label><input type="number" id="p_out" step="0.5"></div>
      <div><label>Processing (₹/kg kernel)</label><input type="number" id="p_proc" step="1"></div>
      <div><label>Kernel selling price (₹/kg)</label><input type="number" id="p_price" step="1"></div>
    </div>
    <div class="grid three" style="margin-top:12px">
      <div><div class="mini">Contribution / kg kernel</div><div class="big" id="pnl_contrib"></div><div class="mini" id="pnl_margin"></div></div>
      <div><div class="mini">Break-even price</div><div class="big" id="pnl_be"></div><div class="mini" id="pnl_beout"></div></div>
      <div><div class="mini">Indicative annual contribution</div><div class="big" id="pnl_annual"></div><div class="mini">on kernel volume sold</div></div>
    </div>
  </div>
  <div class="grid two" style="margin-top:14px">
    <div class="card"><h3>Profit build-up per kg kernel</h3><div id="chWaterfall"></div></div>
    <div class="card"><h3>Contribution by grade</h3><div id="chGrade"></div>
      <div class="legend"><span id="gradeNote" class="mini"></span></div></div>
  </div>
  <div class="card" style="margin-top:14px"><h3>Sensitivity — contribution (₹/kg) vs price moves</h3><div id="tblSens"></div>
    <div class="note">Rows = RCN price change · Columns = selling-price change · green = profit, red = loss.</div></div>

  <h2>Buyer risk &amp; credit <span class="tag">live scorecard</span></h2>
  <div class="grid three">
    <div class="card"><h3>Risk tiers</h3><div id="chTiers"></div></div>
    <div class="card" style="grid-column:span 2"><h3>Buyer risk map (revenue vs recency)</h3><div id="chBuyerMap"></div>
      <div class="legend"><span><i class="dot" style="background:var(--green)"></i>Low</span><span><i class="dot" style="background:var(--amber)"></i>Medium</span><span><i class="dot" style="background:var(--rose)"></i>High</span><span class="mini">bubble size = orders</span></div></div>
  </div>
  <div class="grid two" style="margin-top:14px">
    <div class="card"><h3>Top buyers — scorecard</h3><div id="tblScore"></div></div>
    <div class="card"><h3>Illustrative payment-delay model (ROC)</h3><div id="chRoc"></div>
      <div class="note" id="rocNote"></div></div>
  </div>

  <h2>Anomaly detection <span class="tag">live Z-score</span></h2>
  <div class="grid two">
    <div class="card"><h3>Flagged transactions</h3><div id="chAnom"></div>
      <div class="legend"><span><i class="dot" style="background:#cbd5e1"></i>normal</span><span><i class="dot" style="background:var(--rose)"></i>flagged</span></div></div>
    <div class="card"><h3>Summary</h3><div id="anomSummary"></div></div>
  </div>
  <div class="card" style="margin-top:14px"><h3>Top flagged transactions</h3><div id="tblAnom"></div></div>

  <h2>Inventory yield <span style="font-size:11px;color:var(--muted);font-weight:400">(reference)</span></h2>
  <div class="card"><div id="tblYield"></div></div>

  <h2>My added records</h2>
  <div class="card"><div id="tblAdded"></div></div>
  <p class="note" style="margin:18px 2px">Self-contained · works offline · added records saved in this browser until reset. Isolation Forest and the trained logistic model run in Python (see report); this page recomputes sales, P&amp;L, buyer scores and Z-score anomalies live.</p>
</div>
<div class="toast" id="toast"></div>

<script>
const DATA=__DATA__; const BASE=DATA.tx; const S=DATA.static;
const LS="ankita_final_v1"; let added=[]; try{added=JSON.parse(localStorage.getItem(LS))||[];}catch(e){added=[];}
const DATA_END=new Date(S.data_end);
const isSale=t=>t.f==="Sale"||t.f==="Export"; const allTx=()=>BASE.concat(added);
const sum=a=>a.reduce((s,t)=>s+t.a,0);
const inr=n=>Math.round(n).toLocaleString("en-IN"); const cr=n=>(n/1e7).toFixed(2); const lk=n=>(n/1e5).toFixed(1);
function fyLabel(d){const y=+d.slice(0,4),m=+d.slice(5,7);const st=m>=4?y:y-1;return "FY"+String(st).slice(2)+"-"+String(st+1).slice(2);}
function stats(arr){const v=arr.slice().sort((a,b)=>a-b);const n=v.length;if(!n)return{n:0,total:0,mean:0,median:0,sd:0,min:0,max:0};const total=v.reduce((s,x)=>s+x,0),mean=total/n;const median=n%2?v[(n-1)/2]:(v[n/2-1]+v[n/2])/2;const sd=Math.sqrt(v.reduce((s,x)=>s+(x-mean)*(x-mean),0)/(n>1?n-1:1));return{n,total,mean,median,sd,min:v[0],max:v[n-1]};}
const el=id=>document.getElementById(id);
const C={teal:"#0f766e",amber:"#b45309",slate:"#334155",rose:"#be123c",green:"#15803d",grey:"#94a3b8"};

// ---- generic SVG helpers ----
function groupedBars(cats,s1,s2,fmt){const W=520,H=230,pad=38,bw=(W-2*pad)/cats.length,gap=bw*.18,max=Math.max(1,...s1,...s2),sc=(H-pad-20)/max;let g="";for(let i=0;i<=4;i++){const y=pad+(H-pad-20)*i/4;g+=`<line x1="${pad}" y1="${y}" x2="${W-8}" y2="${y}" stroke="#eef1f5"/><text x="4" y="${y+3}" font-size="9" fill="#9aa3af">${fmt(max*(1-i/4))}</text>`;}cats.forEach((c,i)=>{const x=pad+i*bw,w=(bw-gap)/2;g+=`<rect x="${x+gap/2}" y="${H-20-s1[i]*sc}" width="${w}" height="${s1[i]*sc}" fill="${C.teal}"><title>${c} ${fmt(s1[i])}</title></rect><rect x="${x+gap/2+w}" y="${H-20-s2[i]*sc}" width="${w}" height="${s2[i]*sc}" fill="${C.amber}"><title>${c} ${fmt(s2[i])}</title></rect><text x="${x+bw/2}" y="${H-6}" font-size="10" fill="#555" text-anchor="middle">${c}</text>`;});return `<svg viewBox="0 0 ${W} ${H}">${g}</svg>`;}
function lineChart(labels,vals,color,fmt){const W=520,H=230,pad=40,max=Math.max(1,...vals),min=Math.min(...vals,0),sc=(H-pad-22)/(max-min||1),step=(W-pad-10)/Math.max(1,labels.length-1);let g="";for(let i=0;i<=4;i++){const y=pad+(H-pad-22)*i/4;g+=`<line x1="${pad}" y1="${y}" x2="${W-8}" y2="${y}" stroke="#eef1f5"/><text x="4" y="${y+3}" font-size="9" fill="#9aa3af">${fmt(max-(max-min)*i/4)}</text>`;}g+=`<polyline points="${vals.map((v,i)=>`${pad+i*step},${H-22-(v-min)*sc}`).join(" ")}" fill="none" stroke="${color}" stroke-width="2"/>`;const k=Math.ceil(labels.length/8);labels.forEach((l,i)=>{if(i%k===0)g+=`<text x="${pad+i*step}" y="${H-6}" font-size="8" fill="#777" text-anchor="middle">${l.slice(2)}</text>`;});return `<svg viewBox="0 0 ${W} ${H}">${g}</svg>`;}
function donut(a,b){const total=a+b||1,R=70,r=42,cx=90,cy=90,frac=a/total;function arc(s0,e0,col){const s=s0*2*Math.PI-Math.PI/2,e=e0*2*Math.PI-Math.PI/2,x1=cx+R*Math.cos(s),y1=cy+R*Math.sin(s),x2=cx+R*Math.cos(e),y2=cy+R*Math.sin(e),big=(e0-s0)>.5?1:0;return `<path d="M ${x1} ${y1} A ${R} ${R} 0 ${big} 1 ${x2} ${y2} L ${cx} ${cy} Z" fill="${col}"/>`;}return `<svg viewBox="0 0 180 180" style="max-width:200px;margin:auto">${arc(0,frac,C.teal)+arc(frac,1,C.amber)}<circle cx="${cx}" cy="${cy}" r="${r}" fill="#fff"/><text x="${cx}" y="${cy-2}" font-size="15" font-weight="700" fill="${C.teal}" text-anchor="middle">${(frac*100).toFixed(1)}%</text><text x="${cx}" y="${cy+14}" font-size="9" fill="#777" text-anchor="middle">FCN</text></svg>`;}
function scatter(pts,opts){const W=520,H=270,pad=46;const xs=pts.map(p=>p.x),ys=pts.map(p=>p.y);const xmax=Math.max(1,...xs),ymax=Math.max(1,...ys);let g="";for(let i=0;i<=4;i++){const y=pad+(H-pad-24)*i/4;g+=`<line x1="${pad}" y1="${y}" x2="${W-8}" y2="${y}" stroke="#eef1f5"/><text x="4" y="${y+3}" font-size="8" fill="#9aa3af">${opts.fy(ymax*(1-i/4))}</text>`;}pts.forEach(p=>{const cx=pad+(p.x/xmax)*(W-pad-12),cy=H-24-(p.y/ymax)*(H-pad-24);g+=`<circle cx="${cx}" cy="${cy}" r="${p.r||4}" fill="${p.c}" fill-opacity="0.62" stroke="#fff" stroke-width="0.6"><title>${p.t||""}</title></circle>`;});g+=`<text x="${W/2}" y="${H-4}" font-size="9" fill="#777" text-anchor="middle">${opts.xl}</text>`;return `<svg viewBox="0 0 ${W} ${H}">${g}</svg>`;}
function vbarsSigned(cats,vals,fmt){const W=420,H=230,pad=40,bw=(W-2*pad)/cats.length;const mx=Math.max(...vals.map(Math.abs))*1.25||1,sc=(H-50)/(2*mx),zero=H/2-10;let g=`<line x1="${pad-6}" y1="${zero}" x2="${W-8}" y2="${zero}" stroke="#334155"/>`;cats.forEach((c,i)=>{const x=pad+i*bw,h=vals[i]*sc,col=vals[i]>=0?C.green:C.rose;g+=`<rect x="${x+bw*.2}" y="${vals[i]>=0?zero-h:zero}" width="${bw*.6}" height="${Math.abs(h)}" fill="${col}"/><text x="${x+bw/2}" y="${vals[i]>=0?zero-h-4:zero+h+14}" font-size="10" text-anchor="middle" font-weight="700" fill="${col}">${fmt(vals[i])}</text><text x="${x+bw/2}" y="${H-4}" font-size="10" text-anchor="middle" fill="#555">${c}</text>`;});return `<svg viewBox="0 0 ${W} ${H}">${g}</svg>`;}
function waterfall(price,rcncost,proc,contrib){const steps=[["Price",price,C.teal,false],["−RCN",-rcncost,C.amber,true],["−Proc",-proc,C.amber,true],["Contribution",contrib,C.green,false]];const W=460,H=240,pad=34,bw=(W-2*pad)/4,max=price*1.15,sc=(H-40)/max;let cum=0,g="";steps.forEach((s,i)=>{const[lab,val,col,delta]=s;const x=pad+i*bw;if(i===3){g+=`<rect x="${x+bw*.18}" y="${H-20-contrib*sc}" width="${bw*.64}" height="${contrib*sc}" fill="${col}"/><text x="${x+bw/2}" y="${H-20-contrib*sc-5}" font-size="10" font-weight="700" text-anchor="middle" fill="${col}">₹${Math.round(contrib)}</text>`;}else{const start=cum,h=val*sc;g+=`<rect x="${x+bw*.18}" y="${H-20-(start+ (val>0?val:0))*sc}" width="${bw*.64}" height="${Math.abs(h)}" fill="${col}"/><text x="${x+bw/2}" y="${H-20-(start+(val>0?val:0))*sc-4}" font-size="9" text-anchor="middle" fill="#444">${val>0?"+":""}${Math.round(val)}</text>`;cum+=val;}g+=`<text x="${x+bw/2}" y="${H-5}" font-size="9.5" text-anchor="middle" fill="#555">${lab}</text>`;});return `<svg viewBox="0 0 ${W} ${H}">${g}</svg>`;}
function rocChart(fpr,tpr,auc){const W=300,H=270,pad=40;let g=`<line x1="${pad}" y1="${H-pad}" x2="${W-10}" y2="${pad-10}" stroke="#cbd5e1" stroke-dasharray="4 3"/>`;let pts=fpr.map((f,i)=>`${pad+f*(W-pad-10)},${(H-pad)-tpr[i]*(H-pad-(pad-10))}`).join(" ");g+=`<polyline points="${pts}" fill="none" stroke="${C.teal}" stroke-width="2.4"/>`;g+=`<text x="${pad+8}" y="${pad}" font-size="11" font-weight="700" fill="${C.teal}">AUC = ${auc}</text>`;g+=`<text x="${W/2}" y="${H-6}" font-size="9" fill="#777" text-anchor="middle">False positive rate</text>`;return `<svg viewBox="0 0 ${W} ${H}">${g}</svg>`;}

// ---- compute: sales ----
function computeSales(){const tx=allTx(),sales=tx.filter(isSale),pur=tx.filter(t=>t.f==="Purchase");const A={};tx.forEach(t=>{const fy=fyLabel(t.d);A[fy]=A[fy]||{s:0,p:0};if(isSale(t))A[fy].s+=t.a;else A[fy].p+=t.a;});const fys=Object.keys(A).sort();const M={};sales.forEach(t=>{const ym=t.d.slice(0,7);M[ym]=(M[ym]||0)+t.a;});const months=Object.keys(M).sort();const Co={};tx.forEach(t=>{Co[t.c]=Co[t.c]||{s:0,p:0};if(isSale(t))Co[t.c].s+=t.a;else Co[t.c].p+=t.a;});const mix={FCN:0,RCN:0};sales.forEach(t=>{mix[t.g]=(mix[t.g]||0)+t.a;});const B={};sales.forEach(t=>{const k=t.p||"(blank)";B[k]=(B[k]||0)+t.a;});const buyers=Object.entries(B).sort((a,b)=>b[1]-a[1]);return{tx,sales,pur,totS:sum(sales),totP:sum(pur),A,fys,M,months,Co,mix,buyers,sStat:stats(sales.map(t=>t.a)),pStat:stats(pur.map(t=>t.a))};}

// ---- compute: P&L ----
function pnl(){const rcn=+el("p_rcn").value,out=+el("p_out").value/100,proc=+el("p_proc").value,price=+el("p_price").value;
  const rcnCost=rcn/out,total=rcnCost+proc,contrib=price-total,margin=contrib/price*100;
  const beP=total, beOut=rcn/(price-proc)*100;
  const ws=S.pnl.whole_share/100, wp=S.pnl.whole_price, bp=S.pnl.broken_price;
  // keep grade prices fixed (assumption), recompute their contribution at current cost
  const wc=wp-total, bc=bp-total;
  const steps=[-0.10,-0.05,0,0.05,0.10];const sens=steps.map(dr=>steps.map(dp=>Math.round(price*(1+dp)-(rcn*(1+dr)/out+proc))));
  return{rcn,out,proc,price,rcnCost,total,contrib,margin,beP,beOut,wc,bc,wp,bp,steps,sens};}

// ---- compute: buyers ----
const BADP=new Set(["Cash","Advance From Customers",""]);
function nrm(arr,inv){const lo=Math.min(...arr),hi=Math.max(...arr);return arr.map(x=>hi>lo?(inv?1-(x-lo)/(hi-lo):(x-lo)/(hi-lo)):0);}
function computeBuyers(){const sales=allTx().filter(isSale);const g={};sales.forEach(t=>{if(BADP.has(t.p))return;const k=t.p;(g[k]=g[k]||[]).push(t);});
  let rows=Object.entries(g).map(([buyer,arr])=>{const rev=sum(arr);const orders=arr.length;const dates=arr.map(t=>new Date(t.d));const last=new Date(Math.max(...dates)),first=new Date(Math.min(...dates));const recency=Math.round((DATA_END-last)/86400000);const tenure=Math.max(1,Math.round((last-first)/86400000));const mset=new Set(arr.map(t=>t.d.slice(0,7)));const months=mset.size;return{buyer,rev,orders,avg:rev/orders,recency,tenure,months,freq:orders/Math.max(1,months)};});
  if(!rows.length)return{rows:[],tiers:{Low:0,Medium:0,High:0},exposure:{Low:0,Medium:0,High:0}};
  const recN=nrm(rows.map(r=>r.recency)),frN=nrm(rows.map(r=>r.freq),true),teN=nrm(rows.map(r=>r.tenure),true),avN=nrm(rows.map(r=>r.avg));
  const mult={Low:2.0,Medium:1.0,High:0.4};
  rows.forEach((r,i)=>{r.risk=Math.round((0.35*recN[i]+0.25*frN[i]+0.20*teN[i]+0.20*avN[i])*1000)/10;r.tier=r.risk<=33?"Low":r.risk<=66?"Medium":"High";r.credit=Math.round((r.rev/Math.max(1,r.months))*mult[r.tier]/1000)*1000;});
  rows.sort((a,b)=>b.rev-a.rev);
  const tiers={Low:0,Medium:0,High:0},exposure={Low:0,Medium:0,High:0};
  rows.forEach(r=>{tiers[r.tier]++;exposure[r.tier]+=r.credit;});
  return{rows,tiers,exposure};}

// ---- compute: anomalies (Z-score live) ----
function computeAnom(){const tx=allTx();const grp={};tx.forEach((t,i)=>{const k=t.c+"|"+t.f+"|"+t.g;(grp[k]=grp[k]||[]).push(i);});const z=new Array(tx.length).fill(0);Object.values(grp).forEach(idx=>{const v=idx.map(i=>tx[i].a);const m=v.reduce((s,x)=>s+x,0)/v.length;const sd=Math.sqrt(v.reduce((s,x)=>s+(x-m)*(x-m),0)/v.length)||1;idx.forEach(i=>z[i]=(tx[i].a-m)/sd);});
  const flags=[];tx.forEach((t,i)=>{if(Math.abs(z[i])>3)flags.push({...t,z:z[i]});});flags.sort((a,b)=>Math.abs(b.z)-Math.abs(a.z));return{tx,z,flags};}

// ---- RENDER ----
function render(){
  const m=computeSales();
  const net=m.totS-m.totP,fcS=m.mix.FCN/((m.mix.FCN+m.mix.RCN)||1)*100;
  el("kpis").innerHTML=[["Total Sales","₹"+cr(m.totS)+" cr","₹"+inr(m.totS)],["Total Purchase","₹"+cr(m.totP)+" cr","₹"+inr(m.totP)],["Net","₹"+cr(net)+" cr",net>=0?"surplus":"deficit"],["Transactions",inr(m.tx.length),m.sales.length+" sales · "+m.pur.length+" buys"],["Kernel share",fcS.toFixed(1)+"%","of sales"],["Buyers",computeBuyers().rows.length,"with credit risk scored"]].map(k=>`<div class="card kpi"><div class="label">${k[0]}</div><div class="val">${k[1]}</div><div class="sub">${k[2]}</div></div>`).join("");
  el("chAnnual").innerHTML=groupedBars(m.fys,m.fys.map(f=>m.A[f].s),m.fys.map(f=>m.A[f].p),v=>(v/1e7).toFixed(1));
  el("chMonthly").innerHTML=lineChart(m.months,m.months.map(x=>m.M[x]),C.teal,v=>(v/1e5).toFixed(0));
  let ct=`<table><thead><tr><th>Company</th><th>Sales</th><th>Purchase</th><th>Net</th></tr></thead><tbody>`;Object.keys(m.Co).sort().forEach(c=>{const r=m.Co[c];ct+=`<tr><td>${c}</td><td class="num">${inr(r.s)}</td><td class="num">${inr(r.p)}</td><td class="num">${inr(r.s-r.p)}</td></tr>`;});el("tblCompany").innerHTML=ct+`</tbody></table>`;
  el("chMix").innerHTML=donut(m.mix.FCN,m.mix.RCN);
  let bt=`<table><thead><tr><th>Buyer</th><th>Sales (₹)</th></tr></thead><tbody>`;m.buyers.slice(0,7).forEach(b=>{bt+=`<tr><td>${b[0].length>24?b[0].slice(0,24)+"…":b[0]}</td><td class="num">${inr(b[1])}</td></tr>`;});el("tblBuyers").innerHTML=bt+`</tbody></table>`;
  const ss=m.sStat,ps=m.pStat;let st=`<table><thead><tr><th>Measure</th><th>Sales (₹)</th><th>Purchase (₹)</th></tr></thead><tbody>`;[["Count",ss.n,ps.n],["Total",inr(ss.total),inr(ps.total)],["Mean",inr(ss.mean),inr(ps.mean)],["Median",inr(ss.median),inr(ps.median)],["SD",inr(ss.sd),inr(ps.sd)],["Min",inr(ss.min),inr(ps.min)],["Max",inr(ss.max),inr(ps.max)]].forEach(r=>{st+=`<tr><td>${r[0]}</td><td class="num">${r[1]}</td><td class="num">${r[2]}</td></tr>`;});el("tblStats").innerHTML=st+`</tbody></table>`;
  renderPnL(m); renderBuyers(); renderAnom(); renderYield(); renderAdded();
  el("addedCount").textContent=added.length?added.length+" record(s) added (saved in browser).":"No records added yet.";
  el("dataNote").textContent="Showing "+inr(m.tx.length)+" records ("+inr(BASE.length)+" original + "+added.length+" added).";
}
function renderPnL(m){m=m||computeSales();const p=pnl();
  el("pnl_contrib").textContent="₹"+Math.round(p.contrib)+"/kg";el("pnl_margin").textContent=p.margin.toFixed(1)+"% margin";
  el("pnl_be").textContent="₹"+Math.round(p.beP)+"/kg";el("pnl_beout").textContent="break-even outturn "+p.beOut.toFixed(1)+"%";
  const kernKG=338783;el("pnl_annual").textContent="₹"+cr(kernKG/4*p.contrib)+" cr/yr";
  el("chWaterfall").innerHTML=waterfall(p.price,p.rcnCost,p.proc,p.contrib);
  el("chGrade").innerHTML=vbarsSigned(["Whole","Broken"],[p.wc,p.bc],v=>"₹"+Math.round(v));
  el("gradeNote").textContent=`Whole ≈ ₹${p.wp}/kg · Broken ≈ ₹${p.bp}/kg (assumed); profit comes from whole kernels.`;
  // sensitivity table
  const lab=p.steps.map(s=>(s*100>0?"+":"")+(s*100)+"%");
  let h=`<table class="sens"><thead><tr><th>RCN ↓ / Price →</th>`+lab.map(l=>`<th>${l}</th>`).join("")+`</tr></thead><tbody>`;
  p.sens.forEach((row,i)=>{h+=`<tr><td style="font-weight:600">${lab[i]}</td>`+row.map(v=>{const t=Math.max(-50,Math.min(250,v));const g=t>=0?`background:rgba(21,128,61,${.08+.5*t/250})`:`background:rgba(190,18,60,${.1+.5*(-t)/50})`;return `<td style="${g}">${v}</td>`;}).join("")+`</tr>`;});
  el("tblSens").innerHTML=h+`</tbody></table>`;
}
function renderBuyers(){const b=computeBuyers();
  el("chTiers").innerHTML=(()=>{const t=["Low","Medium","High"],cnt=t.map(x=>b.tiers[x]);const W=300,H=210,pad=30,bw=(W-2*pad)/3,max=Math.max(1,...cnt);let g="";const col={Low:C.green,Medium:C.amber,High:C.rose};t.forEach((x,i)=>{const h=cnt[i]/max*(H-50),X=pad+i*bw;g+=`<rect x="${X+bw*.2}" y="${H-26-h}" width="${bw*.6}" height="${h}" fill="${col[x]}"/><text x="${X+bw/2}" y="${H-26-h-5}" font-size="11" text-anchor="middle" font-weight="700">${cnt[i]}</text><text x="${X+bw/2}" y="${H-8}" font-size="10" text-anchor="middle">${x}</text>`;});return `<svg viewBox="0 0 ${W} ${H}">${g}</svg>`;})();
  const col={Low:C.green,Medium:C.amber,High:C.rose};
  el("chBuyerMap").innerHTML=scatter(b.rows.map(r=>({x:r.recency,y:r.rev,r:Math.max(3,Math.min(16,r.orders*0.7)),c:col[r.tier],t:r.buyer+" · ₹"+inr(r.rev)+" · risk "+r.risk})),{xl:"Days since last order (recency)",fy:v=>(v/1e7).toFixed(1)});
  let s=`<table><thead><tr><th>Buyer</th><th>Revenue</th><th>Ord</th><th>Recency</th><th>Risk</th><th>Tier</th><th>Credit (₹)</th></tr></thead><tbody>`;
  b.rows.slice(0,10).forEach(r=>{s+=`<tr><td>${r.buyer.length>20?r.buyer.slice(0,20)+"…":r.buyer}</td><td class="num">${inr(r.rev)}</td><td class="num">${r.orders}</td><td class="num">${r.recency}</td><td class="num">${r.risk}</td><td><span class="pill ${r.tier}">${r.tier}</span></td><td class="num">${inr(r.credit)}</td></tr>`;});
  el("tblScore").innerHTML=s+`</tbody></table><div class="note">Tiers — Low: ${b.tiers.Low} · Medium: ${b.tiers.Medium} · High: ${b.tiers.High}. Exposure ₹${cr(b.exposure.Low+b.exposure.Medium+b.exposure.High)} cr.</div>`;
  el("chRoc").innerHTML=rocChart(S.roc.fpr,S.roc.tpr,S.lr.auc);
  el("rocNote").textContent=`Illustrative logistic model on simulated delays (AUC ${S.lr.auc}, accuracy ${(S.lr.acc*100).toFixed(0)}%). Real model needs recorded payment dates.`;
}
function renderAnom(){const a=computeAnom();
  el("chAnom").innerHTML=(()=>{const tx=a.tx;const W=520,H=250,pad=44;const amts=tx.map(t=>Math.log10(t.a+1));const ymin=Math.min(...amts),ymax=Math.max(...amts);let g="";for(let i=0;i<=4;i++){const y=pad+(H-pad-24)*i/4;const val=Math.pow(10,ymax-(ymax-ymin)*i/4);g+=`<line x1="${pad}" y1="${y}" x2="${W-8}" y2="${y}" stroke="#eef1f5"/><text x="2" y="${y+3}" font-size="8" fill="#9aa3af">${(val/1e5).toFixed(0)}L</text>`;}tx.forEach((t,i)=>{const x=pad+(i/tx.length)*(W-pad-10);const y=H-24-((amts[i]-ymin)/(ymax-ymin||1))*(H-pad-24);const fl=Math.abs(a.z[i])>3;g+=`<circle cx="${x}" cy="${y}" r="${fl?3.4:1.6}" fill="${fl?C.rose:"#cbd5e1"}" ${fl?'stroke="#000" stroke-width="0.3"':""}><title>${t.c} ${t.f} ${t.g} ₹${inr(t.a)} z=${a.z[i].toFixed(2)}</title></circle>`;});g+=`<text x="${W/2}" y="${H-4}" font-size="9" fill="#777" text-anchor="middle">transactions in order</text>`;return `<svg viewBox="0 0 ${W} ${H}">${g}</svg>`;})();
  el("anomSummary").innerHTML=`<table><tbody>
    <tr><td>Z-score flags (live, |z|&gt;3)</td><td class="num"><b>${a.flags.length}</b></td></tr>
    <tr><td>Report total (Z-score + Isolation Forest)</td><td class="num">${S.anomaly.n_total}</td></tr>
    <tr><td>— by Z-score (Python)</td><td class="num">${S.anomaly.n_zscore}</td></tr>
    <tr><td>— by Isolation Forest (Python)</td><td class="num">${S.anomaly.n_iso}</td></tr>
    <tr><td>Stock months flagged (reconciliation)</td><td class="num">${S.anomaly.recon}</td></tr>
    </tbody></table><div class="note">Isolation Forest runs in Python (see report). This page recomputes the Z-score flags live as you add records.</div>`;
  let t=`<table><thead><tr><th>Date</th><th>Company</th><th>Type</th><th>Goods</th><th>Party</th><th>Amount (₹)</th><th>Z</th></tr></thead><tbody>`;
  a.flags.slice(0,10).forEach(f=>{t+=`<tr><td>${f.d}</td><td>${f.c}</td><td>${f.f}</td><td>${f.g}</td><td>${(f.p||"").slice(0,22)}</td><td class="num">${inr(f.a)}</td><td class="num">${f.z.toFixed(2)}</td></tr>`;});
  el("tblAnom").innerHTML=a.flags.length?t+`</tbody></table>`:`<div class="note">No transactions beyond ±3σ in the current data.</div>`;
}
function renderYield(){let t=`<table><thead><tr><th>Cycle</th><th>Yield %</th><th>Variance</th><th>Z</th><th>Flag</th></tr></thead><tbody>`;S.yield_batches.forEach(b=>{t+=`<tr><td>${b.Company} ${b.Year}</td><td class="num">${b.Yield_pct.toFixed(1)}</td><td class="num">${b.Variance.toFixed(1)}</td><td class="num">${b.Zscore.toFixed(2)}</td><td>${b.Outlier==="Yes"?'<span style="color:#be123c">●</span>':''}</td></tr>`;});el("tblYield").innerHTML=t+`</tbody></table><div class="note">Full period — `+S.yield_company.map(c=>c.Company+" "+c.Recovery_pct.toFixed(1)+"%").join(" · ")+`.</div>`;}
function renderAdded(){if(!added.length){el("tblAdded").innerHTML=`<div class="note">No records added yet.</div>`;return;}let t=`<table><thead><tr><th>#</th><th>Date</th><th>Company</th><th>Type</th><th>Goods</th><th>Party</th><th>Qty</th><th>Amount (₹)</th><th>Pay</th><th></th></tr></thead><tbody>`;added.forEach((r,i)=>{t+=`<tr><td>${i+1}</td><td>${r.d}</td><td>${r.c}</td><td><span class="pill ${isSale(r)?'sale':'pur'}">${r.f}</span></td><td>${r.g}</td><td>${r.p||""}</td><td class="num">${r.q?inr(r.q):"—"}</td><td class="num">${inr(r.a)}</td><td><span class="pill ${r.pay||'Paid'}">${r.pay||'Paid'}</span></td><td><span class="x" data-i="${i}">✕</span></td></tr>`;});el("tblAdded").innerHTML=t+`</tbody></table>`;el("tblAdded").querySelectorAll(".x").forEach(b=>b.onclick=()=>{added.splice(+b.dataset.i,1);save();render();toast("Removed");});}
function save(){try{localStorage.setItem(LS,JSON.stringify(added));}catch(e){}}
function toast(m){const t=el("toast");t.textContent=m;t.classList.add("show");setTimeout(()=>t.classList.remove("show"),1800);}

// init assumptions
el("p_rcn").value=S.pnl.rcn_price;el("p_out").value=(S.pnl.outturn*100);el("p_proc").value=S.pnl.processing;el("p_price").value=S.pnl.kernel_price;
["p_rcn","p_out","p_proc","p_price"].forEach(id=>el(id).oninput=()=>renderPnL());
function autoAmt(){const q=+el("f_qty").value,r=+el("f_rate").value;if(q>0&&r>0)el("f_amt").value=Math.round(q*r);}
el("f_qty").oninput=autoAmt;el("f_rate").oninput=autoAmt;
el("addBtn").onclick=()=>{const d=el("f_date").value,amt=+el("f_amt").value;if(!d){toast("Pick a date");return;}if(!(amt>0)){toast("Enter amount (or qty & rate)");return;}added.push({c:el("f_co").value,f:el("f_flow").value,g:el("f_goods").value,d,p:el("f_party").value.trim(),a:amt,q:+el("f_qty").value||0,rt:+el("f_rate").value||0,inv:el("f_inv").value.trim(),pay:el("f_pay").value});save();["f_amt","f_qty","f_rate","f_inv","f_party"].forEach(id=>el(id).value="");render();toast("Record added — dashboard updated");};
el("resetBtn").onclick=()=>{if(added.length&&confirm("Remove all "+added.length+" added records?")){added=[];save();render();toast("Cleared");}};

// ---- EXPORT (SpreadsheetML, native) ----
function xesc(s){return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function xc(v,st){const num=typeof v==="number"&&isFinite(v);return `<Cell${st?` ss:StyleID="${st}"`:""}><Data ss:Type="${num?"Number":"String"}">${num?v:xesc(v)}</Data></Cell>`;}
function xsheet(name,rows){return `<Worksheet ss:Name="${xesc(name)}"><Table>`+rows.map((r,i)=>`<Row>${r.map(c=>xc(c,i===0?"hdr":(typeof c==="number"?"num":null))).join("")}</Row>`).join("")+`</Table></Worksheet>`;}
function dl(text,name){const b=new Blob([text],{type:"application/vnd.ms-excel"});const a=document.createElement("a");a.href=URL.createObjectURL(b);a.download=name;document.body.appendChild(a);a.click();document.body.removeChild(a);setTimeout(()=>URL.revokeObjectURL(a.href),1000);}
function exportExcel(){const m=computeSales(),p=pnl(),b=computeBuyers(),a=computeAnom();const net=m.totS-m.totP;const sh=[];
  sh.push(xsheet("Summary",[["Metric","Value"],["Generated",new Date().toLocaleString()],["Records",m.tx.length],["Original",BASE.length],["Added",added.length],["Total Sales (₹)",Math.round(m.totS)],["Total Purchase (₹)",Math.round(m.totP)],["Net (₹)",Math.round(net)],["Kernel share of sales (%)",+(m.mix.FCN/((m.mix.FCN+m.mix.RCN)||1)*100).toFixed(1)],["Buyers scored",b.rows.length],["High-risk buyers",b.tiers.High],["Anomalies (Z live)",a.flags.length]]));
  sh.push(xsheet("Annual",[["Financial Year","Sales (₹)","Purchase (₹)","Net (₹)"]].concat(m.fys.map(f=>[f,Math.round(m.A[f].s),Math.round(m.A[f].p),Math.round(m.A[f].s-m.A[f].p)]))));
  const ss=m.sStat,ps=m.pStat;sh.push(xsheet("Descriptive Stats",[["Measure","Sales (₹)","Purchase (₹)"],["Count",ss.n,ps.n],["Total",Math.round(ss.total),Math.round(ps.total)],["Mean",Math.round(ss.mean),Math.round(ps.mean)],["Median",Math.round(ss.median),Math.round(ps.median)],["SD",Math.round(ss.sd),Math.round(ps.sd)],["Min",Math.round(ss.min),Math.round(ps.min)],["Max",Math.round(ss.max),Math.round(ps.max)]]));
  sh.push(xsheet("P&L Model",[["Item","Value"],["RCN price (₹/kg)",p.rcn],["Outturn (%)",+(p.out*100).toFixed(1)],["Processing (₹/kg kernel)",p.proc],["Kernel price (₹/kg)",p.price],["RCN cost per kg kernel",Math.round(p.rcnCost)],["Total cost per kg",Math.round(p.total)],["Contribution per kg",Math.round(p.contrib)],["Margin (%)",+p.margin.toFixed(1)],["Break-even price (₹/kg)",Math.round(p.beP)],["Break-even outturn (%)",+p.beOut.toFixed(1)],["Whole contribution (₹/kg)",Math.round(p.wc)],["Broken contribution (₹/kg)",Math.round(p.bc)]]));
  const lab=p.steps.map(s=>(s*100>0?"+":"")+(s*100)+"%");sh.push(xsheet("Sensitivity",[["RCN \\ Price"].concat(lab)].concat(p.sens.map((row,i)=>[lab[i]].concat(row)))));
  sh.push(xsheet("Buyer Risk Model",[["Buyer","Revenue (₹)","Orders","Avg order (₹)","Recency (d)","Tenure (d)","Active months","Risk score","Tier","Credit limit (₹)"]].concat(b.rows.map(r=>[r.buyer,Math.round(r.rev),r.orders,Math.round(r.avg),r.recency,r.tenure,r.months,r.risk,r.tier,r.credit]))));
  sh.push(xsheet("Buyer Tiers",[["Tier","Buyers","Credit exposure (₹)"]].concat(["Low","Medium","High"].map(t=>[t,b.tiers[t],Math.round(b.exposure[t])]))));
  sh.push(xsheet("Anomalies (Z-score)",[["Date","Company","Flow","Goods","Party","Amount (₹)","Z-score"]].concat(a.flags.map(f=>[f.d,f.c,f.f,f.g,f.p||"",Math.round(f.a),+f.z.toFixed(2)]))));
  sh.push(xsheet("Yield",[["Cycle","Yield %","Variance","Z-score","Outlier"]].concat(S.yield_batches.map(y=>[y.Company+" "+y.Year,+y.Yield_pct.toFixed(1),+y.Variance.toFixed(1),+y.Zscore.toFixed(2),y.Outlier]))));
  const txr=[["Company","Type","Goods","Date","Party","Amount (₹)","Quantity (KG)","Payment","Source"]];m.tx.forEach(t=>txr.push([t.c,t.f,t.g,t.d,t.p||"",Math.round(t.a),t.q||"",t.pay||"",t.q!==undefined||t.pay!==undefined?"added":"original"]));sh.push(xsheet("Transactions",txr));
  const xml='<?xml version="1.0" encoding="UTF-8"?>\n<?mso-application progid="Excel.Sheet"?>\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Styles><Style ss:ID="hdr"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#0F766E" ss:Pattern="Solid"/></Style><Style ss:ID="num"><NumberFormat ss:Format="#,##0"/></Style></Styles>'+sh.join("")+'</Workbook>';
  dl(xml,"Cashew_Final_Report_"+new Date().toISOString().slice(0,10)+".xls");toast("Excel exported ("+sh.length+" sheets)");}
el("exportBtn").onclick=exportExcel;

render();
</script>
</body></html>'''
HTML=HTML.replace("__DATA__",data)
open("../Cashew_Dashboard_Final.html","w",encoding="utf-8").write(HTML)
print("written, size KB:",round(len(HTML.encode('utf-8'))/1024,1))
