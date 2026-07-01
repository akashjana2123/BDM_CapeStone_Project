"""
================================================================================
 BDM Capstone : Analytical Study of Ankita Cashew Processing (with Adrita)
 FINAL analysis script  |  Akash Jana  |  Roll 23f2000990
--------------------------------------------------------------------------------
 Builds the final-phase analysis on top of the mid-term work:
   - Profit & Loss / contribution-margin model (per kg, break-even, sensitivity)
   - Buyer risk scorecard + illustrative logistic-regression delay model
   - Anomaly detection (Z-score + Isolation Forest + reconciliation)
 Inputs (same folder):  "Sales Purchase 22-26 (1).xlsx", "Stock Summary.xlsx"
 Outputs: results_final.json, figures_final/figNN_*.png, cleaning_log.txt
 Requirements: pandas numpy statsmodels scikit-learn matplotlib openpyxl
================================================================================
"""
import os, re, json, datetime, warnings
import numpy as np, pandas as pd, openpyxl
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from matplotlib.colors import LinearSegmentedColormap
from sklearn.ensemble import IsolationForest
from sklearn.linear_model import LogisticRegression
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import train_test_split
from sklearn.metrics import roc_auc_score, roc_curve, accuracy_score
warnings.filterwarnings("ignore")

SALES_FILE="../../Source_and_Reference/Sales Purchase 22-26 (1).xlsx"; STOCK_FILE="../../Source_and_Reference/Stock Summary.xlsx"
FIG="figures_final"; os.makedirs(FIG, exist_ok=True)
DATA_END=pd.Timestamp("2026-02-10")
log=[]
def L(m): print(m); log.append(str(m))

# distinct FINAL palette (deliberately different from the mid-term navy/orange)
TEAL="#0f766e"; AMBER="#b45309"; SLATE="#334155"; ROSE="#be123c"; INDIGO="#4338ca"; GREEN="#15803d"; GREY="#94a3b8"
plt.rcParams.update({"font.size":11,"figure.dpi":130,"axes.grid":True,"grid.alpha":.25,"axes.axisbelow":True,
                     "axes.edgecolor":"#cbd5e1"})
LAKH=FuncFormatter(lambda x,_:f"{x/1e5:.0f}"); CR=FuncFormatter(lambda x,_:f"{x/1e7:.1f}")

# ---------- cleaning (same rules as mid-term) ----------
SHEET={"Adrita RCN Sale":("Adrita","Sale","RCN"),"Adrita FCN Sale":("Adrita","Sale","FCN"),
 "Adrita Purchase RCN":("Adrita","Purchase","RCN"),"Adrita Purchase FCN":("Adrita","Purchase","FCN"),
 "Ankita RCN Sale":("Ankita","Sale","RCN"),"Ankita FCN Sale":("Ankita","Sale","FCN"),
 "Ankita Purchase RCN":("Ankita","Purchase","RCN"),"Ankita Purchase FCN":("Ankita","Purchase","FCN"),
 "Ankita Export":("Ankita","Export","FCN")}
def clean_ledgers(path):
    wb=openpyxl.load_workbook(path,data_only=True); rows=[]; scanned=skip=0
    for sh,(c,f,g) in SHEET.items():
        ws=wb[sh]
        for r in range(1,ws.max_row+1):
            scanned+=1; d=ws.cell(r,1).value
            if not isinstance(d,(datetime.datetime,datetime.date)): skip+=1; continue
            nm=ws.cell(r,3).value
            if nm and "Closing Balance" in str(nm): continue
            deb,cred=ws.cell(r,6).value,ws.cell(r,7).value
            amt=cred if f in("Sale","Export") else deb
            if amt is None: amt=cred if cred is not None else deb
            if amt is None: skip+=1; continue
            rows.append(dict(Company=c,Flow=f,Goods=g,Date=pd.Timestamp(d),
                             Counterparty=str(nm).strip() if nm else "",Amount=float(amt)))
    df=pd.DataFrame(rows); before=len(df); df=df.drop_duplicates(); dup=before-len(df)
    df["FY"]=df.Date.apply(lambda x:x.year if x.month>=4 else x.year-1)
    df["FYlabel"]=df.FY.apply(lambda y:f"FY{str(y)[2:]}-{str(y+1)[2:]}")
    L(f"[CLEAN] scanned {scanned}, skipped {skip}, duplicates removed {dup}, kept {len(df)}")
    return df, dict(scanned=scanned,skipped=skip,duplicates=dup,kept=len(df))

MON={"April":4,"May":5,"June":6,"July":7,"August":8,"September":9,"October":10,"November":11,"December":12,"January":1,"February":2,"March":3}
def kg(s):
    if not isinstance(s,str): return None
    m=re.match(r"\s*(-?[\d,]+(?:\.\d+)?)\s*KG",s,re.I); return float(m.group(1).replace(",","")) if m else None
def gradeof(it):
    u=it.upper()
    if "RAW CASHEW" in u or "NUT IN SHELL" in u: return "RCN"
    if "WHOLE" in u: return "Wholes"
    if "BROKEN" in u: return "Broken"
    return None
def clean_stock(path):
    wb=openpyxl.load_workbook(path,data_only=True); monthly=[]; closing=[]
    for sh in wb.sheetnames:
        comp="Ankita" if "Ankita" in sh else "Adrita"; fy=int("20"+sh.split()[-1]); ws=wb[sh]
        st=None
        for r in range(1,ws.max_row+1):
            if ws.cell(r,2).value=="Quantity" and ws.cell(r,4).value=="Value": st=r+1;break
        if st:
            for r in range(st,ws.max_row+1):
                nm=ws.cell(r,1).value
                if not isinstance(nm,str): continue
                if nm.strip()=="Grand Total": break
                v=ws.cell(r,4).value
                if isinstance(v,(int,float)): closing.append(dict(Company=comp,Item=nm.strip(),Value=float(v)))
        cur=None
        for r in range(1,ws.max_row+1):
            a=ws.cell(r,1).value; nxt=ws.cell(r+1,1).value if r<ws.max_row else None
            if isinstance(a,str) and nxt=="Monthly Summary": cur=a.strip()
            if isinstance(a,str) and a.strip() in MON and cur:
                gg=gradeof(cur)
                if gg:
                    inq=kg(ws.cell(r,2).value) or 0; inval=ws.cell(r,3).value or 0; outq=kg(ws.cell(r,4).value) or 0
                    m=MON[a.strip()]; yr=fy if m>=4 else fy+1
                    monthly.append(dict(Company=comp,Date=pd.Timestamp(year=yr,month=m,day=1),Cat=gg,
                                        In_KG=inq,In_Val=float(inval) if isinstance(inval,(int,float)) else 0,Out_KG=outq))
    return pd.DataFrame(monthly), pd.DataFrame(closing)

# ============================================================
# LOAD
df, cc = clean_ledgers(SALES_FILE)
mv, cb = clean_stock(STOCK_FILE)
sale=df[df.Flow.isin(["Sale","Export"])].copy(); pur=df[df.Flow=="Purchase"].copy()
RES={}; RES["clean_counts"]=cc

# kernel & RCN totals (KG) and realised prices
kern_KG = mv[mv.Cat.isin(["Wholes","Broken"])].Out_KG.sum()
whole_KG = mv[mv.Cat=="Wholes"].Out_KG.sum(); broken_KG=mv[mv.Cat=="Broken"].Out_KG.sum()
rcn_KG = mv[mv.Cat=="RCN"].In_KG.sum()
fcn_sales = sale[sale.Goods=="FCN"].Amount.sum()
rcn=mv[(mv.Cat=="RCN")&(mv.In_KG>0)&(mv.In_Val>0)]
rcn_price = float((rcn.In_Val.sum()/rcn.In_KG.sum()))
kernel_price = float(fcn_sales/kern_KG)

# ============================================================
# 1) PROFIT & LOSS / CONTRIBUTION MARGIN MODEL
# ------------------------------------------------------------
ASSUME = dict(rcn_price=round(rcn_price,1), kernel_price=round(kernel_price,1),
              outturn=0.22, processing_per_kg_kernel=30.0)
whole_share = whole_KG/(whole_KG+broken_KG)
# grade prices chosen so volume-weighted average == realised kernel price, using an industry whole:broken ratio 1.71
ratio=1.71
broken_price = kernel_price/(whole_share*ratio + (1-whole_share))
whole_price = broken_price*ratio
def pnl_per_kg_kernel(price, rcnp, outturn, proc):
    rcn_cost = rcnp/outturn
    total_cost = rcn_cost + proc
    return dict(price=price, rcn_cost=rcn_cost, proc=proc, total_cost=total_cost,
                contribution=price-total_cost, margin_pct=100*(price-total_cost)/price)
base = pnl_per_kg_kernel(kernel_price, rcn_price, ASSUME["outturn"], ASSUME["processing_per_kg_kernel"])
grade_pnl = {
 "Whole": pnl_per_kg_kernel(round(whole_price,0), rcn_price, ASSUME["outturn"], ASSUME["processing_per_kg_kernel"]),
 "Broken": pnl_per_kg_kernel(round(broken_price,0), rcn_price, ASSUME["outturn"], ASSUME["processing_per_kg_kernel"]),
}
# break-even
be_price = base["total_cost"]
be_outturn = rcn_price/(kernel_price-ASSUME["processing_per_kg_kernel"])
# sensitivity grid: rows = RCN price -10..+10%, cols = kernel price -10..+10% -> contribution/kg kernel
steps=[-0.10,-0.05,0,0.05,0.10]
sens=[]
for dr in steps:
    row=[]
    for dp in steps:
        c=pnl_per_kg_kernel(kernel_price*(1+dp), rcn_price*(1+dr), ASSUME["outturn"], ASSUME["processing_per_kg_kernel"])["contribution"]
        row.append(round(c,1))
    sens.append(row)
# annual contribution estimate (per FY) from kernel KG sold and base contribution/kg
kern_by_fy={}
mvk=mv[mv.Cat.isin(["Wholes","Broken"])].copy(); mvk["FY"]=mvk.Date.apply(lambda d:d.year if d.month>=4 else d.year-1)
for fy,g in mvk.groupby("FY"):
    lbl=f"FY{str(fy)[2:]}-{str(fy+1)[2:]}"; kern_by_fy[lbl]=float(g.Out_KG.sum())
annual_contrib={k:round(v*base["contribution"]) for k,v in kern_by_fy.items()}
RES["pnl"]=dict(assume=ASSUME, base=base, grade=grade_pnl, whole_price=round(whole_price), broken_price=round(broken_price),
                whole_share=round(whole_share*100,1), be_price=round(be_price,1), be_outturn=round(be_outturn*100,1),
                sens=sens, sens_steps=steps, kernel_price=round(kernel_price,1), rcn_price=round(rcn_price,1),
                annual_contrib=annual_contrib, total_kernel_KG=round(kern_KG))
L(f"[P&L] kernel price Rs{kernel_price:.0f}/kg, cost Rs{base['total_cost']:.0f}/kg, contribution Rs{base['contribution']:.0f}/kg ({base['margin_pct']:.1f}%)")
L(f"[P&L] break-even price Rs{be_price:.0f}/kg, break-even outturn {be_outturn*100:.1f}%")

# ============================================================
# 2) BUYER RISK SCORECARD + illustrative logistic regression
# ------------------------------------------------------------
S=sale.copy()
bad={"Cash","Advance From Customers",""}
B=[]
for name,g in S.groupby("Counterparty"):
    if name in bad: continue
    rev=g.Amount.sum(); orders=len(g); first=g.Date.min(); last=g.Date.max()
    recency=(DATA_END-last).days; tenure=max(1,(last-first).days)
    months=g.Date.dt.to_period("M").nunique()
    freq=orders/max(1,months)
    B.append(dict(buyer=name,revenue=rev,orders=orders,avg_order=rev/orders,
                  recency=recency,tenure=tenure,active_months=months,freq=freq))
bd=pd.DataFrame(B)
def nrm(x,inv=False):
    x=np.asarray(x,float); lo,hi=np.nanmin(x),np.nanmax(x)
    z=(x-lo)/(hi-lo) if hi>lo else np.zeros_like(x); return 1-z if inv else z
# risk: higher recency, lower freq, lower tenure, higher single-exposure = riskier
bd["risk"]= (0.35*nrm(bd.recency)+0.25*nrm(bd.freq,inv=True)+0.20*nrm(bd.tenure,inv=True)+0.20*nrm(bd.avg_order))*100
bd["risk"]=bd.risk.round(1)
bd["tier"]=pd.cut(bd.risk,[-1,33,66,1000],labels=["Low","Medium","High"])
# credit limit: ~1 month of business scaled by reliability (Low x2.0, Med x1.0, High x0.4)
mult={"Low":2.0,"Medium":1.0,"High":0.4}
bd["monthly"]=bd.revenue/np.maximum(1,bd.active_months)
bd["credit_limit"]=(bd.monthly*bd.tier.astype(str).map(mult)).round(-3)
bd=bd.sort_values("revenue",ascending=False).reset_index(drop=True)

# illustrative logistic regression: SIMULATE a payment-delay outcome from features + noise,
# then fit LR to show the method works (real payment data is not in the source files).
rng=np.random.default_rng(42)
X=bd[["revenue","orders","avg_order","recency","freq","tenure"]].copy()
Xs=StandardScaler().fit_transform(X)
true_coef=np.array([-0.7,-0.6,0.5,0.9,-0.5,-0.4])  # plausible drivers of delay
logit=Xs@true_coef + rng.normal(0,0.8,len(bd))
p=1/(1+np.exp(-logit)); y=(p>0.5).astype(int)
if y.sum()<3 or y.sum()>len(y)-3:  # ensure both classes
    y=(p>np.median(p)).astype(int)
Xtr,Xte,ytr,yte=train_test_split(Xs,y,test_size=0.3,random_state=1,stratify=y)
lr=LogisticRegression(max_iter=500).fit(Xtr,ytr)
prob=lr.predict_proba(Xte)[:,1]; pred=lr.predict(Xte)
auc=roc_auc_score(yte,prob); acc=accuracy_score(yte,pred)
fpr,tpr,_=roc_curve(yte,prob)
coef=dict(zip(["revenue","orders","avg_order","recency","freq","tenure"],lr.coef_[0].round(3)))
RES["buyer"]=dict(n_buyers=len(bd),
    tiers={t:int((bd.tier==t).sum()) for t in ["Low","Medium","High"]},
    exposure={t:float(bd.loc[bd.tier==t,"credit_limit"].sum()) for t in ["Low","Medium","High"]},
    top=bd.head(15)[["buyer","revenue","orders","avg_order","recency","risk","tier","credit_limit"]].to_dict("records"),
    lr=dict(auc=round(float(auc),3),acc=round(float(acc),3),coef=coef,n_delay=int(y.sum()),n_total=int(len(y))),
    roc={"fpr":list(map(float,fpr)),"tpr":list(map(float,tpr))})
L(f"[BUYER] {len(bd)} buyers; tiers {RES['buyer']['tiers']}; illustrative LR AUC={auc:.2f} acc={acc:.2f}")

# ============================================================
# 3) ANOMALY DETECTION (Z-score + Isolation Forest + reconciliation)
# ------------------------------------------------------------
A=df.copy()
A["z"]=A.groupby(["Company","Flow","Goods"]).Amount.transform(lambda s:(s-s.mean())/(s.std(ddof=0) if s.std(ddof=0)>0 else 1))
A["z_flag"]=A.z.abs()>3
# Isolation Forest on engineered features
A["logamt"]=np.log1p(A.Amount); A["m"]=A.Date.dt.month
A["fl"]=(A.Flow=="Purchase").astype(int); A["gd"]=(A.Goods=="RCN").astype(int); A["co"]=(A.Company=="Ankita").astype(int)
feat=A[["logamt","m","fl","gd","co"]].values
iso=IsolationForest(contamination=0.02,random_state=0).fit(feat)
A["iso_score"]=-iso.score_samples(feat)  # higher = more anomalous
A["iso_flag"]=iso.predict(feat)==-1
A["anomaly"]=A.z_flag | A.iso_flag
n_z=int(A.z_flag.sum()); n_iso=int(A.iso_flag.sum()); n_any=int(A.anomaly.sum())
top_anom=A[A.anomaly].sort_values("iso_score",ascending=False).head(12)
top_list=[dict(Company=r.Company,Flow=r.Flow,Goods=r.Goods,Date=r.Date.strftime("%Y-%m-%d"),
               Party=r.Counterparty[:24],Amount=round(r.Amount),Z=round(r.z,2),
               Reason=("Z-score" if r.z_flag and not r.iso_flag else ("Isolation Forest" if r.iso_flag and not r.z_flag else "both")))
          for r in top_anom.itertuples()]
# reconciliation: monthly RCN-in vs kernel-out ratio outliers
rin=mv[mv.Cat=="RCN"].groupby("Date").In_KG.sum(); kout=mv[mv.Cat.isin(["Wholes","Broken"])].groupby("Date").Out_KG.sum()
recon=pd.DataFrame({"rin":rin,"kout":kout}).fillna(0)
recon=recon[(recon.rin>0)]; recon["ratio"]=recon.kout/recon.rin
rz=(recon.ratio-recon.ratio.mean())/recon.ratio.std()
recon_flags=int((rz.abs()>2).sum())
RES["anomaly"]=dict(n_zscore=n_z,n_iso=n_iso,n_total=n_any,pct=round(100*n_any/len(A),1),
                    recon_months_flagged=recon_flags,top=top_list)
L(f"[ANOMALY] z-flag {n_z}, iso-flag {n_iso}, total {n_any} ({100*n_any/len(A):.1f}%), recon months {recon_flags}")

# carry-forward recap numbers (computed fresh)
RES["totals"]=dict(total_sales=float(sale.Amount.sum()),total_purchase=float(pur.Amount.sum()),
                   n_sales=int(len(sale)),n_purchase=int(len(pur)))
RES["sales_by_goods"]={k:float(v) for k,v in sale.groupby("Goods").Amount.sum().items()}
RES["rcn_price"]=round(rcn_price,1); RES["kernel_price"]=round(kernel_price,1)

# descriptive stats for all variables (final list incl new buyer vars)
def desc(s):
    s=pd.Series(s).dropna(); 
    return dict(count=int(s.count()),mean=float(s.mean()),median=float(s.median()),std=float(s.std()),min=float(s.min()),max=float(s.max()),total=float(s.sum()))
RES["descriptive"]={
 "Sale amount (₹)":desc(sale.Amount),"Purchase amount (₹)":desc(pur.Amount),
 "RCN price (₹/KG)":desc(rcn.In_Val/rcn.In_KG),
 "Buyer revenue (₹)":desc(bd.revenue),"Buyer orders (count)":desc(bd.orders),
 "Buyer risk score (0-100)":desc(bd.risk),
 "Contribution per kg kernel (₹)":desc([grade_pnl["Whole"]["contribution"],grade_pnl["Broken"]["contribution"],base["contribution"]]),
}
json.dump(RES,open("results_final.json","w"),indent=2,default=str)
bd.to_pickle("buyers.pkl"); A.to_pickle("anomdf.pkl"); recon.to_pickle("recon.pkl")
open("cleaning_log.txt","a").write("\n[FINAL] "+ " | ".join(log))
L("[DONE] results_final.json written")

# ============================================================
# FINAL FIGURES  (fresh set, distinct palette; NOT reused from mid-term)
# ============================================================
def _save(n): plt.tight_layout(); plt.savefig(f"{FIG}/{n}"); plt.close()

# Recap data
df["FYlabel"]=df.FY.apply(lambda y:f"FY{str(y)[2:]}-{str(y+1)[2:]}")
ann={}
for fy,g in df.groupby("FYlabel"):
    ann[fy]=dict(sales=g[g.Flow.isin(["Sale","Export"])].Amount.sum(),pur=g[g.Flow=="Purchase"].Amount.sum())
fys=sorted(ann)

# Figure 1 — Sales, purchase and net by year (lollipop + net line; distinct design)
fig,ax=plt.subplots(figsize=(8,4.2))
x=np.arange(len(fys))
s=[ann[f]["sales"] for f in fys]; p=[ann[f]["pur"] for f in fys]; net=[a-b for a,b in zip(s,p)]
ax.hlines(x-0.12,0,s,color=TEAL,lw=7,alpha=.85,label="Sales")
ax.hlines(x+0.12,0,p,color=AMBER,lw=7,alpha=.85,label="Purchase")
ax.set_yticks(x); ax.set_yticklabels(fys); ax.xaxis.set_major_formatter(CR); ax.set_xlabel("₹ Crore")
ax.set_title("Figure 1: Sales and Purchase by Year"); ax.legend(loc="lower right")
ax.invert_yaxis(); _save("fig01_overview.png")

# Figure 2 — Forecast with validation (one consolidated panel)
import statsmodels.api as sm
from statsmodels.tsa.statespace.sarimax import SARIMAX
ts=sale.groupby(sale.Date.dt.to_period("M")).Amount.sum(); ts.index=ts.index.to_timestamp(); ts=ts.asfreq("MS").fillna(0); ts=ts[ts.index<="2026-01-01"]
n=len(ts); h=6; tr,te=ts.iloc[:n-h],ts.iloc[n-h:]
m=SARIMAX(tr,order=(1,1,1),seasonal_order=(0,1,1,12),enforce_stationarity=False,enforce_invertibility=False).fit(disp=False)
val=m.forecast(h)
mf=SARIMAX(ts,order=(1,1,1),seasonal_order=(0,1,1,12),enforce_stationarity=False,enforce_invertibility=False).fit(disp=False)
fc=mf.get_forecast(4); fm=fc.predicted_mean; ci=fc.conf_int(alpha=.2)
fig,ax=plt.subplots(figsize=(9.5,4.2))
ax.plot(ts.index,ts.values,color=SLATE,lw=1.5,label="Actual")
ax.axvspan(te.index[0],te.index[-1],color=TEAL,alpha=.07)
ax.plot(te.index,val.values,color=ROSE,lw=2,ls="--",marker="s",ms=4,label="Validation (held-out)")
ax.plot(fm.index,fm.values,color=TEAL,lw=2,marker="D",ms=4,label="Forecast (next 4 mo)")
ax.fill_between(fm.index,np.maximum(0,ci.iloc[:,0]),ci.iloc[:,1],color=TEAL,alpha=.15)
ax.yaxis.set_major_formatter(LAKH); ax.set_ylabel("Sales (₹ Lakh)")
ax.set_title("Figure 2: Sales Forecast — validation window shaded, then forecast")
ax.legend(fontsize=8,ncol=3); plt.xticks(rotation=45,fontsize=8); _save("fig02_forecast.png")

# Figure 3 — Seasonality heatmap (year x month)
piv=sale.copy(); piv["yr"]=piv.Date.dt.year; piv["mo"]=piv.Date.dt.month
H=piv.pivot_table(index="yr",columns="mo",values="Amount",aggfunc="sum").reindex(columns=range(1,13))
fig,ax=plt.subplots(figsize=(9,3.6))
cmap=LinearSegmentedColormap.from_list("t",["#f0fdfa",TEAL,"#0b3b36"])
im=ax.imshow(H.values/1e5,aspect="auto",cmap=cmap)
ax.set_xticks(range(12)); ax.set_xticklabels(["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],fontsize=9)
ax.set_yticks(range(len(H.index))); ax.set_yticklabels(H.index)
for i in range(H.shape[0]):
    for j in range(12):
        v=H.values[i,j]
        if not np.isnan(v): ax.text(j,i,f"{v/1e5:.0f}",ha="center",va="center",fontsize=7,color="#333" if v/1e5<60 else "#fff")
ax.set_title("Figure 3: Sales Seasonality Heatmap (₹ lakh by year × month)")
cb=fig.colorbar(im,ax=ax,fraction=.025); cb.ax.tick_params(labelsize=7); _save("fig03_seasonality_heat.png")

# Figure 4 — Yield control chart
yb=json.load(open("results.json"))["yield_batches"] if os.path.exists("results.json") else []
if yb:
    lbl=[b["Company"][:2]+"-"+b["Year"][2:] for b in yb]; yv=[b["Yield_pct"] for b in yb]
    mu=np.mean(yv); sd=np.std(yv,ddof=1)
    fig,ax=plt.subplots(figsize=(9,4))
    ax.plot(lbl,yv,marker="o",color=TEAL,lw=1.8,zorder=3)
    ax.axhline(mu,color=SLATE,lw=1.2,label=f"mean {mu:.1f}%")
    ax.axhline(mu+1.5*sd,color=ROSE,ls="--",lw=1,label="±1.5σ control")
    ax.axhline(mu-1.5*sd,color=ROSE,ls="--",lw=1)
    ax.axhline(22,color=AMBER,ls=":",lw=1.5,label="target 22%")
    for i,b in enumerate(yb):
        if b["Outlier"]=="Yes": ax.scatter([i],[b["Yield_pct"]],s=120,facecolors="none",edgecolors=ROSE,lw=2,zorder=4)
    ax.set_ylabel("Yield %"); ax.set_title("Figure 4: Yield Control Chart (outliers ringed)")
    ax.legend(fontsize=8); plt.xticks(rotation=45,fontsize=8); _save("fig04_yield_control.png")

# Figure 5 — Contribution waterfall (per kg kernel)
b=base
steps_w=[("Selling price",b["price"],TEAL),("− RCN cost",-b["rcn_cost"],AMBER),("− Processing",-b["proc"],AMBER),("= Contribution",None,GREEN)]
fig,ax=plt.subplots(figsize=(8,4.2)); cum=0
for i,(lab,val,col) in enumerate(steps_w):
    if val is None:
        ax.bar(i,b["contribution"],color=GREEN); ax.text(i,b["contribution"]+5,f"₹{b['contribution']:.0f}",ha="center",fontsize=9,fontweight="bold")
    else:
        start=cum; ax.bar(i,val,bottom=start,color=col)
        ax.text(i,start+val+ (8 if val>0 else -16),("+" if val>0 else "")+f"₹{val:.0f}",ha="center",fontsize=8)
        cum+=val
ax.set_xticks(range(4)); ax.set_xticklabels([s[0] for s in steps_w],fontsize=9)
ax.set_ylabel("₹ per kg kernel"); ax.set_title("Figure 5: Profit Build-up per kg of Kernel")
_save("fig05_waterfall.png")

# Figure 6 — Break-even chart
prices=np.linspace(450,950,60); contrib=prices-(b["rcn_cost"]+b["proc"])
fig,ax=plt.subplots(figsize=(8,4))
ax.plot(prices,contrib,color=TEAL,lw=2)
ax.axhline(0,color=SLATE,lw=1)
ax.axvline(b["total_cost"],color=ROSE,ls="--",lw=1.2,label=f"break-even ₹{b['total_cost']:.0f}/kg")
ax.axvline(b["price"],color=AMBER,ls=":",lw=1.5,label=f"current ₹{b['price']:.0f}/kg")
ax.fill_between(prices,contrib,0,where=(contrib>=0),color=GREEN,alpha=.12)
ax.fill_between(prices,contrib,0,where=(contrib<0),color=ROSE,alpha=.12)
ax.set_xlabel("Kernel selling price (₹/kg)"); ax.set_ylabel("Contribution (₹/kg)")
ax.set_title("Figure 6: Break-even on Kernel Selling Price"); ax.legend(fontsize=8); _save("fig06_breakeven.png")

# Figure 7 — Sensitivity heatmap
sens=np.array(RES["pnl"]["sens"]); steps=RES["pnl"]["sens_steps"]
labs=[f"{int(s*100):+d}%" for s in steps]
fig,ax=plt.subplots(figsize=(7,5))
cmap2=LinearSegmentedColormap.from_list("rg",[ROSE,"#fff7ed",GREEN])
im=ax.imshow(sens,cmap=cmap2,aspect="auto",vmin=-50,vmax=250)
ax.set_xticks(range(5)); ax.set_xticklabels(labs); ax.set_yticks(range(5)); ax.set_yticklabels(labs)
ax.set_xlabel("Kernel price change"); ax.set_ylabel("RCN price change")
for i in range(5):
    for j in range(5): ax.text(j,i,f"{sens[i,j]:.0f}",ha="center",va="center",fontsize=8,color="#222")
ax.set_title("Figure 7: Sensitivity of Contribution (₹/kg kernel)")
fig.colorbar(im,ax=ax,fraction=.045); _save("fig07_sensitivity.png")

# Figure 8 — Per-grade contribution
g=RES["pnl"]["grade"]
fig,ax=plt.subplots(figsize=(6.5,4))
names=["Whole kernel","Broken kernel"]; vals=[g["Whole"]["contribution"],g["Broken"]["contribution"]]
cols=[GREEN if v>=0 else ROSE for v in vals]
ax.bar(names,vals,color=cols,width=.55)
for i,v in enumerate(vals): ax.text(i,v+(6 if v>=0 else -14),f"₹{v:.0f}/kg",ha="center",fontsize=10,fontweight="bold")
ax.axhline(0,color=SLATE,lw=1); ax.set_ylabel("Contribution (₹/kg)")
ax.set_title("Figure 8: Contribution by Grade"); _save("fig08_grade.png")

# Figure 9 — Buyer risk scatter
bd2=bd.copy()
tcol={"Low":GREEN,"Medium":AMBER,"High":ROSE}
fig,ax=plt.subplots(figsize=(8.5,4.6))
for t in ["Low","Medium","High"]:
    s=bd2[bd2.tier==t]
    ax.scatter(s.recency,s.revenue/1e7,s=np.clip(s.orders*6,20,300),alpha=.6,color=tcol[t],label=t,edgecolors="white",lw=.5)
ax.set_xlabel("Days since last order (recency)"); ax.set_ylabel("Revenue (₹ crore)")
ax.set_title("Figure 9: Buyer Risk Map (bubble size = number of orders)")
ax.legend(title="Risk tier",fontsize=8); _save("fig09_buyer_map.png")

# Figure 10 — tier counts + credit exposure
tiers=["Low","Medium","High"]; cnt=[RES["buyer"]["tiers"][t] for t in tiers]; exp=[RES["buyer"]["exposure"][t] for t in tiers]
fig,ax=plt.subplots(figsize=(7.5,4))
ax2=ax.twinx()
ax.bar(np.arange(3)-.18,cnt,.36,color=SLATE,label="Buyers (count)")
ax2.bar(np.arange(3)+.18,[e/1e7 for e in exp],.36,color=TEAL,label="Credit exposure (₹ cr)")
ax.set_xticks(range(3)); ax.set_xticklabels(tiers); ax.set_ylabel("Buyers (count)"); ax2.set_ylabel("Credit exposure (₹ cr)"); ax2.grid(False)
ax.set_title("Figure 10: Buyer Tiers and Recommended Credit Exposure")
l1,la1=ax.get_legend_handles_labels(); l2,la2=ax2.get_legend_handles_labels(); ax.legend(l1+l2,la1+la2,fontsize=8)
_save("fig10_tiers.png")

# Figure 11 — ROC curve (illustrative LR)
roc=RES["buyer"]["roc"]; auc=RES["buyer"]["lr"]["auc"]
fig,ax=plt.subplots(figsize=(5.6,5))
ax.plot(roc["fpr"],roc["tpr"],color=TEAL,lw=2.2,label=f"LR (AUC={auc})")
ax.plot([0,1],[0,1],color=GREY,ls="--",lw=1)
ax.set_xlabel("False positive rate"); ax.set_ylabel("True positive rate")
ax.set_title("Figure 11: ROC — Illustrative Payment-delay Model"); ax.legend(fontsize=9,loc="lower right"); _save("fig11_roc.png")

# Figure 12 — LR coefficients
co=RES["buyer"]["lr"]["coef"]; ks=list(co.keys()); vs=[co[k] for k in ks]
fig,ax=plt.subplots(figsize=(7.5,4))
order=np.argsort(vs); ks=[ks[i] for i in order]; vs=[vs[i] for i in order]
ax.barh(ks,vs,color=[ROSE if v>0 else GREEN for v in vs])
ax.axvline(0,color=SLATE,lw=1); ax.set_xlabel("Coefficient (→ higher delay risk)")
ax.set_title("Figure 12: Drivers of Payment-delay Risk (illustrative)"); _save("fig12_coef.png")

# Figure 13 — anomaly scatter (date vs amount, anomalies highlighted)
fig,ax=plt.subplots(figsize=(9.5,4.2))
norm=A[~A.anomaly]; an=A[A.anomaly]
ax.scatter(norm.Date,norm.Amount,s=10,color=GREY,alpha=.5,label="normal")
ax.scatter(an.Date,an.Amount,s=42,color=ROSE,edgecolors="black",lw=.4,label="anomaly",zorder=3)
ax.set_yscale("log"); ax.set_ylabel("Amount (₹, log scale)")
ax.set_title("Figure 13: Transactions with Anomalies Flagged"); ax.legend(fontsize=8); plt.xticks(rotation=30,fontsize=8); _save("fig13_anom_scatter.png")

# Figure 14 — isolation forest score histogram with flagged
fig,ax=plt.subplots(figsize=(8,4))
ax.hist(A.iso_score,bins=40,color=TEAL,alpha=.85)
thr=A[A.iso_flag].iso_score.min()
ax.axvline(thr,color=ROSE,ls="--",lw=1.4,label=f"flag threshold ({RES['anomaly']['n_iso']} flagged)")
ax.set_xlabel("Isolation Forest anomaly score (higher = more unusual)"); ax.set_ylabel("Transactions")
ax.set_title("Figure 14: Anomaly-score Distribution"); ax.legend(fontsize=8); _save("fig14_iso_hist.png")

# Figure 15 — reconciliation: monthly kernel-out / RCN-in ratio
rr=recon.copy()
mu=rr.ratio.mean(); sd=rr.ratio.std()
fig,ax=plt.subplots(figsize=(9.5,4))
ax.plot(rr.index,rr.ratio,marker="o",ms=3,color=SLATE,lw=1.2)
ax.axhline(mu,color=TEAL,lw=1.2,label=f"mean {mu:.2f}")
ax.axhline(mu+2*sd,color=ROSE,ls="--",lw=1,label="±2σ"); ax.axhline(max(0,mu-2*sd),color=ROSE,ls="--",lw=1)
fl=rr[(rr.ratio-mu).abs()>2*sd]
ax.scatter(fl.index,fl.ratio,s=90,facecolors="none",edgecolors=ROSE,lw=2,zorder=4)
ax.set_ylabel("Kernel-out ÷ RCN-in (monthly)"); ax.set_title("Figure 15: Stock Reconciliation — odd months ringed")
ax.legend(fontsize=8); plt.xticks(rotation=30,fontsize=8); _save("fig15_recon.png")

print("[FIGURES] final figures written to ./"+FIG+"/")
import os as _os; print(sorted([f for f in _os.listdir(FIG) if f.endswith('.png')]))
