const fs=require('fs');
const D=require('docx');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,ImageRun,Footer,
AlignmentType,LevelFormat,HeadingLevel,BorderStyle,WidthType,ShadingType,
VerticalAlign,PageNumber,PageBreak,TabStopType,LeaderType,ExternalHyperlink}=D;
const o=JSON.parse(fs.readFileSync('results.json'));const fc=o.forecast;
const pages=fs.existsSync('toc_pages.json')?JSON.parse(fs.readFileSync('toc_pages.json')):{};
const links=fs.existsSync('links.json')?JSON.parse(fs.readFileSync('links.json')):{};
const NAVY="1F4E79",ORANGE="C55A11",HDR="1F4E79";
const inr=n=>Math.round(n).toLocaleString('en-IN');
const cr=n=>(n/1e7).toFixed(2), lk=n=>(n/1e5).toFixed(1);
const F="figures/";
// ---------- helpers ----------
const H1=t=>new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun(t)]});
const H2=t=>new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun(t)]});
const P=(t,j=true)=>new Paragraph({spacing:{after:120,line:284},alignment:j?AlignmentType.JUSTIFIED:AlignmentType.LEFT,children:[new TextRun({text:t,size:22})]});
const bul=t=>new Paragraph({numbering:{reference:"bul",level:0},spacing:{after:80},children:[new TextRun({text:t,size:22})]});
const eq=t=>new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:60,after:120},children:[new TextRun({text:t,bold:true,size:24})]});
const figcap=(n,t)=>new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:200},children:[new TextRun({text:`Figure ${n}: ${t}`,italics:true,size:18,color:"555555"})]});
const tabcap=(n,t)=>new Paragraph({spacing:{before:120,after:40},children:[new TextRun({text:`Table ${n}: ${t}`,bold:true,size:20,color:"333333"})]});
const fig=(file,w,h,n,cap)=>[new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:60},children:[new ImageRun({type:"png",data:fs.readFileSync(F+file),transformation:{width:w,height:h},altText:{title:cap,description:cap,name:cap}})]}),figcap(n,cap)];
const bd={style:BorderStyle.SINGLE,size:1,color:"BBBBBB"};const borders={top:bd,bottom:bd,left:bd,right:bd};
function table(headers,rows,widths){const total=widths.reduce((a,b)=>a+b,0);
 const mk=(txt,opt={})=>new TableCell({borders,width:{size:opt.w,type:WidthType.DXA},shading:{fill:opt.fill||"FFFFFF",type:ShadingType.CLEAR},margins:{top:50,bottom:50,left:90,right:90},verticalAlign:VerticalAlign.CENTER,children:[new Paragraph({alignment:opt.al||AlignmentType.LEFT,children:[new TextRun({text:String(txt),bold:opt.b||false,color:opt.color||"000000",size:opt.sz||19})]})]});
 const hr=new TableRow({tableHeader:true,children:headers.map((h,i)=>mk(h,{w:widths[i],b:true,color:"FFFFFF",fill:HDR,al:i==0?AlignmentType.LEFT:AlignmentType.CENTER}))});
 const drs=rows.map((r,ri)=>new TableRow({children:r.map((c,i)=>mk(c,{w:widths[i],al:i==0?AlignmentType.LEFT:AlignmentType.CENTER,fill:ri%2?"F2F6FC":"FFFFFF"}))}));
 return new Table({width:{size:total,type:WidthType.DXA},columnWidths:widths,rows:[hr,...drs]});}
// TOC line with dot leader + page number
function toc(label,key,sub){const pg=pages[key]!==undefined?String(pages[key]):"";
 return new Paragraph({tabStops:[{type:TabStopType.RIGHT,position:9020,leader:LeaderType.DOT}],spacing:{after:60},
  indent:sub?{left:360}:undefined,children:[new TextRun({text:label,size:22,bold:!sub}),new TextRun({text:"\t"+pg,size:22})]});}

// shorthand numbers
const asp=Object.fromEntries(o.company_flow?[]:[]); // placeholder
const D2=o.descriptive;
const sa=D2["Sale amount (₹)"],pu=D2["Purchase amount (₹)"];
const fcMix=(o.sales_by_goods.FCN/(o.sales_by_goods.FCN+o.sales_by_goods.RCN)*100).toFixed(1);
const ankY=o.yield_company.find(x=>x.Company=='Ankita'),adrY=o.yield_company.find(x=>x.Company=='Adrita');
const rcnP=o.rcn_price_yearly;const cc=o.clean_counts;
// annual sales/purchase from company_flow
const annual={};
// rebuild annual via fy not available; compute from totals table instead -> read from a precomputed file
const ann=JSON.parse(fs.readFileSync('annual.json'));

let body=[];
// ===== TITLE =====
body.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:1100,after:120},children:[new TextRun({text:"Analytical Study of Ankita Cashew Processing",bold:true,size:36,color:NAVY})]}));
body.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:600},children:[new TextRun({text:"Mid-term Report for the BDM Capstone Project",bold:true,size:26})]}));
["Submitted by","Name: Akash Jana","Roll Number: 23f2000990"].forEach(t=>body.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:80},children:[new TextRun({text:t,size:24})]})));
body.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:160,after:160},children:[new ImageRun({type:"png",data:fs.readFileSync("iitm_logo.png"),transformation:{width:220,height:220},altText:{title:"IITM",description:"IITM logo",name:"IITM"}})]}));
["IITM Online BS Degree Program,","Indian Institute of Technology Madras, Chennai","Tamil Nadu, India, 600036"].forEach(t=>body.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:60},children:[new TextRun({text:t,size:22})]})));
body.push(new Paragraph({children:[new PageBreak()]}));

// ===== CONTENTS (manual, with page numbers) =====
body.push(H1("Contents"));
body.push(toc("Declaration Statement","declaration"));
body.push(toc("1. Executive Summary and Title","s1"));
body.push(toc("2. Proof of Originality of the Data","s2"));
body.push(toc("3. Metadata and Data Collection","s3"));
body.push(toc("4. Descriptive Statistics","s4"));
body.push(toc("5. Methods (How the Analysis Was Done)","s5"));
body.push(toc("5.1 Data Cleaning and Preprocessing","s51",true));
body.push(toc("5.2 Demand Forecasting","s52",true));
body.push(toc("5.3 Inventory Yield Analysis","s53",true));
body.push(toc("6. Results and Findings","s6"));
body.push(toc("6.1 Business Overview","s61",true));
body.push(toc("6.2 Demand and Price Forecasting","s62",true));
body.push(toc("6.3 Inventory and Yield","s63",true));
body.push(toc("6.4 Revenue Mix and Buyers","s64",true));
body.push(toc("6.5 Main Findings","s65",true));
body.push(new Paragraph({children:[new PageBreak()]}));

// ===== DECLARATION =====
body.push(H1("Declaration Statement"));
[`I am working on a project titled "Analytical Study of Ankita Cashew Processing". I thank Ankita Cashew Processing and its sister unit Adrita Cashew Processing for giving me the data and help I needed for this project.`,
`The data in this report is real and correct to the best of my knowledge. It was taken from the business itself (primary source) and checked with care.`,
`I have explained how I collected and studied the data in this report. The results shown here are a true picture of what the data says.`,
`I did this project on my own. I did not copy or work jointly with anyone else. If any copied work is found at any stage, I accept that the authority can take action against me.`,
`The suggestions in this report are only for this business and this project. They cannot be used elsewhere with an IIT Madras tag. IIT Madras does not endorse this.`].forEach(t=>body.push(P(t)));
body.push(new Paragraph({spacing:{before:200},children:[new TextRun({text:"Signature of Candidate: (Digital Signature)",size:22})]}));
body.push(new Paragraph({children:[new TextRun({text:"Name: Akash Jana",size:22})]}));
body.push(new Paragraph({children:[new TextRun({text:"Date: 07 April 2026",size:22})]}));
body.push(new Paragraph({children:[new PageBreak()]}));

// ===== 1 EXEC SUMMARY =====
body.push(H1("1. Executive Summary and Title"));
body.push(P(`Ankita Cashew Processing is a cashew business in West Bengal, India. It buys raw cashew nuts (RCN), processes them into kernels (the eatable cashew, called FCN here), and sells to other traders. Its sister unit, Adrita Cashew Processing, does the same work. This report studies both units together. The work is the mid-term stage of the project, covering data cleaning, demand forecasting, and inventory yield, as set out in the proposal.`));
// 2nd paragraph: proof highlights
body.push(P(`The data is real and is backed by clear proof. The owner, Usha Jana, gave a signed No Objection Certificate (NOC) on the company letterhead, dated 08 June 2026. A short video talk with the owner and the signed letter are also kept in the project portal (links in Section 2). On top of this, the raw Tally accounting files for sales, purchase and stock were shared directly by the business. These items together show that the data comes from the business itself.`));
// dataset descriptive stats in summary
body.push(P(`The cleaned data has ${o.totals.n_sales+o.totals.n_purchase} entries from April 2022 to February 2026. A single sale is about ₹${inr(sa.mean)} on average, with a middle value (median) of ₹${inr(sa.median)} and a standard deviation (SD) of ₹${inr(sa.std)}. The smallest sale is ₹${inr(sa.min)} and the largest is ₹${inr(sa.max)}. Over four years the two units sold goods worth about ₹${cr(o.totals.total_sales)} crore and bought goods worth about ₹${cr(o.totals.total_purchase)} crore. Kernels bring in most of the money, about ${fcMix} percent of sales.`));
// list of methods before results
body.push(P(`The following methods are used in this report, grouped by problem:`));
[`Data cleaning and preprocessing — reading the raw Tally sheets, dropping non-data rows, removing duplicates, and tagging each entry.`,
`Demand forecasting — splitting sales into trend and season (decomposition), a straight-line (linear) trend as a base model, and an ARIMA model for a short forecast. Accuracy is checked with RMSE and MAE on a 6-month test.`,
`Inventory yield analysis — yield per cycle (kernel out ÷ raw nut in), a variance check against a 22% target, Z-score to flag odd cycles, and ABC classification of stock by value.`].forEach(t=>body.push(bul(t)));
body.push(P(`Main results: sales follow a season (highest around January and in May–June); the raw nut price has risen from about ₹${rcnP['2023']} to ₹${rcnP['2026']} per kg; yield is about ${adrY.Recovery_pct}% for Adrita and ${ankY.Recovery_pct}% for Ankita; and a few buyers give most of the sales. Tools used: Microsoft Excel for tables and charts, and Python for cleaning, forecasting and yield work. All money is in ₹ and all weights in KG.`));
body.push(new Paragraph({children:[new PageBreak()]}));

// ===== 2 PROOF =====
body.push(H1("2. Proof of Originality of the Data"));
body.push(P(`The data is real and was given by the business owners. The main proof items are listed below, and the signed papers are in the file "Proof Of Originality.pdf".`));
body.push(tabcap(1,"Business identification and proof details"));
body.push(table(["Item","Detail"],[
 ["Business (main)","Ankita Cashew Processing"],["Sister unit","Adrita Cashew Processing"],
 ["Owner (Ankita)","Usha Jana"],["Owner (Adrita)","Mithu Jana"],
 ["GST number","19BAKPJ1580R1ZY"],
 ["Address","Sitalpur, P.S. Contai, Purba Medinipur, West Bengal, India"],
 ["Contact","+91 8597310103 | ankitacashew@gmail.com"],
 ["NOC given on","08 June 2026 (signed by Usha Jana, Owner)"],
],[3000,6360]));
// links
const vlink=links.video||"[INSERT PORTAL VIDEO LINK BEFORE SUBMISSION]";
const llink=links.letter||"[INSERT PORTAL LETTER LINK BEFORE SUBMISSION]";
body.push(new Paragraph({spacing:{before:120,after:40},children:[new TextRun({text:"2.1 Links (available in the portal section)",bold:true,size:22,color:ORANGE})]}));
const linkLine=(label,url)=>new Paragraph({spacing:{after:80},children:[new TextRun({text:label+" ",size:22,bold:true}),
  url.startsWith("http")?new ExternalHyperlink({link:url,children:[new TextRun({text:url,style:"Hyperlink",size:22})]}):new TextRun({text:url,size:22,italics:true,color:"C55A11"})]});
body.push(linkLine("Video interview with the owner:",vlink));
body.push(linkLine("Signed owner letter (NOC):",llink));
body.push(P(`The proof has three parts: the signed NOC on the company letterhead, the video talk and letter kept in the portal, and the Tally files for sales, purchase and stock. These files were turned into the clean data used in this report.`));
body.push(new Paragraph({children:[new PageBreak()]}));

// ===== 3 METADATA + DATA COLLECTION =====
body.push(H1("3. Metadata and Data Collection"));
body.push(P(`How the data was collected: the data was taken straight from the two units as Tally accounting exports. It came in five kinds — (1) raw nut sales, (2) kernel sales, (3) raw nut purchases, (4) kernel purchases, and (5) year-end stock records with weights for each grade. Ankita also had a small export-sales record. Two files were given: "Sales Purchase 22-26.xlsx" (nine ledger sheets) and "Stock Summary.xlsx" (eight yearly stock sheets). In the ledger sheets several years were stacked one below the other. The data runs from 07 April 2022 to 10 February 2026. Both files were used: the ledgers gave the sales, purchase and price work, and the stock file gave the yield, stock movement and ABC work.`));
body.push(P(`The table below is the data dictionary. It lists each feature, what it means, its unit, its data type, and which business problem it helps with.`));
body.push(tabcap(2,"Data dictionary — features, units, data types and relevance"));
body.push(table(["Feature","Description","Unit","Data type","Helps with"],[
 ["Company","Ankita or Adrita","—","Text","All problems"],
 ["Flow","Sale, Purchase or Export","—","Text","Sales, cost"],
 ["Goods","RCN (raw nut) or FCN (kernel)","—","Text","Yield, mix"],
 ["Date","Date of the entry","Date","Date","Demand, season"],
 ["Counterparty","Buyer or supplier name","—","Text","Buyer risk"],
 ["Amount","Value of the entry","₹","Number","Sales, P&L"],
 ["Financial Year","FY2022-23 to FY2025-26","Year","Text","Yearly trend"],
 ["RCN bought","Raw nuts bought per month","KG","Number","Yield, price"],
 ["RCN value","Value of raw nuts bought","₹","Number","Price/kg"],
 ["Kernel sold","Whole + broken sold per month","KG","Number","Yield"],
 ["Stock value","Closing value of each item","₹","Number","ABC"],
],[1700,2700,900,1200,2860]));
body.push(P(`The number of entries by unit and action is shown next.`));
body.push(tabcap(3,"Record counts by unit and action"));
body.push(table(["Company","Flow","No. of entries","Total value (₹)"],o.company_flow.map(r=>[r.Company,r.Flow,inr(r.count),inr(r.sum)]),[2400,2400,2280,2280]));
body.push(new Paragraph({children:[new PageBreak()]}));

// ===== 4 DESCRIPTIVE STATS =====
body.push(H1("4. Descriptive Statistics"));
body.push(P(`This section shows the key numbers for every variable used in the analysis. Mean is the average. Median is the middle value. SD (standard deviation) shows how far the values spread from the average. Min and Max are the smallest and largest values.`));
body.push(tabcap(4,"Descriptive statistics of all variables used in the analysis"));
const order=["Sale amount (₹)","Purchase amount (₹)","Monthly sales (₹)","RCN buying price (₹/KG)","Monthly RCN bought (KG)","Monthly kernel sold (KG)","Yield per cycle (%)"];
body.push(table(["Variable","Count","Mean","Median","SD","Min","Max"],
 order.map(k=>{const s=D2[k];const f=k.includes('%')?(x=>x.toFixed(1)):(k.includes('/KG')?(x=>x.toFixed(1)):(x=>inr(x)));return [k,String(s.count),f(s.mean),f(s.median),f(s.std),f(s.min),f(s.max)];}),
 [2520,820,1300,1300,1300,1060,1060]));
body.push(P(`What these numbers mean for the business. A sale entry is about ₹${inr(sa.mean)} on average, but the median is lower at ₹${inr(sa.median)}. The average is higher than the median, and the SD (₹${inr(sa.std)}) is large. This tells us a few very big orders pull the average up, while most orders are smaller. The biggest single sale was ₹${inr(sa.max)}, about ${(sa.max/sa.mean).toFixed(0)} times the average. So the business depends on a small number of large orders.`));
body.push(P(`Purchases are bigger per entry than sales (average ₹${inr(pu.mean)} against ₹${inr(sa.mean)}). This is because raw nuts are bought in large lots a few times a year, while kernels are sold in many smaller lots. Monthly sales average about ₹${lk(D2["Monthly sales (₹)"].mean)} lakh and swing from ₹${lk(D2["Monthly sales (₹)"].min)} lakh to ₹${lk(D2["Monthly sales (₹)"].max)} lakh, which shows the strong season.`));
body.push(P(`The raw nut price averages ₹${D2["RCN buying price (₹/KG)"].mean.toFixed(0)} per kg and ranges from ₹${D2["RCN buying price (₹/KG)"].min.toFixed(0)} to ₹${D2["RCN buying price (₹/KG)"].max.toFixed(0)} per kg — a wide range that shows price risk. Yield per cycle averages ${D2["Yield per cycle (%)"].mean.toFixed(1)} percent (SD ${D2["Yield per cycle (%)"].std.toFixed(1)} points), lower than the 22 percent target, which is studied in detail in Section 6.`));
body.push(new Paragraph({children:[new PageBreak()]}));

// ===== 5 METHODS =====
body.push(H1("5. Methods (How the Analysis Was Done)"));
body.push(P(`The methods are grouped by the problem they solve.`));
body.push(H2("5.1 Data Cleaning and Preprocessing"));
body.push(P(`The raw sheets were not ready to use. Each cleaning step is shown below with the actual count from this data set.`));
[`Read all nine ledger sheets: ${inr(cc.scanned)} cells were scanned in total.`,
`Keep only real entries: a row was kept only if it started with a real date. This dropped ${inr(cc.blank_skipped)} non-data rows such as titles, addresses, the "Date / Particulars" header lines, and sub-totals.`,
`Remove duplicates: ${cc.duplicates} exact duplicate rows were found and removed. For example, the same ₹70,000 cash sale (Ankita, 31 Mar 2023) and the same ₹25,00,000 raw nut purchase (Ankita, 30 Jul 2024) each appeared twice. These are double-export entries, so one copy of each was dropped.`,
`Check missing values: after cleaning, ${cc.missing} rows had a missing amount. The sale value was read from the credit column and the purchase value from the debit column, so no amount was left blank.`,
`Tag each entry: every row was marked with its unit, action (sale/purchase), goods (RCN/FCN) and financial year, giving one clean list of ${inr(cc.kept)} entries.`].forEach(t=>body.push(bul(t)));
body.push(P(`The stock file was cleaned in the same way: weights written as text like "172620 KG" were turned into plain numbers, and the monthly inward and outward tables for each grade were pulled out.`));

body.push(H2("5.2 Demand Forecasting"));
body.push(P(`The clean sales were added up by month to make a time line. Three steps were used, as set out in the proposal.`));
body.push(P(`Step 1 — Decomposition. The time line is split into its trend (the slow long-term move) and its season (the repeating monthly pattern). This shows the shape of demand.`));
body.push(P(`Step 2 — Linear trend (base model). A straight line is fitted to the sales over time. Its form is:`));
body.push(eq("Sales = a + b × t"));
body.push(P(`Here t is the month number, a is the starting level, and b is the rise per month. It is a simple base model to compare against.`));
body.push(P(`Step 3 — ARIMA model. ARIMA stands for AutoRegressive Integrated Moving Average. In plain words it has three parts. AR (auto-regressive) means it leans on recent past sales. I (integrated) means it works on the change from one month to the next, which removes the trend. MA (moving average) means it corrects itself using how wrong it was in the recent past. Its form on the differenced series is:`));
body.push(eq("Ŷₜ = c + φ₁Y(t−1) + … + φₚY(t−p) + θ₁e(t−1) + … + θq e(t−q) + eₜ"));
body.push(P(`Here Ŷₜ is the predicted value, the φ (phi) terms weight past sales, the θ (theta) terms weight past errors e, and c is a constant.`));
body.push(P(`Checking accuracy. The last 6 months were kept aside as a test. Each model guessed those months, and the guesses were compared with what really happened, using two error scores (a smaller score is better):`));
body.push(eq("RMSE = √[ (1/n) × Σ (Aᵢ − Fᵢ)² ]"));
body.push(eq("MAE = (1/n) × Σ | Aᵢ − Fᵢ |"));
body.push(P(`Here Aᵢ is the actual sales, Fᵢ is the forecast, and n is the number of test months. RMSE squares the errors, so it punishes big misses more. MAE is the plain average error.`));

body.push(H2("5.3 Inventory Yield Analysis"));
body.push(P(`Yield means how much kernel comes out of the raw nuts. Each unit-year is treated as one processing cycle, because the business does not yet keep records for smaller batches. The yield is worked out as:`));
body.push(eq("Yield (%) = ( Kernel Output [KG] ÷ Raw Nut Input [KG] ) × 100"));
body.push(P(`Variance check. Each cycle's yield is compared with a target of 22 percent (a normal figure for this trade). Variance = actual yield − target. A negative variance means the cycle is below target.`));
body.push(P(`Z-score (to flag odd cycles). The Z-score tells how far a cycle's yield is from the average, in units of SD:`));
body.push(eq("Z = ( x − μ ) ÷ σ"));
body.push(P(`Here x is one cycle's yield, μ (mu) is the average yield of all cycles, and σ (sigma) is the standard deviation. A cycle is flagged if its Z-score is beyond plus or minus 1.5.`));
body.push(P(`ABC classification. The stock items are sorted by value. The few items that hold most of the value are class A (watch closely), the next group is B, and the rest is C. This shows where the money sits.`));
body.push(P(`One limit should be noted. The two units also buy and sell kernels as a trade, and they hold raw nuts in stock from one year to the next. So a single year's yield can look high or low just because of timing. For this reason the full-period figure is the more reliable one.`));
body.push(new Paragraph({children:[new PageBreak()]}));

// ===== 6 RESULTS =====
body.push(H1("6. Results and Findings"));
body.push(P(`The charts below are grouped by problem: business overview, demand and price forecasting, inventory and yield, and revenue and buyers.`));

body.push(H2("6.1 Business Overview"));
body.push(tabcap(5,"Annual sales and purchase (both units together)"));
body.push(table(["Financial Year","Sales (₹)","Purchase (₹)","Net (₹)"],ann.map(r=>[r.fy,inr(r.sales),inr(r.purchase),inr(r.sales-r.purchase)]),[2400,2320,2320,2320]));
body.push(...fig("fig01_annual.png",540,300,1,"Annual sales versus purchase (₹ crore)."));
body.push(P(`Sales grew each year up to about ₹${cr(Math.max(...ann.map(r=>r.sales)))} crore in FY2024-25, then eased in the part-year FY2025-26 (data only to February 2026). In three of the four years sales were higher than purchases. Only in FY2022-23 were purchases a little higher than sales, as the units were building up raw nut stock that year.`));

body.push(H2("6.2 Demand and Price Forecasting"));
body.push(...fig("fig02_decomposition.png",500,385,2,"Monthly sales split into trend and season."));
body.push(P(`Comment on Figure 2. The top panel is the actual monthly sales. The middle panel is the trend — it drifts slowly upward, so demand is growing over the years. The bottom panel is the season — the same up-and-down shape repeats every 12 months, which proves the sales have a fixed seasonal pattern, not random noise. Together they show a rising business with a strong yearly cycle.`));
body.push(...fig("fig03_seasonality.png",520,270,3,"Total sales by month, all years pooled."));
{const ms=Object.entries({Apr:0});/*placeholder*/}
body.push(P(`Sales are highest around January and again in May–June, and slowest in July and March. This fits wedding and festival demand for kernels and the buying of raw nuts before the monsoon. Knowing this helps plan how much to make and buy in each part of the year.`));
body.push(P(`Validation before forecasting. Before trusting any forecast, the models were tested on the last 6 months (which were hidden from them). Figure 4 shows the test, and Table 6 gives the error scores.`));
body.push(...fig("fig04_validation.png",520,260,4,"Model check on the 6-month holdout — actual versus predicted."));
body.push(tabcap(6,"Forecast accuracy on the 6-month test (lower is better)"));
body.push(table(["Model","RMSE (₹)","MAE (₹)"],[
 ["ARIMA",inr(fc.sarima_rmse),inr(fc.sarima_mae)],
 ["Linear trend (base)",inr(fc.lin_rmse),inr(fc.lin_mae)],
 ["Naive (same month last year)",inr(fc.naive_rmse),inr(fc.naive_mae)],
],[3360,3000,3000]));
body.push(...fig("fig05_accuracy.png",470,270,5,"Forecast accuracy (RMSE and MAE)."));
body.push(P(`On the test, the simple linear trend gave the smallest error (MAE about ₹${lk(fc.lin_mae)} lakh), a little better than ARIMA. Both were far better than the naive guess. The monthly sales jump around a lot, so no model is perfect, but the trend is clearly upward at about ₹${inr(fc.lin_slope)} more sales per month.`));
body.push(...fig("fig06_forecast.png",560,235,6,"Sales forecast for the next 4 months (ARIMA)."));
body.push(P(`The ARIMA forecast for the next four months is roughly ₹${cr(fc.forecast[Object.keys(fc.forecast)[0]].mean)} to ₹${cr(fc.forecast[Object.keys(fc.forecast)[2]].mean)} crore per month, in line with the usual season.`));
body.push(...fig("fig07_rcn_price.png",550,230,7,"Raw cashew nut buying price per kg."));
body.push(tabcap(7,"Average raw nut price by year (₹/KG)"));
body.push(table(["Year","Average RCN price (₹/KG)"],Object.entries(rcnP).map(([y,p])=>[y,String(p)]),[3000,3000]));
body.push(P(`The raw nut price moves a lot within a year and has risen over time, from about ₹${rcnP['2023']} per kg in 2023 to about ₹${rcnP['2026']} per kg in 2026. A higher raw nut price raises costs, so the buying time and the selling price both need care. This is the price-change problem named in the proposal, now shown with real numbers.`));

body.push(H2("6.3 Inventory and Yield"));
body.push(...fig("fig08_movement.png",560,235,8,"Monthly stock movement — raw nuts in versus kernels out."));
body.push(P(`Raw nuts are bought in large lots a few times a year, mostly before and during the season. Kernels are sold in smaller, steadier amounts through the year. So the business holds raw nut stock for some months before turning it into kernels and selling it.`));
body.push(tabcap(8,"Yield per cycle with target, variance and Z-score"));
body.push(table(["Cycle","Raw Nut In (KG)","Kernel Out (KG)","Yield %","Target %","Variance","Z-score","Outlier?"],
 o.yield_batches.map(r=>[`${r.Company} ${r.Year}`,inr(r.RCN_Purchased),inr(r.Kernel_Out),r.Yield_pct.toFixed(1),"22.0",r.Variance.toFixed(1),r.Zscore.toFixed(2),r.Outlier]),
 [1860,1660,1520,920,900,1020,900,580]));
body.push(...fig("fig09_yield.png",520,255,9,"Yield of each cycle against the 22% target."));
body.push(...fig("fig10_zscore.png",520,245,10,"Yield Z-score. Bars past ±1.5 are flagged."));
body.push(tabcap(9,"Yield by unit over the full period"));
body.push(table(["Unit","Raw Nut In (KG)","Kernel Out (KG)","Yield %","Whole share %"],
 o.yield_company.map(r=>[r.Company,inr(r.RCN),inr(r.Kern),r.Recovery_pct.toFixed(1),r.Whole_share.toFixed(1)]),[2360,2000,1800,1600,1600]));
body.push(P(`Over the full period the yield is about ${adrY.Recovery_pct} percent for Adrita and ${ankY.Recovery_pct} percent for Ankita. Adrita is close to the 22 percent target. Ankita is lower, mainly because it holds more raw nut stock across years, not because of poor work. Most cycles are below target. The Z-score test flags one cycle, Ankita FY2022-23, where the yield (10.8 percent) is well below the rest — a clear case to look into. Whole kernels make up about ${ankY.Whole_share} percent of the graded output, which is a good quality mix.`));
body.push(tabcap(10,"ABC classification of stock by value"));
body.push(table(["Stock group","Value (₹)","% of total","Cumulative %","Class"],
 o.abc.map(r=>[r.Group,inr(r.Value),r.Pct.toFixed(1),r.CumPct.toFixed(1),r.Class]),[2760,1900,1300,1500,900]));
body.push(...fig("fig11_abc.png",560,275,11,"ABC chart of stock by value (Pareto)."));
body.push(P(`Just three items — raw nuts, whole kernels and broken kernels — hold about 98 percent of the stock value. Raw nuts alone are the "A" group at ${o.abc[0].Pct.toFixed(1)} percent. So the owners should watch raw nut and kernel stock most closely, since that is where the money is.`));

body.push(H2("6.4 Revenue Mix and Buyers"));
body.push(...fig("fig12_mix.png",350,320,12,"Revenue mix — kernels versus raw nut."));
body.push(P(`Kernels bring in about ${fcMix} percent of the money. Processing, not raw nut trading, is the heart of the business.`));
body.push(...fig("fig13_buyers.png",520,275,13,"Top buyers by sales value."));
body.push(P(`The top five buyers bring in about ${o.buyer_concentration_top5.toFixed(0)} percent of all sales. So the business leans on a few big buyers. This is why the next stage of the project will score buyers for payment risk, as planned in the proposal.`));

body.push(H2("6.5 Main Findings"));
[`Over four years, sales were about ₹${cr(o.totals.total_sales)} crore and purchases about ₹${cr(o.totals.total_purchase)} crore. Sales were higher than purchases in three of the four years; in FY2022-23 purchases were slightly higher.`,
`Kernels bring in about ${Math.round(fcMix)} percent of the money; processing is the heart of the business.`,
`Sales follow a season, highest in January and in May–June, with a slowly rising trend.`,
`The linear trend forecast was a little better than ARIMA on the test; both beat the naive guess.`,
`The raw nut price has risen over the years and moves a lot, so buying time and selling price need care.`,
`Yield is about ${adrY.Recovery_pct}% (Adrita) and ${ankY.Recovery_pct}% (Ankita); most cycles are below 22%, and Ankita FY2022-23 is a low outlier.`,
`Raw nuts and kernels hold about 98% of stock value (ABC A and B), so they need the closest watch.`,
`The top five buyers give about ${o.buyer_concentration_top5.toFixed(0)}% of sales, so the business depends on a few buyers.`].forEach(t=>body.push(bul(t)));

const doc=new Document({
 styles:{default:{document:{run:{font:"Arial",size:22}}},
  paragraphStyles:[
   {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:30,bold:true,font:"Arial",color:NAVY},paragraph:{spacing:{before:240,after:160},outlineLevel:0}},
   {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:25,bold:true,font:"Arial",color:ORANGE},paragraph:{spacing:{before:160,after:100},outlineLevel:1}}]},
 numbering:{config:[{reference:"bul",levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:560,hanging:280}}}}]}]},
 sections:[{properties:{page:{size:{width:12240,height:15840},margin:{top:1440,right:1440,bottom:1440,left:1440}}},
   footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Mid-term Report — Ankita Cashew Processing   |   Page ",size:16,color:"888888"}),new TextRun({children:[PageNumber.CURRENT],size:16,color:"888888"})]})]})},
   children:body}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync("Ankita_Cashew_Midterm_Report.docx",b);console.log("written; pages known:",Object.keys(pages).length)});
