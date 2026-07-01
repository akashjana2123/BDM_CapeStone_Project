const fs=require('fs');
const D=require('docx');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,ImageRun,Footer,
AlignmentType,LevelFormat,HeadingLevel,BorderStyle,WidthType,ShadingType,
VerticalAlign,PageNumber,PageBreak,TabStopType,LeaderType,ExternalHyperlink}=D;
const RF=JSON.parse(fs.readFileSync('results_final.json'));
const RM=JSON.parse(fs.readFileSync('results.json'));
const ANN=JSON.parse(fs.readFileSync('annual.json'));
const pages=fs.existsSync('toc_pages_final.json')?JSON.parse(fs.readFileSync('toc_pages_final.json')):{};
const links=fs.existsSync('links.json')?JSON.parse(fs.readFileSync('links.json')):{};
const TEAL="0F766E",AMBER="B45309",SLATE="334155",HDR="0F766E";
const inr=n=>Math.round(n).toLocaleString('en-IN');
const cr=n=>(n/1e7).toFixed(2), lk=n=>(n/1e5).toFixed(1);
const F="figures_final/";
const H1=t=>new Paragraph({heading:HeadingLevel.HEADING_1,children:[new TextRun(t)]});
const H2=t=>new Paragraph({heading:HeadingLevel.HEADING_2,children:[new TextRun(t)]});
const P=(t,j=true)=>new Paragraph({spacing:{after:120,line:284},alignment:j?AlignmentType.JUSTIFIED:AlignmentType.LEFT,children:[new TextRun({text:t,size:22})]});
const bul=t=>new Paragraph({numbering:{reference:"bul",level:0},spacing:{after:80},children:[new TextRun({text:t,size:22})]});
const eq=t=>new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:60,after:120},children:[new TextRun({text:t,bold:true,size:23})]});
const figcap=(n,t)=>new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:200},children:[new TextRun({text:`Figure ${n}: ${t}`,italics:true,size:18,color:"555555"})]});
const tabcap=(n,t)=>new Paragraph({spacing:{before:120,after:40},children:[new TextRun({text:`Table ${n}: ${t}`,bold:true,size:20,color:"333333"})]});
const fig=(file,w,h,n,cap)=>[new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:60},children:[new ImageRun({type:"png",data:fs.readFileSync(F+file),transformation:{width:w,height:h},altText:{title:cap,description:cap,name:cap}})]}),figcap(n,cap)];
const bd={style:BorderStyle.SINGLE,size:1,color:"BBBBBB"};const borders={top:bd,bottom:bd,left:bd,right:bd};
function table(headers,rows,widths){const total=widths.reduce((a,b)=>a+b,0);
 const mk=(txt,opt={})=>new TableCell({borders,width:{size:opt.w,type:WidthType.DXA},shading:{fill:opt.fill||"FFFFFF",type:ShadingType.CLEAR},margins:{top:50,bottom:50,left:90,right:90},verticalAlign:VerticalAlign.CENTER,children:[new Paragraph({alignment:opt.al||AlignmentType.LEFT,children:[new TextRun({text:String(txt),bold:opt.b||false,color:opt.color||"000000",size:opt.sz||19})]})]});
 const hr=new TableRow({tableHeader:true,children:headers.map((h,i)=>mk(h,{w:widths[i],b:true,color:"FFFFFF",fill:HDR,al:i==0?AlignmentType.LEFT:AlignmentType.CENTER}))});
 const drs=rows.map((r,ri)=>new TableRow({children:r.map((c,i)=>mk(c,{w:widths[i],al:i==0?AlignmentType.LEFT:AlignmentType.CENTER,fill:ri%2?"EEF6F4":"FFFFFF"}))}));
 return new Table({width:{size:total,type:WidthType.DXA},columnWidths:widths,rows:[hr,...drs]});}
function toc(label,key,sub){const pg=pages[key]!==undefined?String(pages[key]):"";
 return new Paragraph({tabStops:[{type:TabStopType.RIGHT,position:9020,leader:LeaderType.DOT}],spacing:{after:60},indent:sub?{left:360}:undefined,children:[new TextRun({text:label,size:22,bold:!sub}),new TextRun({text:"\t"+pg,size:22})]});}

const sa=RF.descriptive["Sale amount (₹)"], pu=RF.descriptive["Purchase amount (₹)"];
const fcMix=(RF.sales_by_goods.FCN/(RF.sales_by_goods.FCN+RF.sales_by_goods.RCN)*100).toFixed(1);
const pnl=RF.pnl, base=pnl.base, by=RF.buyer, an=RF.anomaly;
const fc=RM.forecast; const yc=RM.yield_company;
const ankY=yc.find(x=>x.Company=='Ankita'), adrY=yc.find(x=>x.Company=='Adrita');

let body=[];
// TITLE
body.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:1100,after:120},children:[new TextRun({text:"Analytical Study of Ankita Cashew Processing",bold:true,size:36,color:TEAL})]}));
body.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:600},children:[new TextRun({text:"Final Report for the BDM Capstone Project",bold:true,size:26})]}));
["Submitted by","Name: Akash Jana","Roll Number: 23f2000990"].forEach(t=>body.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:80},children:[new TextRun({text:t,size:24})]})));
body.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:160,after:160},children:[new ImageRun({type:"png",data:fs.readFileSync("iitm_logo.png"),transformation:{width:220,height:220},altText:{title:"IITM",description:"IITM",name:"IITM"}})]}));
["IITM Online BS Degree Program,","Indian Institute of Technology Madras, Chennai","Tamil Nadu, India, 600036"].forEach(t=>body.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:60},children:[new TextRun({text:t,size:22})]})));
body.push(new Paragraph({children:[new PageBreak()]}));
// CONTENTS
body.push(H1("Contents"));
[["Declaration Statement","declaration",0],["1. Executive Summary and Title","s1",0],["2. Proof of Originality of the Data","s2",0],
 ["3. Metadata and Data Collection","s3",0],["4. Descriptive Statistics","s4",0],["5. Methods (How the Analysis Was Done)","s5",0],
 ["5.1 Data Cleaning and Preprocessing","s51",1],["5.2 Demand Forecasting","s52",1],["5.3 Inventory Yield Analysis","s53",1],
 ["5.4 Profit and Loss Modelling","s54",1],["5.5 Buyer Risk and Credit Scoring","s55",1],["5.6 Anomaly Detection","s56",1],
 ["6. Results and Findings","s6",0],["6.1 Business Overview","s61",1],["6.2 Demand and Price Forecasting","s62",1],
 ["6.3 Inventory and Yield","s63",1],["6.4 Profit and Loss","s64",1],["6.5 Buyer Risk and Credit","s65",1],["6.6 Anomaly Detection","s66",1],
 ["7. Final Recommendations","s7",0],["8. Conclusion and Limitations","s8",0]
].forEach(([l,k,s])=>body.push(toc(l,k,s)));
body.push(new Paragraph({children:[new PageBreak()]}));
// DECLARATION
body.push(H1("Declaration Statement"));
[`I am working on a project titled "Analytical Study of Ankita Cashew Processing". I thank Ankita Cashew Processing and its sister unit Adrita Cashew Processing for giving me the data and help I needed for this project.`,
`The data in this report is real and correct to the best of my knowledge. It was taken from the business itself (primary source) and checked with care.`,
`I have explained how I collected and studied the data in this report. The results shown here are a true picture of what the data says.`,
`I did this project on my own. I did not copy or work jointly with anyone else. If any copied work is found at any stage, I accept that the authority can take action against me.`,
`The suggestions in this report are only for this business and this project. They cannot be used elsewhere with an IIT Madras tag. IIT Madras does not endorse this.`].forEach(t=>body.push(P(t)));
body.push(new Paragraph({spacing:{before:200},children:[new TextRun({text:"Signature of Candidate: (Digital Signature)",size:22})]}));
body.push(new Paragraph({children:[new TextRun({text:"Name: Akash Jana",size:22})]}));
body.push(new Paragraph({children:[new TextRun({text:"Date: 01 May 2026",size:22})]}));
body.push(new Paragraph({children:[new PageBreak()]}));

// 1 EXEC SUMMARY
body.push(H1("1. Executive Summary and Title"));
body.push(P(`Ankita Cashew Processing is a cashew business in West Bengal, India. It buys raw cashew nuts (RCN), processes them into kernels (the eatable cashew, called FCN here), and sells to other traders. Its sister unit, Adrita Cashew Processing, does the same work. This report studies both units together. It is the final report of the project, and it builds on the mid-term work (cleaning, forecasting and yield) by adding profit-and-loss modelling, buyer risk scoring and anomaly detection, ending with clear recommendations.`));
body.push(P(`The data is real and is backed by clear proof. The owner, Usha Jana, gave a signed No Objection Certificate (NOC) on the company letterhead, dated 08 June 2026. A short video talk with the owner and the signed letter are kept in the project portal (links in Section 2). The raw Tally accounting files for sales, purchase and stock were shared directly by the business. Together these show that the data comes from the business itself.`));
body.push(P(`The cleaned data has ${RF.totals.n_sales+RF.totals.n_purchase} entries from April 2022 to February 2026. A single sale is about ₹${inr(sa.mean)} on average, with a middle value (median) of ₹${inr(sa.median)} and a standard deviation (SD) of ₹${inr(sa.std)}; the largest sale is ₹${inr(sa.max)}. Over four years the two units sold about ₹${cr(RF.totals.total_sales)} crore and bought about ₹${cr(RF.totals.total_purchase)} crore of goods. Kernels bring in about ${fcMix} percent of sales.`));
body.push(P(`The methods used in this report, grouped by problem, are:`));
[`Data cleaning and preprocessing — read the raw Tally sheets, drop non-data rows, remove duplicates, tag each entry.`,
`Demand forecasting — decomposition into trend and season, a linear-trend base model, and an ARIMA model; accuracy checked with RMSE and MAE on a 6-month test.`,
`Inventory yield analysis — yield per cycle, variance against a 22% target, and a Z-score check for odd cycles.`,
`Profit and loss modelling — a per-kilogram contribution-margin model, break-even, and a price sensitivity grid, by grade.`,
`Buyer risk and credit scoring — a buyer scorecard, credit-limit recommendations, and an illustrative logistic-regression delay model.`,
`Anomaly detection — Z-score thresholding and the Isolation Forest method, plus a stock reconciliation check.`].forEach(t=>body.push(bul(t)));
body.push(P(`Main results: the contribution is about ₹${Math.round(base.contribution)} per kg of kernel (${base.margin_pct.toFixed(1)}% margin) and the break-even kernel price is about ₹${pnl.be_price}/kg; whole kernels earn a profit while broken kernels barely cover cost. Of ${by.n_buyers} buyers, ${by.tiers.High} are high-risk and need tight credit control. Anomaly checks flagged ${an.n_total} transactions (${an.pct}%) for review. Tools used: Microsoft Excel for tables and charts, and Python (pandas, statsmodels, scikit-learn, matplotlib) for the analysis. All money is in ₹ and all weights in KG.`));
body.push(new Paragraph({children:[new PageBreak()]}));

// 2 PROOF
body.push(H1("2. Proof of Originality of the Data"));
body.push(P(`The data is real and was given by the business owners. The main proof items are below, and the signed papers are in the file "Proof Of Originality.pdf".`));
body.push(tabcap(1,"Business identification and proof details"));
body.push(table(["Item","Detail"],[["Business (main)","Ankita Cashew Processing"],["Sister unit","Adrita Cashew Processing"],["Owner (Ankita)","Usha Jana"],["Owner (Adrita)","Mithu Jana"],["GST number","19BAKPJ1580R1ZY"],["Address","Sitalpur, P.S. Contai, Purba Medinipur, West Bengal, India"],["Contact","+91 8597310103 | ankitacashew@gmail.com"],["NOC given on","08 June 2026 (signed by Usha Jana, Owner)"]],[3000,6360]));
body.push(new Paragraph({spacing:{before:120,after:40},children:[new TextRun({text:"2.1 Links (available in the portal section)",bold:true,size:22,color:AMBER})]}));
const vlink=links.video||"[INSERT PORTAL VIDEO LINK BEFORE SUBMISSION]";
const llink=links.letter||"[INSERT PORTAL LETTER LINK BEFORE SUBMISSION]";
const linkLine=(label,url)=>new Paragraph({spacing:{after:80},children:[new TextRun({text:label+" ",size:22,bold:true}),url.startsWith("http")?new ExternalHyperlink({link:url,children:[new TextRun({text:url,style:"Hyperlink",size:22})]}):new TextRun({text:url,size:22,italics:true,color:AMBER})]});
body.push(linkLine("Video interview with the owner:",vlink));
body.push(linkLine("Signed owner letter (NOC):",llink));
body.push(P(`The proof has three parts: the signed NOC on the company letterhead, the video talk and letter kept in the portal, and the Tally files for sales, purchase and stock. These files were turned into the clean data used in this report.`));
body.push(new Paragraph({children:[new PageBreak()]}));

// 3 METADATA
body.push(H1("3. Metadata and Data Collection"));
body.push(P(`How the data was collected: the data was taken straight from the two units as Tally accounting exports. It came in five kinds — raw nut sales, kernel sales, raw nut purchases, kernel purchases, and year-end stock records with weights for each grade. Ankita also had a small export-sales record. Two files were given: "Sales Purchase 22-26.xlsx" (nine ledger sheets) and "Stock Summary.xlsx" (eight yearly stock sheets). The data runs from 07 April 2022 to 10 February 2026. Both files were used in the analysis.`));
body.push(P(`The table below is the data dictionary. It lists each feature, what it means, its unit, its data type, and which problem it helps with.`));
body.push(tabcap(2,"Data dictionary — features, units, data types and relevance"));
body.push(table(["Feature","Description","Unit","Type","Helps with"],[
 ["Company","Ankita or Adrita","—","Text","All"],["Flow","Sale, Purchase or Export","—","Text","Sales, cost"],
 ["Goods","RCN or FCN","—","Text","Yield, P&L"],["Date","Date of entry","Date","Date","Demand, season"],
 ["Counterparty","Buyer or supplier","—","Text","Buyer risk"],["Amount","Value of entry","₹","Number","Sales, P&L"],
 ["Financial Year","FY2022-23 … FY2025-26","Year","Text","Yearly trend"],["RCN bought","Raw nuts bought/month","KG","Number","Yield, price"],
 ["RCN value","Value of raw nuts","₹","Number","RCN price/kg"],["Kernel sold","Whole+broken sold/month","KG","Number","Yield, P&L"],
 ["Stock value","Closing value per item","₹","Number","ABC"],["Buyer revenue (derived)","Total sales to a buyer","₹","Number","Risk score"],
 ["Recency (derived)","Days since buyer's last order","Days","Number","Risk score"],["Contribution (derived)","Profit per kg of kernel","₹/KG","Number","P&L"],
],[1800,2560,820,900,3280]));
body.push(P(`Record counts by unit and action are shown next.`));
body.push(tabcap(3,"Record counts by unit and action"));
body.push(table(["Company","Flow","Entries","Total value (₹)"],RM.company_flow.map(r=>[r.Company,r.Flow,inr(r.count),inr(r.sum)]),[2400,2400,2280,2280]));
body.push(new Paragraph({children:[new PageBreak()]}));

// 4 DESCRIPTIVE
body.push(H1("4. Descriptive Statistics"));
body.push(P(`This section shows the key numbers for every variable used in the analysis, including the new variables built for the final stage (buyer revenue, buyer orders, risk score and per-kg contribution). Mean is the average, median is the middle value, SD is the spread, and min/max are the smallest and largest values.`));
body.push(tabcap(4,"Descriptive statistics of all variables used in the analysis"));
const dorder=["Sale amount (₹)","Purchase amount (₹)","RCN price (₹/KG)","Buyer revenue (₹)","Buyer orders (count)","Buyer risk score (0-100)","Contribution per kg kernel (₹)"];
body.push(table(["Variable","Count","Mean","Median","SD","Min","Max"],dorder.map(k=>{const s=RF.descriptive[k];const f=(k.includes("count")||k.includes("0-100")||k.includes("/KG")||k.includes("per kg"))?(x=>(+x).toFixed(1)):(x=>inr(x));return [k,String(s.count),f(s.mean),f(s.median),f(s.std),f(s.min),f(s.max)];}),[2520,820,1300,1300,1300,1060,1060]));
body.push(P(`What these numbers mean for the business. A sale entry is about ₹${inr(sa.mean)} on average, but the median is lower at ₹${inr(sa.median)} and the SD (₹${inr(sa.std)}) is large. So a few very big orders pull the average up while most orders are smaller; the biggest single sale (₹${inr(sa.max)}) is about ${(sa.max/sa.mean).toFixed(0)} times the average. The business therefore leans on a small number of large orders, which is exactly why buyer risk (Section 6.5) matters.`));
body.push(P(`Purchases are bigger per entry than sales (average ₹${inr(pu.mean)} versus ₹${inr(sa.mean)}) because raw nuts are bought in large lots a few times a year. The raw nut price averages about ₹${RF.descriptive["RCN price (₹/KG)"].mean.toFixed(0)}/kg and ranges from ₹${RF.descriptive["RCN price (₹/KG)"].min.toFixed(0)} to ₹${RF.descriptive["RCN price (₹/KG)"].max.toFixed(0)}/kg — a wide range that drives the price risk studied in the P&L. Across ${RF.descriptive["Buyer revenue (₹)"].count} buyers, the average buyer brings ₹${inr(RF.descriptive["Buyer revenue (₹)"].mean)} of business but the median is much lower, again showing a few big buyers and many small ones. The risk score averages ${RF.descriptive["Buyer risk score (0-100)"].mean.toFixed(0)} out of 100, and the per-kg contribution averages about ₹${RF.descriptive["Contribution per kg kernel (₹)"].mean.toFixed(0)}/kg across grades.`));
body.push(new Paragraph({children:[new PageBreak()]}));

// 5 METHODS
body.push(H1("5. Methods (How the Analysis Was Done)"));
body.push(P(`The methods are grouped by the problem they solve.`));
body.push(H2("5.1 Data Cleaning and Preprocessing"));
const cc=RF.clean_counts;
body.push(P(`The raw sheets were not ready to use. Each cleaning step is shown below with the actual count from this data set.`));
[`Read all nine ledger sheets: ${inr(cc.scanned)} cells were scanned.`,
`Keep only real entries: a row was kept only if it started with a real date. This dropped ${inr(cc.skipped)} non-data rows such as titles, addresses, the "Date / Particulars" headers and sub-totals.`,
`Remove duplicates: ${cc.duplicates} exact duplicate rows were found and removed — for example the same ₹70,000 cash sale (Ankita, 31 Mar 2023) and the same ₹25,00,000 raw nut purchase (Ankita, 30 Jul 2024) each appeared twice (double-export entries).`,
`Check missing values: 0 rows had a missing amount after cleaning (sale value read from the credit column, purchase value from the debit column).`,
`Tag each entry: every row was marked with its unit, action, goods and financial year, giving one clean list of ${inr(cc.kept)} entries.`].forEach(t=>body.push(bul(t)));
body.push(H2("5.2 Demand Forecasting"));
body.push(P(`Monthly sales were studied in three ways. First, decomposition splits the series into trend (the slow move) and season (the repeating monthly pattern). Second, a linear trend (a straight line, Sales = a + b × t) is used as a simple base model. Third, an ARIMA model is used for a short forecast.`));
body.push(P(`ARIMA means AutoRegressive Integrated Moving Average. In plain words: AR uses recent past sales, I works on the month-to-month change to remove the trend, and MA corrects using recent errors. Its form on the differenced series is:`));
body.push(eq("Ŷₜ = c + φ₁Y(t−1) + … + φₚY(t−p) + θ₁e(t−1) + … + θq e(t−q) + eₜ"));
body.push(P(`Accuracy was checked by holding out the last 6 months and comparing the forecast with the actual values, using two error scores (smaller is better):`));
body.push(eq("RMSE = √[ (1/n) × Σ (Aᵢ − Fᵢ)² ]      MAE = (1/n) × Σ | Aᵢ − Fᵢ |"));
body.push(H2("5.3 Inventory Yield Analysis"));
body.push(P(`Yield is how much kernel comes out of the raw nuts. Each unit-year is one processing cycle. Yield and the Z-score check are:`));
body.push(eq("Yield (%) = ( Kernel Output [KG] ÷ Raw Nut Input [KG] ) × 100      Z = ( x − μ ) ÷ σ"));
body.push(P(`Variance = actual yield − 22% target. A cycle is flagged if its Z-score is beyond ±1.5 (x = a cycle's yield, μ = mean yield, σ = standard deviation).`));
body.push(H2("5.4 Profit and Loss Modelling"));
body.push(P(`A per-kilogram contribution model was built for the kernels. Because the ledgers do not hold separate cost heads (labour, electricity, packing, transport), processing and overhead is taken as an assumption of ₹${pnl.assume.processing_per_kg_kernel}/kg of kernel, and a normal outturn of ${(pnl.assume.outturn*100).toFixed(0)}% is used (with sensitivity around both). The raw nut cost per kg of kernel is the raw nut price divided by the outturn. The contribution per kg of kernel is:`));
body.push(eq("Contribution = Selling price − ( RCN price ÷ Outturn ) − Processing cost"));
body.push(P(`Break-even is the point where contribution is zero. The break-even selling price equals the total cost per kg; the break-even outturn is RCN price ÷ (selling price − processing cost). A sensitivity grid then shows how the contribution changes when the raw nut price and the selling price each move by ±10%.`));
body.push(H2("5.5 Buyer Risk and Credit Scoring"));
body.push(P(`For each buyer, features were built from the sales history: total revenue, number of orders, average order size, recency (days since the last order), tenure (days from first to last order) and order frequency. A risk score from 0 to 100 was made by combining these (higher recency, lower frequency, lower tenure and larger single exposure all raise the score). Buyers were then placed in Low, Medium or High tiers, and a credit limit was suggested in proportion to one month of their business and their tier.`));
body.push(P(`The proposal also asks for a logistic-regression model of payment delay. The ledgers do not record payment dates or outstanding amounts, so a real delay model cannot be trained from this data. To still demonstrate the method, an illustrative logistic regression was fitted on a simulated delay outcome built from the buyer features. Logistic regression estimates the probability of delay as:`));
body.push(eq("P(delay) = 1 ÷ ( 1 + e^−( β₀ + β₁x₁ + β₂x₂ + … + βₖxₖ ) )"));
body.push(P(`Here x₁…xₖ are the buyer features and β are the weights the model learns. Accuracy is judged by the ROC curve and its area (AUC), where 1.0 is perfect and 0.5 is a coin toss. The model is clearly marked as illustrative; the business should start capturing invoice and payment dates so a real model can be trained later.`));
body.push(H2("5.6 Anomaly Detection"));
body.push(P(`Two methods find unusual transactions. First, a Z-score is computed within each company-flow-goods group, and a transaction is flagged if its Z-score is beyond ±3 (about 0.3% of normal data). Second, the Isolation Forest method is used. It builds many random trees that keep splitting the data; an unusual point gets separated from the rest in only a few splits, so a short "isolation path" means a high anomaly score. Finally, a reconciliation check compares each month's kernel sold against raw nuts bought and flags months whose ratio is more than two standard deviations from normal.`));
body.push(new Paragraph({children:[new PageBreak()]}));

// 6 RESULTS
body.push(H1("6. Results and Findings"));
body.push(P(`The charts are grouped by problem. None of them are reused from the mid-term report; all were drawn fresh for this final report.`));
body.push(H2("6.1 Business Overview"));
body.push(tabcap(5,"Annual sales and purchase (both units together)"));
body.push(table(["Financial Year","Sales (₹)","Purchase (₹)","Net (₹)"],ANN.map(r=>[r.fy,inr(r.sales),inr(r.purchase),inr(r.sales-r.purchase)]),[2400,2320,2320,2320]));
body.push(...fig("fig01_overview.png",540,290,1,"Sales and purchase by year."));
body.push(P(`Sales grew to about ₹${cr(Math.max(...ANN.map(r=>r.sales)))} crore in FY2024-25 and then eased in the part-year FY2025-26. In three of the four years sales were higher than purchases; only in FY2022-23 were purchases a little higher, as the units were building raw nut stock.`));

body.push(H2("6.2 Demand and Price Forecasting"));
body.push(...fig("fig02_forecast.png",560,250,2,"Sales forecast with the validation window shaded, then the next-4-month forecast."));
body.push(P(`Comment on Figure 2. The grey line is the actual monthly sales. The shaded band is the 6-month validation window: before trusting any forecast, the model first predicted these held-out months (the red dashed line) so its error could be measured against what really happened. Only after that check is the forward forecast (teal) drawn, with a light band for uncertainty. This order — validate first, predict second — is what makes the forecast trustworthy.`));
body.push(...fig("fig03_seasonality_heat.png",560,225,3,"Sales seasonality heatmap (₹ lakh by year and month)."));
body.push(P(`The heatmap shows the same months stay strong across years — January and the May–June window are the darkest (highest) cells, while July and March are lighter. This repeating pattern is the season the business can plan around.`));
body.push(tabcap(6,"Forecast accuracy on the 6-month test (lower is better)"));
body.push(table(["Model","RMSE (₹)","MAE (₹)"],[["ARIMA",inr(fc.sarima_rmse),inr(fc.sarima_mae)],["Linear (base)",inr(fc.lin_rmse),inr(fc.lin_mae)],["Naive (last year)",inr(fc.naive_rmse),inr(fc.naive_mae)]],[3360,3000,3000]));
body.push(P(`On the test the simple linear trend was a little better than ARIMA (MAE about ₹${lk(fc.lin_mae)} lakh), and both beat the naive guess. The trend is rising at about ₹${inr(fc.lin_slope)} more sales per month. The raw nut price has also risen over the years (from about ₹${RM.rcn_price_yearly['2023']} to ₹${RM.rcn_price_yearly['2026']} per kg), which feeds straight into the P&L below.`));

body.push(H2("6.3 Inventory and Yield"));
body.push(...fig("fig04_yield_control.png",560,250,4,"Yield control chart with mean, ±1.5σ limits and the 22% target."));
body.push(tabcap(7,"Yield per cycle (kernel out ÷ raw nut in)"));
body.push(table(["Cycle","Yield %","Variance","Z-score","Outlier?"],RM.yield_batches.map(r=>[`${r.Company} ${r.Year}`,r.Yield_pct.toFixed(1),r.Variance.toFixed(1),r.Zscore.toFixed(2),r.Outlier]),[2600,1700,1700,1700,1660]));
body.push(P(`Over the full period the yield is about ${adrY.Recovery_pct}% for Adrita and ${ankY.Recovery_pct}% for Ankita. The control chart rings one out-of-pattern cycle, Ankita FY2022-23 (10.8%), which sits below the lower limit — the same cycle the mid-term flagged. Most cycles are below the 22% target, which is partly a timing effect because raw nuts bought in one year are processed in the next. This low yield is the single biggest lever on the P&L that follows.`));

body.push(H2("6.4 Profit and Loss"));
body.push(P(`The kernels sell for about ₹${pnl.kernel_price}/kg on average (from the data). With a ${(pnl.assume.outturn*100).toFixed(0)}% outturn, the raw nut cost works out to about ₹${Math.round(base.rcn_cost)}/kg of kernel, and adding the assumed ₹${pnl.assume.processing_per_kg_kernel}/kg processing gives a total cost of about ₹${Math.round(base.total_cost)}/kg. That leaves a contribution of about ₹${Math.round(base.contribution)}/kg, a margin of ${base.margin_pct.toFixed(1)} percent.`));
body.push(tabcap(8,"Profit build-up per kg of kernel (base case)"));
body.push(table(["Item","₹ per kg kernel"],[["Selling price",inr(base.price)],["Raw nut cost (price ÷ outturn)",inr(Math.round(base.rcn_cost))],["Processing & overhead (assumed)",inr(base.proc)],["Total cost",inr(Math.round(base.total_cost))],["Contribution (profit)",inr(Math.round(base.contribution))],["Margin",base.margin_pct.toFixed(1)+"%"]],[5200,4160]));
body.push(...fig("fig05_waterfall.png",520,275,5,"Profit build-up per kg of kernel."));
body.push(...fig("fig06_breakeven.png",520,260,6,"Break-even on the kernel selling price."));
body.push(P(`Figure 6 shows the business breaks even at about ₹${pnl.be_price}/kg; since kernels sell above that, processing is profitable. In outturn terms, break-even is about ${pnl.be_outturn}% — below that yield, processing alone loses money, which is why raising yield matters so much.`));
body.push(tabcap(9,"Contribution by grade (per kg)"));
body.push(table(["Grade","Assumed price (₹/kg)","Contribution (₹/kg)"],[["Whole kernel",inr(pnl.whole_price),inr(Math.round(pnl.grade.Whole.contribution))],["Broken kernel",inr(pnl.broken_price),inr(Math.round(pnl.grade.Broken.contribution))]],[3120,3120,3120]));
body.push(...fig("fig08_grade.png",380,235,8,"Contribution by grade — whole versus broken."));
body.push(P(`The grade split is the key P&L insight. Whole kernels (about ${pnl.whole_share}% of output) earn roughly ₹${Math.round(pnl.grade.Whole.contribution)}/kg, while broken kernels barely cover their share of cost (about ₹${Math.round(pnl.grade.Broken.contribution)}/kg). So the profit comes from keeping breakage low and selling more whole kernels.`));
body.push(...fig("fig07_sensitivity.png",430,300,7,"Sensitivity of contribution to raw nut and selling-price changes (₹/kg)."));
body.push(P(`The sensitivity grid (Figure 7) shows the contribution stays positive for most price moves, but a 10% rise in the raw nut price together with a 10% fall in the selling price would wipe the margin out. This is why buying at the right time and protecting the selling price both matter.`));

body.push(H2("6.5 Buyer Risk and Credit"));
body.push(...fig("fig09_buyer_map.png",560,300,9,"Buyer risk map — revenue versus recency, sized by number of orders."));
body.push(P(`Each bubble is a buyer. The reliable buyers (green) sit at the left (they ordered recently) and high up (they bring more revenue) with bigger bubbles (more orders). The high-risk buyers (red) sit to the right — they have not ordered in a long time — and are mostly small. Of ${by.n_buyers} buyers, ${by.tiers.Low} are Low risk, ${by.tiers.Medium} Medium and ${by.tiers.High} High.`));
body.push(...fig("fig10_tiers.png",520,250,10,"Buyer tiers and recommended credit exposure."));
body.push(tabcap(10,"Buyer tiers, counts and recommended credit exposure"));
body.push(table(["Tier","Buyers","Recommended credit exposure (₹)"],["Low","Medium","High"].map(t=>[t,String(by.tiers[t]),inr(by.exposure[t])]),[3120,3120,3120]));
body.push(tabcap(11,"Top buyers — scorecard (highest revenue)"));
body.push(table(["Buyer","Revenue (₹)","Orders","Recency (d)","Risk","Tier","Credit limit (₹)"],by.top.slice(0,8).map(b=>[b.buyer.length>22?b.buyer.slice(0,22)+"…":b.buyer,inr(b.revenue),String(b.orders),String(b.recency),(+b.risk).toFixed(0),b.tier,inr(b.credit_limit)]),[2600,1500,820,1020,720,1000,1700]));
body.push(...fig("fig11_roc.png",330,295,11,"ROC curve of the illustrative payment-delay model."));
body.push(...fig("fig12_coef.png",520,250,12,"Drivers of payment-delay risk (illustrative model)."));
body.push(P(`The illustrative logistic-regression model reaches an AUC of ${by.lr.auc} and ${(by.lr.acc*100).toFixed(0)}% accuracy on its test set, which shows the method works end to end. Its weights behave sensibly: higher recency raises delay risk, while more orders, higher frequency and longer tenure lower it. This is a demonstration only — once the business records real invoice and payment dates, the same model can be trained on true delays.`));

body.push(H2("6.6 Anomaly Detection"));
body.push(...fig("fig13_anom_scatter.png",560,250,13,"All transactions with anomalies flagged (amount on a log scale)."));
body.push(...fig("fig14_iso_hist.png",520,250,14,"Isolation Forest anomaly-score distribution with the flag threshold."));
body.push(P(`The two methods together flagged ${an.n_total} of ${RF.totals.n_sales+RF.totals.n_purchase} transactions (${an.pct}%) for a closer look — ${an.n_zscore} by the Z-score rule and ${an.n_iso} by Isolation Forest. Most flags are unusually small raw-nut purchases sitting among mostly large ones, which are worth checking for data-entry errors or part-payments.`));
body.push(tabcap(12,"Top flagged transactions for review"));
body.push(table(["Date","Company","Flow","Goods","Party","Amount (₹)","Flagged by"],an.top.slice(0,8).map(a=>[a.Date,a.Company,a.Flow,a.Goods,a.Party,inr(a.Amount),a.Reason]),[1400,1200,1100,820,2200,1340,1300]));
body.push(...fig("fig15_recon.png",560,250,15,"Stock reconciliation — monthly kernel-out ÷ raw-nut-in, odd months ringed."));
body.push(P(`The reconciliation chart flags ${an.recon_months_flagged} months where the kernel-out to raw-nut-in ratio is far from normal. These are usually months with large raw-nut buying but little kernel selling (stock building), but they are exactly the months to check for weight or recording mistakes.`));
body.push(new Paragraph({children:[new PageBreak()]}));

// 7 RECOMMENDATIONS
body.push(H1("7. Final Recommendations"));
[`Lift yield and cut breakage. Yield is the biggest lever on profit; the business loses money below about ${pnl.be_outturn}% outturn. Start batch-wise tracking so low-yield batches are caught early, and aim more output at whole kernels, which carry almost all the margin.`,
`Buy raw nuts at the right time. The raw nut price has risen and swings a lot. Use the season and the price trend to buy more before the usual price rise, and avoid buying at peaks.`,
`Protect the selling price. The margin is thin and a joint adverse move in raw-nut and selling prices can erase it. Hold firm on whole-kernel pricing and sell broken kernels in value-added forms where possible.`,
`Control buyer credit. Set credit limits by tier (Table 10). Keep tight terms for the ${by.tiers.High} high-risk buyers and reward the reliable ones. Most of all, start recording invoice and payment dates so a real payment-delay model can replace the illustrative one.`,
`Act on the anomaly flags. Review the ${an.n_total} flagged transactions and the ${an.recon_months_flagged} odd stock months for entry errors, weight gaps or part-payments, and run these checks monthly.`,
`Build the data habit. The main limit throughout this project was missing detail — no batch records, no cost heads, no payment dates. Capturing these in Tally will unlock per-batch P&L, true ageing and stronger forecasts.`].forEach(t=>body.push(bul(t)));

// 8 CONCLUSION
body.push(H1("8. Conclusion and Limitations"));
body.push(P(`This study turned four years of raw Tally records into a clear picture of the business. Sales are seasonal and slowly rising, processing earns a thin but positive margin that depends heavily on yield and on selling whole kernels, a few buyers carry most of the revenue and need credit control, and a small set of transactions deserve a second look. Acting on the six recommendations above should improve margin, cash flow and control.`));
body.push(P(`Limitations should be stated plainly. The yield is measured per year, not per batch, because batch records are not kept. The P&L uses an assumed processing cost since cost heads are not separated in the ledgers. The payment-delay model is illustrative because payment dates are not recorded. Each of these can be removed once the business captures the extra data, and the methods in this report are ready to use the moment it does.`));

const doc=new Document({
 styles:{default:{document:{run:{font:"Arial",size:22}}},
  paragraphStyles:[
   {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:30,bold:true,font:"Arial",color:TEAL},paragraph:{spacing:{before:240,after:160},outlineLevel:0}},
   {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:25,bold:true,font:"Arial",color:AMBER},paragraph:{spacing:{before:160,after:100},outlineLevel:1}}]},
 numbering:{config:[{reference:"bul",levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:560,hanging:280}}}}]}]},
 sections:[{properties:{page:{size:{width:12240,height:15840},margin:{top:1440,right:1440,bottom:1440,left:1440}}},
   footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Final Report — Ankita Cashew Processing   |   Page ",size:16,color:"888888"}),new TextRun({children:[PageNumber.CURRENT],size:16,color:"888888"})]})]})},
   children:body}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync("Ankita_Cashew_Final_Report.docx",b);console.log("written; toc pages:",Object.keys(pages).length)});
