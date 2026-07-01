# BDM Capstone — Analytical Study of Ankita Cashew Processing
Akash Jana · Roll 23f2000990 · IIT Madras BS Degree

## Folder structure

```
Source_and_Reference/     Shared inputs (do not edit)
    Sales Purchase 22-26 (1).xlsx   raw sales & purchase ledgers
    Stock Summary.xlsx              raw year-end stock records
    Proposal.pdf                    project proposal
    Proof Of Originality.pdf        signed NOC + proof
    Midterm_Example_Sheet.pdf       reference example
    iitm_logo.png, links.json       used by the report generators

Midterm/                  Mid-term submission
    Ankita_Cashew_Midterm_Report.docx
    Ankita_Adrita_Midterm_Analysis.xlsx
    Cashew_Dashboard.html
    code/                 scripts + intermediates + figures/
        analysis_ankita.py     cleaning, forecasting, yield  ->  figures/, results.json, annual.json
        gen_report.js          builds the report  ->  ../Ankita_Cashew_Midterm_Report.docx
        build_dashboard.py     builds the dashboard  ->  ../Cashew_Dashboard.html

Final/                    Final submission
    Ankita_Cashew_Final_Report.docx / .pdf
    Ankita_Adrita_Final_Analysis_native_charts.xlsx
    Cashew_Dashboard_Final.html
    code/                 scripts + intermediates + figures_final/
        analysis_final.py         P&L, buyer-risk logistic, anomaly detection  ->  figures_final/, results_final.json
        gen_final_report.js       builds the report  ->  ../Ankita_Cashew_Final_Report.docx
        build_final_excel.py      builds the Excel   ->  ../Ankita_Adrita_Final_Analysis_native_charts.xlsx
        build_dashboard_final.py  builds the dashboard  ->  ../Cashew_Dashboard_Final.html

node_modules/, package.json, requirements.txt   shared build tools (kept at root)
```

## How to re-run (from inside a code/ folder)

```
pip install -r ../../requirements.txt      # once
python analysis_ankita.py                   # (or analysis_final.py) regenerates figures + json
node gen_report.js                          # (or gen_final_report.js) rebuilds the report
python build_final_excel.py                 # rebuilds the Excel   (Final only)
python build_dashboard.py                   # rebuilds the dashboard
```

The scripts read the raw data from `../../Source_and_Reference/` and write the
finished report / Excel / dashboard up into the `Midterm/` or `Final/` folder.
Run each script from its own `code/` folder so the relative paths line up.

Notes:
- The final stage reuses two mid-term outputs (`results.json`, `annual.json`); copies are kept in `Final/code/` so it runs on its own.
- `Final/code/Ankita_Adrita_Final_Analysis_OLD_image_charts.xlsx` is the earlier picture-chart version, kept only for reference; the current Excel uses native charts.
- Amounts are in ₹ (INR), weights in KG. Data covers 07-Apr-2022 to 10-Feb-2026.
