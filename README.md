# Ankita Cashew Processing — BDM Mid-term Analysis

This folder holds the code, data and outputs for the mid-term report
**"Analytical Study of Ankita Cashew Processing"** (Akash Jana, Roll 23f2000990).

There are two programs:

| File | Language | What it does |
|------|----------|--------------|
| `analysis_ankita.py` | Python | Cleans the data, runs all the analysis (descriptive stats, demand forecasting, inventory yield, ABC), and draws every chart. |
| `gen_report.js` | Node.js | Builds the Word report (`.docx`) from the Python outputs. |

Run the **Python file first**, then the **Node file**. The Node file needs the
files that Python produces (`results.json`, `annual.json`, and the `figures/` folder).

---

## 1. Input files needed

Keep these in the same folder as the scripts:

| File | Used by | Note |
|------|---------|------|
| `Sales Purchase 22-26 (1).xlsx` | Python | Raw sales & purchase ledgers (source data) |
| `Stock Summary.xlsx` | Python | Raw year-end stock records (source data) |
| `iitm_logo.png` | Node | Logo shown on the report title page |
| `toc_pages.json` | Node | Page numbers for the Contents page (already filled) |
| `links.json` | Node | Portal video / letter links (optional — see Step 4) |

---

## 2. Software you need

- **Python 3.10 or newer**
- **Node.js 18 or newer**
- (Optional) **LibreOffice** — only if you want to turn the `.docx` into a PDF.

Check they are installed:

```bash
python --version
node --version
```

---

## 3. Run the Python analysis

Step 3a — install the Python libraries (one time):

```bash
pip install -r requirements.txt
```

Step 3b — run the script:

```bash
python analysis_ankita.py
```

This reads the two source `.xlsx` files and writes:

| Output | What it is |
|--------|-----------|
| `results.json` | Every computed number (totals, stats, forecast, yield, ABC) |
| `annual.json` | Sales and purchase per financial year |
| `cleaning_log.txt` | The data-cleaning counts (rows scanned, duplicates removed, etc.) |
| `figures/fig01..fig13_*.png` | All 13 charts used in the report |

You will see the cleaning counts and forecast/yield numbers printed on screen.

---

## 4. Build the Word report (Node)

Step 4a — install the Node library (one time):

```bash
npm install
```

(this installs `docx`, listed in `package.json`.)

Step 4b — (optional) add the portal links. Open `links.json` and paste the URLs:

```json
{ "video": "https://...", "letter": "https://..." }
```

If you leave them blank, the report keeps a placeholder you can fill in Word later.

Step 4c — build the report:

```bash
node gen_report.js
```

This writes **`Ankita_Cashew_Midterm_Report.docx`**.

---

## 5. (Optional) Make a PDF

If LibreOffice is installed:

```bash
soffice --headless --convert-to pdf Ankita_Cashew_Midterm_Report.docx
```

---

## 6. Quick run (all steps together)

```bash
pip install -r requirements.txt
python analysis_ankita.py
npm install
node gen_report.js
```

---

## 7. Other files in this folder

- **`Ankita_Adrita_Midterm_Analysis.xlsx`** — the Excel workbook. It holds the
  cleaned data, all tables, and the **same charts** as the report (the chart
  images come from the `figures/` folder, so the report and the Excel match).
- **`requirements.txt`** — Python library list.
- **`package.json`** — Node library list.

---

## 8. Notes

- The Contents page numbers come from `toc_pages.json`. They are correct for the
  current report. If you change the report content a lot and the layout shifts,
  the page numbers may need to be set again before the final version.
- All money is in Indian Rupees (₹). All weights are in kilograms (KG).
- The data covers 07 April 2022 to 10 February 2026.
