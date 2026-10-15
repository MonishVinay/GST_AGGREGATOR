# GST Invoice Cumulative Aggregator

A web application designed to batch-process **50 to 150+ individual Excel tax invoices** and consolidate them into a single formatted cumulative report matching the standard Indian GST accounting format.

---

## Functional Modules

The application is unified into a single web portal running on `http://localhost:3000` with direct functional tabs:

### 1. Consolidate Invoices (Batch Aggregator)
- **High-Capacity Batch Processing**: Upload up to **150+ Excel invoices** (`.xlsx`, `.xls`) simultaneously via drag-and-drop.
- **Formatted Cumulative Register**:
  - Matches the standard sales register format with super-header (`OUT INV`, `Period Title`) and columns: `INVOICE NO`, `INV DATE`, `Company Name`, `GST NO`, `HSN Code`, `Qty`, `BASIC AMOUNT`, `SGST`, `CGST`, `IGST`, **`TOTAL AMOUNT`**, `description / remarks`.
  - Line-by-line total amount and `=SUM(K3:K...)` formula for cumulative grand totals.
- **Intelligent GST & SEZ Recognition**:
  - Inter-state supplies split to **IGST (18%)**, intra-state to **CGST (9%) + SGST (9%)**.
  - Identifies SEZ / zero-rated invoices and badges them as **`GST NIL sez`**.
- **Flexible Display & Aggregation Modes**:
  - **Standard View**: Full item breakdown.
  - **Club Same HSN Items**: Merges identical HSN codes on the same invoice.
  - **1 Row per Invoice (Top HSN)**: Collapses each invoice to single row under highest amount HSN.
- **Data Integrity & Consistency Audit**:
  - Real-time audit for missing sequence numbers (e.g. TS-01...TS-50), incomplete mandatory fields, and duplicate invoices.
- **1-Click Transfer**: Seamlessly send consolidated data directly to the **HSN-wise GST Summary Generator**.

### 2. HSN-wise GST Summary Generator
- **Modified Cumulative Sheet Processing**:
  - Accepts cumulative sales registers generated from Module 1, even after manual modifications (rows added or deleted).
- **Exact Reference Style & Workbook Formatting**:
  - **Navy Blue Header** (`#162E4A`) matching corporate GST summary guidelines.
  - **Light Blue Grand Total Row** (`#DDEBF7`) with `=SUM(...)` formulas.
- **HSN Grouping & Business Rules**:
  - Identical HSN codes are grouped together with summed quantities, basic amounts, SGST, CGST, and IGST.
  - **Zero GST Separation**: `Zero GST` (SEZ / NIL / Export) supplies are separated from standard `Taxable (SGST / CGST / IGST)` supplies even under identical HSN codes.
  - **9987XX Service Code Unification**: All service HSN codes starting with `9987` (e.g. `998711`, `998719`, `998799`) are unified under `998719`.
  - **Safe Total Row Exclusion**: Original cumulative sheet "Total" / "Grand Total" rows are safely identified and excluded to prevent double-counting.
- **Built-in Live Data Integrity & Discrepancy Reporting**:
  - **Missing Required Data**: Flags rows with missing HSN codes, non-standard HSN formats, missing quantities, or missing basic values.
  - **Tax Discrepancies**: Flags intra-state SGST != CGST mismatches, dual-tax errors (both IGST and SGST/CGST on the same row), and unclarified zero tax entries.
  - **Sequence Gap Detection**: Identifies any missing invoice numbers across the cumulative register.
  - **Consolidation Log**: Transparent log of which 9987XX codes were unified into 998719 and which total rows were excluded.

### 3. GST Portal Reconciliation & Filing (Coming Soon)
- Direct comparison with GSTR-1, GSTR-2B, and GSTR-3B portal returns.

---

## 1. Running Locally on Your Laptop (Windows)

You can run this server on your laptop immediately:

1. **Option A (Simplest - 1-Click)**:
   - Double-click the [`start.bat`](file:///c:/Users/Tagore/Desktop/New%20folder%20%282%29/start.bat) file in this directory.
   - It will start the server and automatically launch your browser at `http://localhost:3000`.

2. **Option B (PowerShell / Terminal)**:
   ```powershell
   npm start
   ```
   or if Node is not in your global system PATH:
   ```powershell
   .\node-v20.18.0-win-x64\node.exe server.js
   ```
   Then open `http://localhost:3000` in your web browser.

---

## 2. Deploying on Render (Free Cloud Hosting)

To host this website online on [Render](https://render.com) so anyone can access it via a web link:

1. **Create a GitHub Repository**:
   - Push this project folder to your GitHub:
     ```bash
     git init
     git add .
     git commit -m "GST Invoice Cumulative Aggregator"
     git branch -M main
     git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
     git push -u origin main
     ```

2. **Deploy on Render**:
   - Sign in to [Render.com](https://render.com).
   - Click **New +** &rarr; Select **Web Service**.
   - Connect your GitHub repository.
   - Render will automatically detect the settings:
     - **Runtime:** `Node`
     - **Build Command:** `npm install`
     - **Start Command:** `npm start`
     - **Instance Type:** `Free`
   - Click **Deploy Web Service**.

Render will build and start your application in ~1 minute, providing you with a free live `https://...onrender.com` URL.

---

## Project Structure

```
├── public/
│   ├── index.html       # Web Dashboard interface
│   ├── style.css        # Spreadsheet-styled UI
│   └── app.js           # Client upload, table rendering & export logic
├── src/
│   ├── parser.js        # Multi-sheet Excel invoice extractor
│   └── excelGenerator.js# Formatted .xlsx & .csv report builder
├── sample_invoices/     # Test invoices matching actual invoice cases
├── server.js            # Express API server (up to 250 files)
├── render.yaml          # Render cloud blueprint
├── package.json         # Dependencies & scripts
└── start.bat            # Windows 1-click launcher
```
