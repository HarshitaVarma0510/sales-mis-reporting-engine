# Sales MIS Reporting Engine (Pure JavaScript & ExcelJS)

[![Architecture](https://img.shields.io/badge/Architecture-100%25%20Standalone%20Client--Side-blue.svg)](#)
[![Stack](https://img.shields.io/badge/Stack-Vanilla%20JS%20%7C%20ExcelJS%20%7C%20Chart.js%20%7C%20Tailwind-indigo.svg)](#)
[![Status](https://img.shields.io/badge/Audit-0.00%20Variance%20Reconciled-emerald.svg)](#)
[![Zero Backend](https://img.shields.io/badge/Backend-Zero%20Dependencies%20(No%20Server)-success.svg)](#)

A high-performance, boardroom-ready **Sales MIS (Management Information System) Executive Reporting Engine** engineered entirely in pure client-side modern JavaScript. The engine runs 100% inside any modern web browser with **zero backend servers, zero external runtimes, and zero database dependencies**, providing instant transaction ingestion, automated data cleansing, interactive multi-dimensional dashboards, and institutional 4-sheet financial Excel model compilation.

---

## Key Highlights

- **Zero-Backend Architecture:** Runs locally via standard `file:///` protocol. All file ingestion, cleansing, metric calculations, and spreadsheet compilation execute in transient browser memory via modern Web APIs (`FileReader`, `ArrayBuffer`, and `Blob`).
- **Institutional 4-Sheet ExcelJS Export:** Compiles a boardroom-grade financial workbook (`Sales_MIS_Final_Report.xlsx`) with live native Excel formulas (`=L*M`, `=N*O`, `=N-P`, `=SUM(...)`), Segoe UI typography, corporate palette fills, frozen header panes, and auto-fitted column widths.
- **Certified Data Hygiene:** Automatically purges the critical `ORD-0038` duplicate transaction (row 252 excluded, row 39 preserved), validates positive metric boundaries (`Quantity > 0`, `Unit_Price > 0`), and standardizes dates.
- **0.00 Mathematical Parity:** Complete mathematical tie-out between transactional records and multi-dimensional pivot summaries with zero variance across all top-line and realized metrics.
- **Dynamic DOM Hydration:** Starts in a clean, uninitialized placeholder state (`—`, `0`, `0.0%`) upon page refresh, instantly hydrating and auto-generating the Excel report when the user clicks **"Use Demo Dataset"** or uploads a spreadsheet.
- **Interactive Multi-Dimensional Slicers:** Real-time cross-filtering across Region, Channel, Category, and Order Status, paired with a time-granularity toggle (Daily | Weekly | Monthly) and instant Chart.js updates.

---

## Certified Mathematical Baseline (0.00 Variance)

The engine enforces absolute mathematical reconciliation between the cleaned transactions register and aggregate summary tables:

| Core Financial Metric | Verified Value | Variance | Audit Status |
| :--- | :---: | :---: | :---: |
| **Total Gross Sales** | **₹31,65,485.00** | **₹0.00** | **100% RECONCILED** |
| **Total Discount Amount** | **₹2,06,718.50** | **₹0.00** | **100% RECONCILED** |
| **Total Net Sales (Booked Revenue)** | **₹29,58,766.51** | **₹0.00** | **100% RECONCILED** |
| **Delivered Net Sales (Realized)** | **₹14,04,302.88** | **₹0.00** | **100% RECONCILED** |
| **Total Valid Orders** | **250 Orders** | **0** | **100% RECONCILED** |
| **Total Physical Units Sold** | **655 Units** | **0** | **100% RECONCILED** |
| **Average Order Value (AOV)** | **₹11,835.07** | **₹0.00** | **100% RECONCILED** |
| **Cancellation Rate** | **16.40% (41 orders)** | **0.00%** | **100% RECONCILED** |
| **Return Rate** | **16.40% (41 orders)** | **0.00%** | **100% RECONCILED** |
| **Total Operational Leakage** | **32.80% (82 orders)** | **0.00%** | **100% RECONCILED** |
| **Operational Cash Bleed** | **₹11,33,668.38** | **₹0.00** | **100% RECONCILED** |

---

## 4-Sheet ExcelJS Export Specification

The generated workbook (`Sales_MIS_Final_Report.xlsx`) incorporates professional financial modeling standards:

1. **`Executive_Dashboard` (Active Landing Sheet):**
   - Corporate executive title banner (`#0F172A`), generation timestamp, and certified `0.00 Mathematical Variance Verified` audit badge.
   - 8 KPI cards matching the web portal layout with executive color accents (Delivered Emerald, Warning Amber, Critical Leakage Rose).
   - Strategic Takeaways panel featuring the 5 key management insights.
   - Territory snapshot and fulfillment funnel tables with calculated margin and realization rates.

2. **`Cleaned_Data` (Cleaned Sales Register):**
   - Exactly 250 validated, deduplicated order records across 19 standard columns.
   - Live native Excel calculation formulas:
     * **Gross Sales (Col O):** `=L{row}*M{row}` (`=Quantity * Unit_Price`)
     * **Discount Amount (Col P):** `=O{row}*N{row}` (`=Gross_Sales * Discount_Pct`)
     * **Net Sales (Col Q):** `=O{row}-P{row}` (`=Gross_Sales - Discount_Amount`)
   - Bottom summary row with `=SUM(...)` formulas and double-line accounting border.
   - Frozen header row and native Auto-Filter active on all columns.

3. **`Summary_Tables` (Performance Pivot Matrices):**
   - Multi-dimensional financial pivot tables aggregated by **Region**, **Product Category**, **Sales Representative**, **Sales Channel**, and **Order Status**.
   - Metrics include Order Counts, Units, Gross Sales, Discount, Net Sales, AOV, Order Share %, Revenue Share %, Delivery Realization %, and Leakage %.
   - **Section 6 Mathematical Tie-Out Block** explicitly verifying zero variance between the Cleaned Register and Summary Tables.

4. **`Data_Quality_Log` (Audit Trail & Exception Register):**
   - Complete audit trail documenting the exclusion of `ORD-0038` duplicate at row 252 (preserving row 39), saving ₹15,600.00 in artificial revenue inflation.
   - Documents boundary rule verifications (`Quantity > 0`, `Unit_Price > 0`), zero null checks, and date normalization hashes.

---

## Executive Business Insights

1. **Realized Revenue Gap & Delivery Funnel (₹15.54L Unrealized Pipeline):**
   Out of ₹29,58,766.51 in total booked net sales, only ₹14,04,302.88 (49.6% order realization across 124 orders) has successfully delivered. An unrealized revenue gap of ₹15,54,463.63 (52.5% of total pipeline) remains at risk across pending fulfillment, cancellations, and returns.
2. **Operational Leakage Alert (32.80% Loss Across 82 Orders):**
   Severe operational leakage stands at 32.80%, evenly split between cancellations (16.40% / 41 orders / ₹5,03,529.50) and returns (16.40% / 41 orders / ₹6,30,138.88). Together, 82 compromised orders drain ₹11,33,668.38 in top-line cash flow.
3. **Category Dominance vs Volume Asymmetry (Displays Drives 56.8% Revenue):**
   The 'Displays' category is the primary revenue anchor, generating ₹16,79,959.00 (56.8% of company net sales) across 63 orders with an elite AOV of ₹26,666.02. In contrast, 'Accessories' absorbs 39.2% of order volume (98 orders, 271 units) but contributes only 13.8% of net sales.
4. **Regional Divergence & Expansion Opportunities (North Lags by 42.6%):**
   Central leads the business generating ₹6,88,149.00 (AOV ₹14,336.44), followed closely by East (₹6,54,790.88). In contrast, North significantly lags at ₹3,94,925.50 (42.6% below Central), highlighting a localized logistics or regional distribution bottleneck.
5. **Actionable Management Recommendation: High-Ticket Protocols (>₹15,000):**
   Premium transactions (>₹15,000, predominantly Displays and Furniture) account for over 68% of total return value. Management should immediately enforce mandatory pre-dispatch OTP/telephonic customer confirmation on orders over ₹15,000 and partner with white-glove logistics carriers to reduce return rates by 30–35%, reclaiming up to ₹2,00,000+ monthly.

---

## Clean Repository Structure

```
├── index.html          # Standalone client-side application (Dashboard & Ingestion UI)
├── excelGenerator.js   # Institutional 4-Sheet ExcelJS Financial Model Engine
├── dev_audit.md        # Technical Audit & Mathematical Engineering Specification
└── README.md           # Executive Overview & Quickstart Guide
```

---

## How to Run

### Zero-Setup Instant Run:
1. Locate **`index.html`** in this directory.
2. Double-click **`index.html`** to open it directly in any modern web browser (Google Chrome, Microsoft Edge, Mozilla Firefox, Apple Safari).
3. Click **"Use Demo Dataset"** to immediately test the cleansing engine, view the interactive visual analytics, and auto-download the verified 4-sheet Excel report.
4. Alternatively, drag and drop any raw sales Excel spreadsheet (`.xlsx`) onto the dropzone and click **"Upload Raw Excel & Generate MIS"**.

### Optional Static Web Server:
If you prefer running via a local static web server:

```bash
# Using npx serve or any static HTTP server
npx serve .

# OR using VS Code Live Server extension
```

Then visit `http://localhost:8000` in your browser.
