# Ledger Insight AI — Project Documentation

**Built by Lovenish Purohit**  
**Stack:** Next.js 14, TypeScript, Tailwind CSS, SheetJS (xlsx), Papa Parse  
**Deployed:** Vercel  
**Repo:** https://github.com/lovenishpurohit111/ledger-insight-ai

---

## Overview

A web application that accepts General Ledger files (CSV / XLSX) and generates:
- Financial dashboard (P&L, Balance Sheet, Cash Flow)
- Month-over-Month P&L analysis
- Financial Insights (burn rate, ratios, Pareto, anomalies)
- Downloadable Excel workbook with live formulas linked to raw data
- CSV and PDF exports

---

## Architecture

```
app/
  upload/
    page.tsx              — Main app page (upload + dashboard)
    upload-utils.ts       — File parsing (CSV + QBO Excel), validation
    components/
      FileDropzone.tsx    — Drag & drop upload zone
      ValidationPanel.tsx — Error display
      PreviewTable.tsx    — Transaction preview table
      QBOGuide.tsx        — Interactive QBO export guide (6 steps)
      ColumnMapper.tsx    — Drag-drop column mapping for non-standard files

src/lib/
  accounting.ts           — Account type classification, parsers
  analyzeLedger.ts        — Duplicate detection, inconsistent vendor logic
  generateBalanceSheet.ts — Balance Sheet with current/non-current split
  generateCashFlow.ts     — Cash Flow (indirect method)
  generateInsights.ts     — Burn rate, ratios, Pareto, anomalies, audit flags
  generateMoMPL.ts        — Month-over-Month P&L generator
  generatePL.ts           — P&L with COGS/Gross Profit separation
  exportUtils.ts          — Excel (7 sheets), CSV, PDF export
  verifyCategoryFree.ts   — Free keyword-based category mismatch detection

public/samples/
  sample-ledger.xlsx      — Balanced sample GL (18 transactions, double-entry)
  sample-ledger.csv       — Same data as CSV
```

---

## Key Features

### File Support
- **CSV** — standard format with required headers
- **XLSX** — standard format OR QuickBooks Online General Ledger export
- **QBO format:** Auto-detects header row (up to row 20), skips `Total for X` subtotals
- **Column Mapper:** If headers don't match, shows drag-and-drop mapping UI

### Mandatory Columns
These 3 are strictly required — file is rejected without them:
1. `Distribution account` — account name
2. `Distribution account type` — Asset / Liability / Equity / Income / Expense / Cost of Goods Sold
3. `Transaction type` — Invoice, Payment, Bill, Deposit, etc.

### QBO Export Guide
Interactive 6-step guide on landing page:
1. Open Reports → For My Accountant → General Ledger
2. Click Customize
3. Set Date Range
4. Add Columns → `Distribution account type` ⚠️ (not in QBO by default)
5. Run Report
6. Export to Excel → upload here

---

## Accounting Logic

### P&L (`generatePL.ts`)
- Revenue: Income / Revenue / Sales account types
- COGS: Cost of Goods Sold / COGS types (separated from OpEx)
- Gross Profit = Revenue − COGS
- Operating Expenses: Expense types only (excludes COGS)
- Net Profit = Gross Profit − OpEx
- Monthly breakdown uses `extractMonthKey()` supporting ISO and MM/DD/YYYY dates

### Balance Sheet (`generateBalanceSheet.ts`)
- Uses **last balance per account** (LOOKUP pattern — empty balance = $0 for paid-off accounts)
- CPE (Current Period Earnings) = Net Profit from P&L, only injected if not already in equity accounts
- Current vs Non-Current classification for assets and liabilities
- Financial ratios: Current Ratio, Quick Ratio, Debt-to-Equity, Debt Ratio
- Tolerance: ±$1.00 for `isBalanced` check

### Cash Flow (`generateCashFlow.ts`)
- Indirect method: starts from Net Profit
- Adjusts for working capital changes

### Insights (`generateInsights.ts`)
- Burn rate & cash runway (from bank account balances)
- Top 10 vendors by spend (Pareto)
- Top 10 revenue sources
- Anomaly detection: Z-score >3σ on transaction amounts
- Audit flags: round numbers ≥$500, weekend transactions, sequence gaps
- Tax estimate: 21% flat on net profit
- Category mismatch: `verifyCategoryFree.ts` uses keyword signals

### Inconsistent Vendors (`analyzeLedger.ts`)
- Only flags vendors using multiple accounts **within the same family** (expense+expense, income+income)
- Normal debit/credit pairs (expense + asset) are NOT flagged

---

## Excel Export (`exportUtils.ts`)

### Sheet Structure
| Sheet | Description |
|-------|-------------|
| Dashboard | KPI cards, P&L summary linked to P&L tab, BS snapshot, audit flags |
| Raw Ledger | Full GL data, color-coded by type, auto-filter, frozen headers |
| P & L | SUMIF formulas from Raw Ledger, monthly breakdown with SUMPRODUCT |
| Balance Sheet | LOOKUP last balance per account, reconciliation with IF formula for status |
| Month-over-Month | SUMPRODUCT per cell per month, MoM% row, frozen first col |
| Cash Flow | Net Profit linked from `='P & L'!B15` |
| Flags | Inconsistent vendors + duplicates |

### Formula Approach
- **P&L totals:** `IFERROR(SUMIF('Raw Ledger'!B$2:B$N,"Income",'Raw Ledger'!I$2:I$N)+...)`
- **Monthly:** `IFERROR(SUMPRODUCT(((type=X)>0)*(LEFT(date,7)="YYYY-MM")*ISNUMBER(amt)*amt),0)`
- **Balance Sheet:** `IFERROR(LOOKUP(2,1/(acct=name),balance),0)`
- **CPE:** `='P & L'!B15` (not LOOKUP — CPE doesn't exist in Raw Ledger)
- **BS Status:** `=IF(ABS(variance)<=1,"✔ BALANCED","✘ Off by "&TEXT(...))`
- **Date normalization:** All dates written to Raw Ledger as ISO `YYYY-MM-DD` so `LEFT(date,7)` works

### Design System
- Palette: Charcoal `#1A1F2E`, `#2C3E50`, soft green `#276749`, soft red `#9B2C2C`
- No heavy borders — bottom hairlines on data, medium accent on totals
- Alternating white / `#F0F2F5` rows
- Color-coded Raw Ledger: green=income, red=expense, blue=asset, yellow=liability

### Critical: TDZ Fix
All module-level declarations in `exportUtils.ts` use `var` (not `const`) to prevent Temporal Dead Zone errors in Next.js production minification. Export functions use `async import()` in `page.tsx` for the same reason.

---

## UI Structure

### Upload View
- Two-column layout (large screens): left = dropzone, right = QBO guide + headers
- Full viewport width
- Theme toggle (Dark / Light)

### Dashboard View (after upload)
- Sticky navbar with tabs + export buttons
- Tabs: Overview · 🔍 Insights · P&L · Balance Sheet · Cash Flow · Month-over-Month · Transactions

### Month-over-Month Tab
- Date range filter (From / To month selectors)
- Toggle: `$ Amount` / `% MoM`
- Sticky account column + frozen header row
- Color: green = revenue up / expense down, red = revenue down / expense up

---

## Known Constraints

1. **`Distribution account type` required** — QBO doesn't export this by default. User must add via Customize → Change columns before exporting.
2. **Large files** — SUMPRODUCT formulas use `DEND = rowCount + 1` to bound range. Not using full column references to avoid `#VALUE!`.
3. **Date formats** — Supports ISO (`YYYY-MM-DD`) and US (`MM/DD/YYYY`). Normalized to ISO in Raw Ledger.
4. **CPE double-count guard** — If equity already has a "Net Income" account, CPE is not injected again.

---

## Development Notes

### Git Remote
```bash
git remote set-url origin https://USERNAME:PAT@github.com/lovenishpurohit111/ledger-insight-ai.git
```

### Local Dev
```bash
npm install
npm run dev   # http://localhost:3000/upload
```

### Type Check
```bash
npx tsc --noEmit
```

### Deploy
Auto-deploys to Vercel on push to `main`.

---

## Changelog (Major)

| Feature | Description |
|---------|-------------|
| QBO parser | Auto-detect header row, skip subtotals, infer account type |
| Mandatory column enforcement | Reject files missing Distribution account type |
| Column Mapper | Drag-drop UI for non-standard headers |
| Balance Sheet fix | Last-balance-per-account, empty = $0, CPE from P&L |
| BS Tolerance | ±$1.00 for real-world rounding drift |
| COGS separation | Gross Profit line in P&L |
| MoM date filter | From/To month selectors |
| Insights tab | Burn rate, ratios, Pareto, anomalies, audit flags, tax estimate |
| Category mismatch | Free keyword-based detection, no paid API |
| Excel redesign | 7 sheets, live formulas, modern minimal design |
| TDZ fix | Dynamic import + var declarations for production minification |
| Date normalization | All dates → ISO before writing to Excel |
