import * as XLSX from 'xlsx';
import type { LedgerAnalysis } from './analyzeLedger';
import type { BalanceSheet } from './generateBalanceSheet';
import type { CashFlowStatement } from './generateCashFlow';
import type { ProfitAndLoss } from './generatePL';
import type { MoMPL } from './generateMoMPL';
import { monthLabel } from './generateMoMPL';
import type { LedgerRow } from '../../app/upload/upload-utils';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type CS = any;

const $f = (n: number) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(n);
const toNum = (s: string): number => {
  if (!s) return 0;
  const neg = s.trim().startsWith('(') || s.trim().startsWith('-');
  const n = parseFloat(s.replace(/[^0-9.]/g, ''));
  return isNaN(n) ? 0 : neg ? -n : n;
};
const pctStr = (n: number, d = 1) => `${(n * 100).toFixed(d)}%`;
const base = (f: string) => f.replace(/\.[^.]+$/, '') || 'ledger';

// ── Palette (matches the dashboard image feel) ────────────────────────────────
const C = {
  // Blues / Navy
  NAVY:   '1F3864', NAVY2: '2E74B5', NAVY3: '2F5496', NAVY_LT: 'D6DCE4', BLUE_LT: 'BDD7EE',
  // Greens
  GREEN:  '375623', GREEN2:'70AD47', GRN_LT:'E2EFDA', GRN_HDR:'548235',
  // Reds
  RED:    'C00000', RED2:  'FF0000', RED_LT: 'FFDCDC',
  // Ambers / Gold
  AMBER:  'ED7D31', GOLD:  'BF8F00', YLW:   'FFFF00', YLW_LT: 'FFF2CC',
  // Grays / Neutrals
  WHITE:  'FFFFFF', OFF:   'F2F2F2', LGRAY: 'D9D9D9', GRAY:   'A6A6A6', DGRAY:  '595959',
  BLACK:  '000000',
  // Status colours (like image)
  URGENT:  'FF0000', HIGH:    'FFC000', MEDIUM:  '92D050', LOW:     '00B0F0',
  TODO:    'FF7070', INPROG:  'FFEB9C', DONE:    'C6EFCE',
  // Accent
  TEAL:   '1F6B75', PURPLE: '7030A0',
};

// ── Style factory ─────────────────────────────────────────────────────────────
const F = (bold: boolean, sz: number, rgb: string, italic = false, strike = false) =>
  ({ bold, sz, name: 'Calibri', color: { rgb }, italic, strike });
const Fill = (rgb: string): CS => ({ fgColor: { rgb }, patternType: 'solid' });
const Aln  = (h: string, v = 'center', wrap = false): CS => ({ horizontal: h, vertical: v, wrapText: wrap });
const Bdr  = (style: string, rgb: string) => ({ style, color: { rgb } });

const bdrBox  = (c = C.LGRAY): CS => ({ top: Bdr('thin', c), bottom: Bdr('thin', c), left: Bdr('thin', c), right: Bdr('thin', c) });
const bdrData = (): CS => ({ bottom: Bdr('hair', C.LGRAY), right: Bdr('hair', C.LGRAY) });
const bdrTop  = (): CS => ({ top: Bdr('medium', C.LGRAY), bottom: Bdr('double', C.LGRAY) });

const FMT_MONEY = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"_);_(@_)';
const FMT_MONEY0= '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)';
const FMT_PCT   = '0.0%';
const FMT_PCTD  = '+0.0%;[Red]-0.0%;"—"';
const FMT_NUM   = '#,##0';

// Style presets
const ss = {
  // Headers
  hdrC: (bg: string, fg = C.WHITE, sz = 11): CS => ({ font: F(true, sz, fg), fill: Fill(bg), alignment: Aln('center'), border: bdrBox(bg) }),
  hdrL: (bg: string, fg = C.WHITE, sz = 11): CS => ({ font: F(true, sz, fg), fill: Fill(bg), alignment: Aln('left', 'center'), border: bdrBox(bg) }),
  // Section
  secH: (bg: string, fg = C.WHITE): CS => ({ font: F(true, 10, fg), fill: Fill(bg), alignment: Aln('left', 'center') }),
  // Data cells
  c: (bold = false, fg = C.BLACK, ha = 'left', bg?: string): CS => ({
    font: F(bold, 10, fg), ...(bg ? { fill: Fill(bg) } : {}),
    alignment: Aln(ha, 'center'), border: bdrData(),
  }),
  m: (bold = false, fg = C.BLACK, bg?: string): CS => ({
    font: F(bold, 10, fg), ...(bg ? { fill: Fill(bg) } : {}),
    numFmt: FMT_MONEY, alignment: Aln('right', 'center'), border: bdrData(),
  }),
  p: (bold = false, fg = C.BLACK, bg?: string): CS => ({
    font: F(bold, 10, fg), ...(bg ? { fill: Fill(bg) } : {}),
    numFmt: FMT_PCT, alignment: Aln('right', 'center'), border: bdrData(),
  }),
  // Totals
  tL: (bg: string, fg = C.WHITE): CS => ({ font: F(true, 11, fg), fill: Fill(bg), alignment: Aln('left', 'center'),  border: bdrTop() }),
  tM: (bg: string, fg = C.WHITE): CS => ({ font: F(true, 11, fg), fill: Fill(bg), numFmt: FMT_MONEY, alignment: Aln('right', 'center'), border: bdrTop() }),
  tP: (bg: string, fg = C.WHITE): CS => ({ font: F(true, 11, fg), fill: Fill(bg), numFmt: FMT_PCT,   alignment: Aln('right', 'center'), border: bdrTop() }),
  // Special: KPI card value
  kpiV: (fg = C.NAVY, sz = 22): CS => ({ font: F(true, sz, fg), alignment: Aln('center', 'center'), border: bdrBox(C.LGRAY) }),
  kpiL: (): CS => ({ font: F(false, 9, C.DGRAY, true), alignment: Aln('center', 'bottom'), fill: Fill(C.OFF) }),
  // Title/subtitle
  titl: (bg: string, sz = 18): CS => ({ font: F(true, sz, C.WHITE), fill: Fill(bg), alignment: Aln('left', 'center') }),
  sub:  (bg: string): CS => ({ font: F(false, 9, C.NAVY_LT, true), fill: Fill(bg), alignment: Aln('left', 'center') }),
  note: (): CS => ({ font: F(false, 9, C.DGRAY, true), alignment: Aln('left', 'center') }),
  // Status pills
  pill: (bg: string, fg = C.BLACK): CS => ({ font: F(true, 9, fg), fill: Fill(bg), alignment: Aln('center', 'center'), border: bdrBox(bg) }),
};

// ── Helpers ───────────────────────────────────────────────────────────────────
const L  = (n: number) => n < 26 ? String.fromCharCode(65 + n) : String.fromCharCode(64 + Math.floor(n / 26)) + String.fromCharCode(65 + n % 26);
const MG = (r: number, c: number, r2: number, c2: number) => ({ s: { r, c }, e: { r: r2, c: c2 } });
const W  = (n: number) => ({ wch: n });

const wv = (ws: CS, r: number, c: number, v: string | number, t: string, style: CS) => { ws[`${L(c)}${r}`] = { v, t, s: style }; };
const wf = (ws: CS, r: number, c: number, f: string, v: number, style: CS) => { ws[`${L(c)}${r}`] = { t: 'n', f, v, s: style }; };
const bg = (ws: CS, r: number, fromC: number, toC: number, rgb: string) => {
  for (let c = fromC; c <= toC; c++) { const a = `${L(c)}${r}`; if (!ws[a]) ws[a] = { t: 's', v: '' }; ws[a].s = { fill: Fill(rgb) }; }
};
const setRef = (ws: CS, maxR: number, maxC: number) => { ws['!ref'] = `A1:${L(maxC)}${maxR}`; };

// ── Formula constants ─────────────────────────────────────────────────────────
const RL = 'Raw Ledger';
let DEND = 5000;

const TINC  = ['Income','income','Revenue','revenue','Sales','sales','Other Income','other income'];
const TCOGS = ['Cost of Goods Sold','cost of goods sold','COGS','cogs','Cost of Sales'];
const TEXP  = ['Expense','expense','Expenses','expenses','Other Expense','other expense'];
const TAST  = ['Asset','asset','Bank','bank','Accounts Receivable (A/R)','Other Current Assets','Fixed Assets','Other Assets','Inventory'];
const TLIA  = ['Liability','liability','Accounts Payable (A/P)','Credit Card','Other Current Liabilities','Long Term Liabilities','Other Liability'];
const TEQ   = ['Equity','equity','Retained Earnings','Opening Balance Equity'];

// Clean SUMIF chain — one per type string, all chained with +
const SUMIF_chain = (types: string[], col: 'I' | 'J') =>
  `IFERROR(${types.map(t => `SUMIF('${RL}'!B$2:B$${DEND},"${t}",'${RL}'!${col}$2:${col}$${DEND})`).join('+')},0)`;

// SUMPRODUCT for type group + month — IFERROR on TEXT prevents #VALUE on empty/text date cells
const SUMPRODUCT_month = (types: string[], mk: string) => {
  const typeFilter = types.map(t => `('${RL}'!B$2:B$${DEND}="${t}")`).join('+');
  return `IFERROR(SUMPRODUCT(((${typeFilter})>0)*(IFERROR(TEXT('${RL}'!C$2:C$${DEND},"YYYY-MM"),"")="${mk}")*ISNUMBER('${RL}'!I$2:I$${DEND})*('${RL}'!I$2:I$${DEND})),0)`;
};

// SUMPRODUCT for specific account + month
const SUMPRODUCT_acct = (acct: string, mk: string) =>
  `IFERROR(SUMPRODUCT(('${RL}'!A$2:A$${DEND}="${acct.replace(/"/g, '""')}")*(IFERROR(TEXT('${RL}'!C$2:C$${DEND},"YYYY-MM"),"")="${mk}")*ISNUMBER('${RL}'!I$2:I$${DEND})*('${RL}'!I$2:I$${DEND})),0)`;

// LOOKUP for last balance
const LOOKUP_bal = (acct: string) =>
  `IFERROR(LOOKUP(2,1/('${RL}'!A$2:A$${DEND}="${acct.replace(/"/g, '""')}"),'${RL}'!J$2:J$${DEND}),0)`;

// ═════════════════════════════════════════════════════════════════════════════
// CSV
// ═════════════════════════════════════════════════════════════════════════════
export function exportCsv(fileName: string, analysis: LedgerAnalysis, pl: ProfitAndLoss, bs: BalanceSheet, cf: CashFlowStatement): void {
  const e = (v: string | number) => { const s = String(v); return s.includes(',') ? `"${s}"` : s; };
  const sec = (t: string, h: string[], rows: (string | number)[][]) =>
    [t, h.map(e).join(','), ...rows.map(r => r.map(e).join(',')), ''].join('\n');
  const parts = [
    sec('P&L', ['Item','Amount','Margin'],[
      ['Revenue',$f(pl.totalRevenue),'100%'],
      ['COGS',$f(pl.totalCogs),pctStr(pl.totalRevenue?pl.totalCogs/pl.totalRevenue:0)],
      ['Gross Profit',$f(pl.grossProfit),pctStr(pl.grossMargin)],
      ['OpEx',$f(pl.totalExpenses),pctStr(pl.totalRevenue?pl.totalExpenses/pl.totalRevenue:0)],
      ['Net Profit',$f(pl.netProfit),pctStr(pl.netMargin)],
    ]),
    '\n',
    sec('Balance Sheet',['Category','Account','Balance'],[
      ...bs.assets.map(({account,value})=>['Asset',account,$f(value)]),
      ...bs.liabilities.map(({account,value})=>['Liability',account,$f(value)]),
      ...bs.equity.map(({account,value})=>['Equity',account,$f(value)]),
    ]),
  ];
  const blob = new Blob([parts.join('\n')], { type: 'text/csv' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `${base(fileName)}.csv`; a.click();
}

// ═════════════════════════════════════════════════════════════════════════════
// EXCEL
// ═════════════════════════════════════════════════════════════════════════════
export function exportExcel(
  fileName: string, analysis: LedgerAnalysis, pl: ProfitAndLoss,
  bs: BalanceSheet, cf: CashFlowStatement, mom?: MoMPL, rawRows?: LedgerRow[],
): void {
  const wb = XLSX.utils.book_new();
  const gd = new Date().toLocaleDateString('en-US', { year:'numeric', month:'long', day:'numeric' });
  const co = base(fileName);
  const rows = rawRows ?? [];
  DEND = rows.length > 1 ? rows.length + 1 : 5000;

  const HDRS = ['Distribution account','Distribution account type','Transaction date','Transaction type','Num','Name','Description','Split','Amount','Balance'];

  // ══ Sheet 1: DASHBOARD ════════════════════════════════════════════════════
  {
    const ws: CS = {};

    // ── Banner row 1 (full width A1:L1) ──
    for (let c = 0; c <= 11; c++) { ws[`${L(c)}1`] = { v: '', t: 's', s: { fill: Fill(C.NAVY) } }; }
    wv(ws, 1, 0, `📊  FINANCIAL ANALYSIS DASHBOARD`, 's', { ...ss.titl(C.NAVY, 20), font: F(true, 20, C.WHITE) });
    wv(ws, 2, 0, `${co}`, 's', { ...ss.sub(C.NAVY), font: F(true, 12, C.BLUE_LT) });
    wv(ws, 2, 4, `Generated: ${gd}`, 's', ss.sub(C.NAVY));
    wv(ws, 2, 8, `Transactions: ${analysis.totalTransactions.toLocaleString()}`, 's', ss.sub(C.NAVY));
    for (let c = 1; c <= 11; c++) { ws[`${L(c)}2`] = ws[`${L(c)}2`] ?? { v:'', t:'s', s:ss.sub(C.NAVY) }; }

    // ── Row 3: separator ──
    for (let c = 0; c <= 11; c++) { ws[`${L(c)}3`] = { v:'', t:'s', s:{ fill: Fill(C.NAVY2) } }; }

    // ── KPI CARDS (rows 5-8, groups of 2 cols each) ──
    const kpis = [
      { label:'Total Revenue',   val: pl.totalRevenue,   fmt: FMT_MONEY0, fg: C.GRN_HDR, },
      { label:'Gross Profit',    val: pl.grossProfit,    fmt: FMT_MONEY0, fg: C.GREEN2,  },
      { label:'Net Profit',      val: pl.netProfit,      fmt: FMT_MONEY0, fg: pl.netProfit>=0?C.GRN_HDR:C.RED, },
      { label:'Gross Margin',    val: pl.grossMargin,    fmt: FMT_PCT,    fg: C.NAVY2,   },
      { label:'Net Margin',      val: pl.netMargin,      fmt: FMT_PCT,    fg: C.NAVY2,   },
    ];

    kpis.forEach((kpi, i) => {
      const col = i * 2;
      const labelStyle: CS = {
        font: F(true, 9, C.DGRAY),
        fill: Fill(C.OFF),
        alignment: Aln('center', 'bottom'),
        border: bdrBox(C.LGRAY),
      };
      const valStyle: CS = {
        font: F(true, 18, kpi.fg),
        fill: Fill(C.WHITE),
        numFmt: kpi.fmt,
        alignment: Aln('center', 'center'),
        border: { ...bdrBox(C.LGRAY), bottom: Bdr('medium', kpi.fg) },
      };
      wv(ws, 5, col, kpi.label.toUpperCase(), 's', labelStyle);
      ws[`${L(col)}5`].s.border = { ...ws[`${L(col)}5`].s.border, top: Bdr('medium', kpi.fg), left: Bdr('medium', kpi.fg), right: Bdr('medium', kpi.fg) };
      wv(ws, 6, col, typeof kpi.val === 'number' ? kpi.val : 0, 'n', valStyle);
      wv(ws, 7, col, '', 's', { fill: Fill(C.WHITE) }); // spacer
    });

    // ── BS STATUS card (rightmost) ──
    const bsOK = bs.isBalanced;
    wv(ws, 5, 10, 'BALANCE SHEET', 's', { font: F(true, 9, C.DGRAY), fill: Fill(C.OFF), alignment: Aln('center', 'bottom'), border: { top: Bdr('medium', bsOK?C.GRN_HDR:C.RED), left: Bdr('medium', bsOK?C.GRN_HDR:C.RED), right: Bdr('medium', bsOK?C.GRN_HDR:C.RED) } });
    wv(ws, 6, 10, bsOK ? '✔  BALANCED' : '✘  NOT BALANCED', 's', {
      font: F(true, 14, bsOK ? C.GREEN : C.RED),
      fill: Fill(bsOK ? C.GRN_LT : C.RED_LT),
      alignment: Aln('center', 'center'),
      border: { ...bdrBox(bsOK ? C.GRN_HDR : C.RED), bottom: Bdr('medium', bsOK ? C.GRN_HDR : C.RED) },
    });

    // ── SECTION: P&L SUMMARY (rows 9-18) ──
    const PL_START = 10;
    wv(ws, PL_START, 0, '  PROFIT & LOSS SUMMARY', 's', ss.hdrL(C.NAVY, C.WHITE, 11));
    wv(ws, PL_START, 4, 'Formula Source', 's', ss.hdrC(C.NAVY2, C.WHITE, 9));
    bg(ws, PL_START, 1, 3, C.NAVY);
    bg(ws, PL_START, 5, 11, C.NAVY);

    const plRows: [string, number, string, string][] = [
      ['Revenue',        pl.totalRevenue,   C.GRN_LT,  `=SUMIF('${RL}'!B:B,"Income",'${RL}'!I:I)+SUMIF('${RL}'!B:B,"Sales",'${RL}'!I:I)`],
      ['Cost of Goods',  pl.totalCogs,      C.OFF,      `=SUMIF('${RL}'!B:B,"Cost of Goods Sold",'${RL}'!I:I)`],
      ['Gross Profit',   pl.grossProfit,    C.GRN_LT,  `=B${PL_START+1}-B${PL_START+2}`],
      ['Operating Exp',  pl.totalExpenses,  C.OFF,      `=SUMIF('${RL}'!B:B,"Expense",'${RL}'!I:I)+SUMIF('${RL}'!B:B,"Expenses",'${RL}'!I:I)`],
      ['Net Profit',     pl.netProfit,      pl.netProfit>=0?C.GRN_LT:C.RED_LT, `=B${PL_START+3}-B${PL_START+4}`],
    ];

    plRows.forEach(([label, val, rowBg, formula], i) => {
      const r = PL_START + 1 + i;
      const isBold = label === 'Gross Profit' || label === 'Net Profit';
      const fg = label === 'Net Profit' ? (val >= 0 ? C.GREEN : C.RED) : C.BLACK;
      wv(ws, r, 0, `  ${label}`, 's', ss.c(isBold, fg, 'left', rowBg));
      ws[`${L(0)}${r}`].s.border = { left: Bdr('medium', C.NAVY), ...bdrData() };
      wv(ws, r, 1, val, 'n', { ...ss.m(isBold, fg, rowBg), numFmt: FMT_MONEY });
      wv(ws, r, 2, pl.totalRevenue ? val / pl.totalRevenue : 0, 'n', ss.p(isBold, fg, rowBg));
      wv(ws, r, 3, formula, 's', { font: F(false, 8, C.NAVY3, true), fill: Fill(C.BLUE_LT), alignment: Aln('left', 'center'), numFmt:'@' });
      ws[`${L(3)}${r}`].s.border = { right: Bdr('medium', C.NAVY), ...bdrData() };
    });

    // ── SECTION: Balance Sheet snapshot (rows 10-18, cols 5-11) ──
    wv(ws, PL_START, 5, '  BALANCE SHEET SNAPSHOT', 's', ss.hdrL(C.TEAL, C.WHITE, 11));
    bg(ws, PL_START, 6, 11, C.TEAL);

    const bsSections = [
      { label:'ASSETS',      val:bs.totals.assetsTotal,      color:C.NAVY2 },
      { label:'LIABILITIES', val:bs.totals.liabilitiesTotal, color:C.TEAL },
      { label:'EQUITY',      val:bs.totals.equityTotal,      color:C.GREEN },
      { label:'VARIANCE',    val:bs.variance,                color:bsOK?C.GRN_HDR:C.RED },
    ];

    bsSections.forEach(({ label, val, color }, i) => {
      const r = PL_START + 1 + i;
      wv(ws, r, 5, label, 's', ss.c(true, color, 'left', C.OFF));
      wv(ws, r, 6, val, 'n', { ...ss.m(true, color), numFmt: FMT_MONEY, ...(i===3?{fill:Fill(bsOK?C.GRN_LT:C.RED_LT)}:{}) });
    });

    // ── SECTION: Top Accounts row 20-26 ──
    const TA_ROW = 20;
    wv(ws, TA_ROW, 0, '  TOP 5 EXPENSE ACCOUNTS', 's', ss.hdrL(C.RED, C.WHITE, 10));
    bg(ws, TA_ROW, 1, 3, C.RED);
    wv(ws, TA_ROW, 4, '  TOP 5 INCOME ACCOUNTS', 's', ss.hdrL(C.GRN_HDR, C.WHITE, 10));
    bg(ws, TA_ROW, 5, 7, C.GRN_HDR);

    const topExp = bs.liabilities.slice(0,5);
    const incAccts = bs.assets.slice(0,5);

    // Top expense from analysis
    analysis.inconsistentVendors.slice(0,5).forEach((v, i) => {
      const r = TA_ROW + 1 + i; const rb = i%2===0?C.WHITE:C.OFF;
      wv(ws, r, 0, v.vendor, 's', ss.c(false, C.BLACK, 'left', rb));
      wv(ws, r, 1, v.reason, 's', ss.c(false, C.DGRAY, 'left', rb));
    });
    bs.liabilities.slice(0,5).forEach((e, i) => {
      const r = TA_ROW + 1 + i; const rb = i%2===0?C.WHITE:C.OFF;
      wv(ws, r, 4, e.account, 's', ss.c(false, C.BLACK, 'left', rb));
      wv(ws, r, 5, e.value, 'n', ss.m(false, C.BLACK, rb));
    });

    // ── FLAGS ROW ──
    const FL = 28;
    wv(ws, FL, 0, `⚠  ${analysis.inconsistentVendors.length} Inconsistent Vendors`, 's', ss.pill(analysis.inconsistentVendors.length>0?C.YLW:C.GRN_LT, analysis.inconsistentVendors.length>0?C.BLACK:C.GREEN));
    wv(ws, FL, 2, `⚠  ${analysis.duplicates.length} Duplicate Transactions`, 's', ss.pill(analysis.duplicates.length>0?C.RED_LT:C.GRN_LT, analysis.duplicates.length>0?C.RED:C.GREEN));
    wv(ws, FL, 4, `📁  ${rows.length.toLocaleString()} Total Rows in Raw Ledger`, 's', ss.pill(C.BLUE_LT, C.NAVY));
    wv(ws, FL, 6, `📅  Period: ${mom?.months?.[0] ?? '—'} → ${mom?.months?.slice(-1)[0] ?? '—'}`, 's', ss.pill(C.NAVY_LT, C.NAVY));

    ws['!ref'] = `A1:L${FL + 2}`;
    ws['!cols'] = [W(22),W(16),W(12),W(36),W(22),W(16),W(12),W(12),W(12),W(12),W(22),W(12)];
    ws['!rows'] = [{ hpt:36 },{hpt:20},{hpt:8},{},{hpt:20},{hpt:44},{hpt:8},{},{hpt:24}];
    ws['!merges'] = [
      MG(0,0,0,11), // title banner
      MG(1,0,1,3), MG(1,4,1,7), MG(1,8,1,11), // subtitle cells
      MG(2,0,2,11), // separator
      // KPI cards rows 5-6 (5 cards × 2 cols = cols 0-9, BS status at cols 10-11)
      ...Array.from({length:5},(_,i)=>MG(4,i*2,4,i*2+1)),
      ...Array.from({length:5},(_,i)=>MG(5,i*2,5,i*2+1)),
      MG(4,10,4,11), // BS status label
      MG(5,10,5,11), // BS status value
      // P&L section header (no overlapping merges)
      MG(9,0,9,3), MG(9,4,9,11),
      // Monthly section header
      MG(19,0,19,3), MG(19,4,19,7),
      // Flag pills
      MG(27,0,27,1), MG(27,2,27,3), MG(27,4,27,5), MG(27,6,27,7),
    ];
    ws['!freeze'] = { xSplit:0, ySplit:4 };
    XLSX.utils.book_append_sheet(wb, ws, 'Dashboard');
  }

  // ══ Sheet 2: RAW LEDGER ═══════════════════════════════════════════════════
  {
    const ws: CS = {};

    // Banner
    for (let c=0;c<10;c++) ws[`${L(c)}1`] = {v:'',t:'s',s:{fill:Fill(C.NAVY)}};
    wv(ws, 1, 0, `RAW LEDGER  —  ${co}  (${rows.length.toLocaleString()} transactions)`, 's', ss.titl(C.NAVY, 12));

    // Headers row 2
    const hdrColors = [C.NAVY,C.NAVY3,C.TEAL,C.TEAL,C.GRAY,C.DGRAY,C.DGRAY,C.GRAY,C.GRN_HDR,C.GRN_HDR];
    HDRS.forEach((h, i) => wv(ws, 2, i, h, 's', ss.hdrC(hdrColors[i] ?? C.NAVY, C.WHITE, 9)));

    // Data rows — amounts and balances as NUMBERS
    rows.forEach((row, ri) => {
      const rn = ri + 3;
      const rb = ri%2===0 ? C.WHITE : C.OFF;
      HDRS.forEach((h, ci) => {
        const raw = row[h as keyof typeof row] ?? '';
        const isAmt = ci===8||ci===9;
        const isDate = ci===2;
        if (isAmt) {
          const num = toNum(String(raw));
          wv(ws, rn, ci, num, 'n', { ...ss.m(false,num<0?C.RED:C.BLACK,rb), numFmt:FMT_MONEY });
        } else {
          wv(ws, rn, ci, String(raw), 's', ss.c(false,C.BLACK,'left',rb));
          if(isDate) ws[`${L(ci)}${rn}`].s.alignment = Aln('center','center');
        }
      });
      // Colour code account type column (col B = index 1)
      const typeVal = String(row['Distribution account type'] ?? '').toLowerCase();
      let typeBg = rb;
      if(typeVal.includes('income')||typeVal.includes('revenue')||typeVal.includes('sales')) typeBg = C.GRN_LT;
      else if(typeVal.includes('expense')||typeVal.includes('cost')) typeBg = C.RED_LT;
      else if(typeVal.includes('asset')||typeVal.includes('bank')) typeBg = C.BLUE_LT;
      else if(typeVal.includes('liabilit')||typeVal.includes('payable')||typeVal.includes('credit')) typeBg = C.YLW_LT;
      else if(typeVal.includes('equity')) typeBg = C.NAVY_LT;
      ws[`${L(1)}${rn}`].s.fill = Fill(typeBg);
    });

    ws['!ref'] = `A1:J${rows.length+2}`;
    ws['!cols'] = [W(28),W(24),W(13),W(14),W(7),W(18),W(30),W(18),W(13),W(13)];
    ws['!rows'] = [{hpt:24},{hpt:22}];
    ws['!merges'] = [MG(0,0,0,9)];
    ws['!freeze'] = {xSplit:2,ySplit:2};
    ws['!autofilter'] = {ref:'A2:J2'};
    XLSX.utils.book_append_sheet(wb, ws, RL);
  }

  // ══ Sheet 3: P & L ════════════════════════════════════════════════════════
  {
    const ws: CS = {};
    const npC = pl.netProfit>=0?C.GREEN:C.RED;
    const months = Object.entries(pl.monthlyBreakdown).sort();
    const DATA_START = 18; // monthly data starts here

    // Banner
    for(let c=0;c<6;c++) for(let r=1;r<=3;r++) ws[`${L(c)}${r}`]={v:'',t:'s',s:{fill:Fill(C.NAVY)}};
    wv(ws,1,0,`PROFIT & LOSS STATEMENT  —  ${co}`,'s',ss.titl(C.NAVY,16));
    wv(ws,2,0,`Source: Live SUMIF/SUMPRODUCT formulas from '${RL}' sheet`,'s',ss.sub(C.NAVY));
    wv(ws,3,0,`Generated: ${gd}`,'s',ss.sub(C.NAVY));

    // ── Summary table rows 5-14 ──
    wv(ws,5,0,'LINE ITEM','s',ss.hdrL(C.TEAL,C.WHITE,10));
    wv(ws,5,1,'Amount','s',ss.hdrC(C.TEAL,C.WHITE,10));
    wv(ws,5,2,'% Revenue','s',ss.hdrC(C.TEAL,C.WHITE,10));
    wv(ws,5,3,'Notes','s',ss.hdrC(C.NAVY2,C.WHITE,9));

    // Revenue
    bg(ws,6,0,3,C.NAVY2); wv(ws,6,0,'REVENUE','s',ss.secH(C.NAVY2));
    wv(ws,7,0,'  Total Revenue','s',ss.c(false,C.BLACK,'left',C.GRN_LT));
    wf(ws,7,1,SUMIF_chain(TINC,'I'),pl.totalRevenue,ss.m(true,C.GRN_HDR,C.GRN_LT));
    wf(ws,7,2,'IF(B7=0,0,B7/B7)',1,ss.p(false,C.BLACK,C.GRN_LT));
    wv(ws,7,3,`SUMIF Income/Revenue types from '${RL}'!B:I`,'s',ss.note());

    // COGS
    bg(ws,8,0,3,C.TEAL); wv(ws,8,0,'COST OF GOODS SOLD','s',ss.secH(C.TEAL));
    wv(ws,9,0,'  Total COGS','s',ss.c());
    wf(ws,9,1,SUMIF_chain(TCOGS,'I'),pl.totalCogs,ss.m());
    wf(ws,9,2,'IF(B7=0,0,B9/B7)',pl.totalRevenue?pl.totalCogs/pl.totalRevenue:0,ss.p());
    wv(ws,9,3,'SUMIF Cost of Goods Sold types','s',ss.note());

    // Gross Profit
    wv(ws,10,0,'GROSS PROFIT','s',ss.tL(C.GRN_HDR));
    wf(ws,10,1,'B7-B9',pl.grossProfit,ss.tM(C.GRN_HDR));
    wf(ws,10,2,'IF(B7=0,0,B10/B7)',pl.grossMargin,ss.tP(C.GRN_HDR));
    wv(ws,10,3,'= Revenue − COGS','s',{...ss.note(),fill:Fill(C.GRN_LT)});

    // OpEx
    bg(ws,11,0,3,C.TEAL); wv(ws,11,0,'OPERATING EXPENSES','s',ss.secH(C.TEAL));
    wv(ws,12,0,'  Total OpEx','s',ss.c());
    wf(ws,12,1,SUMIF_chain(TEXP,'I'),pl.totalExpenses,ss.m());
    wf(ws,12,2,'IF(B7=0,0,B12/B7)',pl.totalRevenue?pl.totalExpenses/pl.totalRevenue:0,ss.p());
    wv(ws,12,3,'SUMIF Expense types','s',ss.note());

    // Net Profit
    wv(ws,13,0,'NET PROFIT / (LOSS)','s',ss.tL(npC));
    wf(ws,13,1,'B10-B12',pl.netProfit,ss.tM(npC));
    wf(ws,13,2,'IF(B7=0,0,B13/B7)',pl.netMargin,ss.tP(npC));
    wv(ws,13,3,'= Gross Profit − OpEx','s',{...ss.note(),fill:Fill(pl.netProfit>=0?C.GRN_LT:C.RED_LT)});

    // ── Monthly breakdown rows 17+ ──
    bg(ws,16,0,5,C.NAVY); wv(ws,16,0,`MONTHLY BREAKDOWN  (${months.length} months · Live from '${RL}')`, 's', ss.hdrL(C.NAVY,C.WHITE,10));
    ['Month','Revenue','COGS','OpEx','Net Profit','Margin %'].forEach((h,i)=>wv(ws,17,i,h,'s',ss.hdrC(C.NAVY2,C.WHITE,10)));

    months.forEach(([mk,{revenue,cogs,expenses}],i)=>{
      const r=18+i; const rb=i%2===0?C.WHITE:C.OFF;
      const net=revenue-(cogs??0)-expenses;
      wv(ws,r,0,mk,'s',ss.c(false,C.BLACK,'left',rb));
      wf(ws,r,1,SUMPRODUCT_month(TINC,mk),revenue,ss.m(false,C.BLACK,rb));
      wf(ws,r,2,SUMPRODUCT_month(TCOGS,mk),cogs??0,ss.m(false,C.BLACK,rb));
      wf(ws,r,3,SUMPRODUCT_month(TEXP,mk),expenses,ss.m(false,C.BLACK,rb));
      wf(ws,r,4,`${L(1)}${r}-${L(2)}${r}-${L(3)}${r}`,net,ss.m(false,net>=0?C.GREEN:C.RED,rb));
      wf(ws,r,5,`IF(${L(1)}${r}=0,0,${L(4)}${r}/${L(1)}${r})`,revenue?net/revenue:0,ss.p(false,net>=0?C.GREEN:C.RED,rb));
    });

    const totR=18+months.length;
    wv(ws,totR,0,'TOTAL','s',ss.tL(C.NAVY));
    wf(ws,totR,1,`SUM(B18:B${totR-1})`,pl.totalRevenue,ss.tM(C.NAVY));
    wf(ws,totR,2,`SUM(C18:C${totR-1})`,pl.totalCogs,ss.tM(C.NAVY));
    wf(ws,totR,3,`SUM(D18:D${totR-1})`,pl.totalExpenses,ss.tM(C.NAVY));
    wf(ws,totR,4,`SUM(E18:E${totR-1})`,pl.netProfit,ss.tM(npC));
    wf(ws,totR,5,`IF(B${totR}=0,0,E${totR}/B${totR})`,pl.netMargin,ss.tP(npC));

    ws['!ref']=`A1:F${totR}`;
    ws['!cols']=[W(32),W(18),W(14),W(16),W(16),W(12)];
    ws['!rows']=[{hpt:36},{hpt:20},{hpt:16}];
    ws['!merges']=[
      MG(0,0,0,5),MG(1,0,1,5),MG(2,0,2,5), // title banner
      MG(5,0,5,3),   // summary header row
      MG(6,0,6,3),   // REVENUE section header
      MG(8,0,8,3),   // COGS section header
      MG(11,0,11,3), // EXPENSES section header
      MG(15,0,15,5), // monthly section header
    ];
    ws['!freeze']={xSplit:0,ySplit:5};
    XLSX.utils.book_append_sheet(wb, ws, 'P & L');
  }

  // ══ Sheet 4: BALANCE SHEET ════════════════════════════════════════════════
  {
    const ws: CS = {};
    let r=1;

    for(let c=0;c<4;c++) for(let rr=1;rr<=3;rr++) ws[`${L(c)}${rr}`]={v:'',t:'s',s:{fill:Fill(C.NAVY)}};
    wv(ws,1,0,`BALANCE SHEET  —  ${co}`,'s',ss.titl(C.NAVY,16));
    wv(ws,2,0,`LOOKUP formulas pull last balance per account from '${RL}' sheet`,'s',ss.sub(C.NAVY));
    wv(ws,3,0,`Generated: ${gd}`,'s',ss.sub(C.NAVY));

    r=5;
    wv(ws,r,0,'ACCOUNT','s',ss.hdrL(C.TEAL));
    wv(ws,r,1,'Balance ($)','s',ss.hdrC(C.TEAL));
    wv(ws,r,2,'% of Total','s',ss.hdrC(C.TEAL));
    wv(ws,r,3,'Type','s',ss.hdrC(C.NAVY2));
    wv(ws,r,4,'Source Formula','s',ss.hdrC(C.NAVY2,C.WHITE,9));
    r++;

    const writeBS = (title: string, entries: BalanceSheet['assets'], total: number, bg_c: string, typeLbl: string) => {
      bg(ws,r,0,4,bg_c); wv(ws,r,0,`  ${title}`,'s',ss.secH(bg_c)); r++;
      const sr=r;
      entries.forEach((e,i)=>{
        const rb=i%2===0?C.WHITE:C.OFF;
        wv(ws,r,0,`  ${e.account}`,'s',ss.c(false,C.BLACK,'left',rb));
        wf(ws,r,1,LOOKUP_bal(e.account),e.value,ss.m(false,C.BLACK,rb));
        wf(ws,r,2,`IF(SUM(B${sr}:B${sr+entries.length-1})=0,0,B${r}/SUM(B${sr}:B${sr+entries.length-1}))`,total?e.value/total:0,ss.p(false,C.BLACK,rb));
        wv(ws,r,3,typeLbl,'s',ss.pill(bg_c==='1F3864'?C.BLUE_LT:bg_c==='17375E'?C.YLW_LT:C.GRN_LT,C.BLACK));
        wv(ws,r,4,`=LOOKUP('${RL}'!J, acct="${e.account}")`,'s',{...ss.note(),numFmt:'@'});
        r++;
      });
      wv(ws,r,0,'Total','s',ss.tL(bg_c));
      wf(ws,r,1,`SUM(B${sr}:B${r-1})`,total,ss.tM(bg_c));
      wf(ws,r,2,'1',1,ss.tP(bg_c));
      r+=2;
    };

    writeBS('ASSETS',      bs.assets,      bs.totals.assetsTotal,      C.NAVY3,  'Asset');
    writeBS('LIABILITIES', bs.liabilities, bs.totals.liabilitiesTotal, C.TEAL,   'Liability');
    writeBS('EQUITY',      bs.equity,      bs.totals.equityTotal,      C.GRN_HDR,'Equity');

    // Reconciliation
    bg(ws,r,0,4,C.GOLD); wv(ws,r,0,'  BALANCE SHEET RECONCILIATION','s',ss.secH(C.GOLD)); r++;
    const aR=6+bs.assets.length;
    const lR=aR+3+bs.liabilities.length;
    const eR=lR+3+bs.equity.length;
    [
      ['Total Assets',        bs.totals.assetsTotal,                          `B${aR}`],
      ['Liabilities + Equity',bs.totals.liabilitiesTotal+bs.totals.equityTotal,`B${lR}+B${eR}`],
      ['Variance',            bs.variance,                                    `B${aR}-(B${lR}+B${eR})`],
    ].forEach(([lbl,val,fml],i)=>{
      const isVar=i===2; const fg=isVar?(bs.isBalanced?C.GREEN:C.RED):C.BLACK;
      wv(ws,r,0,String(lbl),'s',ss.c(true,fg));
      wf(ws,r,1,String(fml),Number(val),{...ss.m(true,fg),...(isVar&&!bs.isBalanced?{fill:Fill(C.RED_LT)}:isVar?{fill:Fill(C.GRN_LT)}:{})});
      r++;
    });
    wv(ws,r,0,'Status','s',ss.c(true));
    wv(ws,r,1,bs.isBalanced?'✔  BALANCED':'✘  NOT BALANCED','s',{
      font:F(true,12,bs.isBalanced?C.GREEN:C.RED), fill:Fill(bs.isBalanced?C.GRN_LT:C.RED_LT),
      alignment:Aln('center','center'), border:bdrBox(bs.isBalanced?C.GREEN:C.RED),
    });

    ws['!ref']=`A1:E${r+1}`;
    ws['!cols']=[W(34),W(18),W(12),W(12),W(36)];
    ws['!rows']=[{hpt:36},{hpt:20},{hpt:16}];
    ws['!merges']=[MG(0,0,0,4),MG(1,0,1,4),MG(2,0,2,4),MG(4,0,4,4)];
    ws['!freeze']={xSplit:0,ySplit:5};
    XLSX.utils.book_append_sheet(wb, ws, 'Balance Sheet');
  }

  // ══ Sheet 5: MONTH-OVER-MONTH ════════════════════════════════════════════
  if (mom && mom.months.length > 0) {
    const ws: CS = {};
    const months=mom.months; const nC=months.length+2;

    for(let c=0;c<nC;c++) for(let r=1;r<=3;r++) ws[`${L(c)}${r}`]={v:'',t:'s',s:{fill:Fill(C.NAVY)}};
    wv(ws,1,0,`MONTH-OVER-MONTH P&L  —  ${co}`,'s',ss.titl(C.NAVY,16));
    wv(ws,2,0,`SUMPRODUCT formulas from '${RL}' · ${months.length} months · ${monthLabel(months[0])} → ${monthLabel(months[months.length-1])}`,'s',ss.sub(C.NAVY));
    wv(ws,3,0,`Generated: ${gd}`,'s',ss.sub(C.NAVY));

    // Headers row 5
    wv(ws,5,0,'Account','s',ss.hdrL(C.NAVY,C.WHITE,10));
    months.forEach((m,i)=>wv(ws,5,i+1,monthLabel(m),'s',ss.hdrC(C.NAVY2,C.WHITE,9)));
    wv(ws,5,months.length+1,'Total','s',ss.hdrC(C.TEAL,C.WHITE,10));

    // MoM % row 6
    wv(ws,6,0,'MoM Δ%','s',ss.hdrL(C.TEAL,C.WHITE,9));
    months.forEach((_,i)=>{
      if(i===0){wv(ws,6,1,'Baseline','s',ss.hdrC(C.TEAL,C.WHITE,9));return;}
      const cur=mom.monthlyRevenue[months[i]]??0; const prv=mom.monthlyRevenue[months[i-1]]??0;
      const pct=prv!==0?(cur-prv)/Math.abs(prv):0;
      wf(ws,6,i+1,`IF(${L(i)}${8+mom.incomeCategories.length}=0,0,(${L(i+1)}${8+mom.incomeCategories.length}-${L(i)}${8+mom.incomeCategories.length})/ABS(${L(i)}${8+mom.incomeCategories.length}))`,pct,{...ss.p(false,cur>=prv?C.GREEN:C.RED,C.YLW_LT),numFmt:FMT_PCTD});
    });
    wv(ws,6,months.length+1,'','s',ss.hdrC(C.TEAL,C.WHITE,9));

    let r=7;
    // Income
    bg(ws,r,0,nC-1,C.GRN_HDR); wv(ws,r,0,'  ▸  INCOME','s',ss.secH(C.GRN_HDR)); r++;
    const incS=r;
    mom.incomeCategories.forEach((cat,ri)=>{
      const rb=ri%2===0?C.WHITE:C.OFF;
      wv(ws,r,0,`  ${cat.name}`,'s',ss.c(false,C.BLACK,'left',rb));
      months.forEach((m,i)=>{const v=cat.months[m]??0;wf(ws,r,i+1,SUMPRODUCT_acct(cat.name,m),v,ss.m(false,v!==0?C.BLACK:C.LGRAY,rb));});
      wf(ws,r,months.length+1,`SUM(${L(1)}${r}:${L(months.length)}${r})`,cat.total,ss.m(true,C.BLACK,rb));
      r++;
    });
    const revTR=r;
    wv(ws,r,0,'Total Revenue','s',ss.tL(C.GRN_HDR));
    months.forEach((_,i)=>wf(ws,r,i+1,`SUM(${L(i+1)}${incS}:${L(i+1)}${r-1})`,mom.monthlyRevenue[months[i]]??0,ss.tM(C.GRN_HDR)));
    wf(ws,r,months.length+1,`SUM(${L(1)}${r}:${L(months.length)}${r})`,mom.totalRevenue,ss.tM(C.GRN_HDR));
    r+=2;

    // Expenses
    bg(ws,r,0,nC-1,C.RED); wv(ws,r,0,'  ▸  EXPENSES','s',ss.secH(C.RED)); r++;
    const expS=r;
    mom.expenseCategories.forEach((cat,ri)=>{
      const rb=ri%2===0?C.WHITE:C.OFF;
      wv(ws,r,0,`  ${cat.name}`,'s',ss.c(false,C.BLACK,'left',rb));
      months.forEach((m,i)=>{const v=cat.months[m]??0;wf(ws,r,i+1,SUMPRODUCT_acct(cat.name,m),v,ss.m(false,v!==0?C.BLACK:C.LGRAY,rb));});
      wf(ws,r,months.length+1,`SUM(${L(1)}${r}:${L(months.length)}${r})`,cat.total,ss.m(true,C.BLACK,rb));
      r++;
    });
    const expTR=r;
    wv(ws,r,0,'Total Expenses','s',ss.tL(C.RED));
    months.forEach((_,i)=>wf(ws,r,i+1,`SUM(${L(i+1)}${expS}:${L(i+1)}${r-1})`,mom.monthlyExpenses[months[i]]??0,ss.tM(C.RED)));
    wf(ws,r,months.length+1,`SUM(${L(1)}${r}:${L(months.length)}${r})`,mom.totalExpenses,ss.tM(C.RED));
    r+=2;

    const npC2=mom.totalNetProfit>=0?C.GRN_HDR:C.RED;
    wv(ws,r,0,'NET PROFIT','s',ss.tL(npC2));
    months.forEach((_,i)=>{
      const v=mom.monthlyNetProfit[months[i]]??0;
      wf(ws,r,i+1,`${L(i+1)}${revTR}-${L(i+1)}${expTR}`,v,ss.tM(v>=0?C.GRN_HDR:C.RED));
    });
    wf(ws,r,months.length+1,`${L(months.length+1)}${revTR}-${L(months.length+1)}${expTR}`,mom.totalNetProfit,ss.tM(npC2));

    ws['!ref']=`A1:${L(nC-1)}${r}`;
    ws['!cols']=[W(30),...months.map(()=>W(12)),W(14)];
    ws['!rows']=[{hpt:36},{hpt:20},{hpt:16}];
    ws['!merges']=[MG(0,0,0,nC-1),MG(1,0,1,nC-1),MG(2,0,2,nC-1)];
    ws['!freeze']={xSplit:1,ySplit:5};
    XLSX.utils.book_append_sheet(wb, ws, 'Month-over-Month');
  }

  // ══ Sheet 6: CASH FLOW ════════════════════════════════════════════════════
  {
    const ws: CS = {};
    const ocfC=cf.operatingCashFlow>=0?C.GRN_HDR:C.RED;

    for(let c=0;c<2;c++) for(let r=1;r<=3;r++) ws[`${L(c)}${r}`]={v:'',t:'s',s:{fill:Fill(C.NAVY)}};
    wv(ws,1,0,`CASH FLOW STATEMENT  —  ${co}`,'s',ss.titl(C.NAVY,16));
    wv(ws,2,0,`Net Profit linked from 'P & L'!B13`,'s',ss.sub(C.NAVY));
    wv(ws,3,0,`Generated: ${gd}`,'s',ss.sub(C.NAVY));

    bg(ws,5,0,1,C.NAVY2); wv(ws,5,0,'OPERATING ACTIVITIES','s',ss.secH(C.NAVY2)); wv(ws,5,1,'Amount','s',ss.hdrC(C.NAVY2));
    wv(ws,6,0,'  Net Profit  (→ P & L tab B13)','s',ss.c(false,C.BLACK,'left',C.BLUE_LT));
    wf(ws,6,1,`'P & L'!B13`,pl.netProfit,{...ss.m(true,C.NAVY2,C.BLUE_LT),numFmt:FMT_MONEY});

    wv(ws,7,0,'  Working capital adjustments:','s',ss.c(false,C.DGRAY));

    let r=8;
    cf.adjustments.forEach((adj,i)=>{
      const rb=i%2===0?C.WHITE:C.OFF;
      wv(ws,r,0,`    ${adj.account}`,'s',ss.c(false,C.BLACK,'left',rb));
      wv(ws,r,1,adj.impact,'n',ss.m(false,adj.impact>=0?C.GREEN:C.RED,rb));
      r++;
    });

    wv(ws,r,0,'NET OPERATING CASH FLOW','s',ss.tL(ocfC));
    wf(ws,r,1,`B6+SUM(B8:B${r-1})`,cf.operatingCashFlow,ss.tM(ocfC));
    r+=2;
    wv(ws,r,0,`📌  Net Profit linked from 'P & L'!B13  ·  Full transactions in '${RL}' tab`,'s',ss.note());

    ws['!ref']=`A1:B${r}`;
    ws['!cols']=[W(40),W(20)];
    ws['!rows']=[{hpt:36},{hpt:20},{hpt:16}];
    ws['!merges']=[MG(0,0,0,1),MG(1,0,1,1),MG(2,0,2,1)];
    ws['!freeze']={xSplit:0,ySplit:4};
    XLSX.utils.book_append_sheet(wb, ws, 'Cash Flow');
  }

  // ══ Sheet 7: FLAGS ════════════════════════════════════════════════════════
  {
    const ws: CS = {};
    let r=1;

    for(let c=0;c<5;c++) for(let rr=1;rr<=3;rr++) ws[`${L(c)}${rr}`]={v:'',t:'s',s:{fill:Fill(C.RED)}};
    wv(ws,1,0,`FLAGS & AUDIT NOTES  —  ${co}`,'s',ss.titl(C.RED,16));
    wv(ws,2,0,`${analysis.inconsistentVendors.length} inconsistent vendors  ·  ${analysis.duplicates.length} duplicate transactions`,'s',ss.sub(C.RED));

    r=5;
    bg(ws,r,0,4,C.GOLD);
    ['Vendor','Reason','Accounts Affected','',''].forEach((h,i)=>wv(ws,r,i,h,'s',ss.hdrC(i===0?C.GOLD:C.NAVY2,C.WHITE,10)));
    r++;
    analysis.inconsistentVendors.forEach((v,i)=>{
      const rb=i%2===0?C.YLW_LT:C.WHITE;
      wv(ws,r,0,v.vendor,'s',ss.c(true,C.BLACK,'left',rb));
      wv(ws,r,1,v.reason,'s',ss.c(false,C.BLACK,'left',rb));
      wv(ws,r,2,v.accounts.join(', '),'s',ss.c(false,C.BLACK,'left',rb));
      r++;
    });
    if(!analysis.inconsistentVendors.length){wv(ws,r,0,'✔  None found','s',ss.c(true,C.GREEN));r++;}

    r++;
    bg(ws,r,0,4,C.RED);
    ['Vendor','Amount','Date','Count',''].forEach((h,i)=>wv(ws,r,i,h,'s',ss.hdrC(i===0?C.RED:C.NAVY2,C.WHITE,10)));
    r++;
    analysis.duplicates.forEach((d,i)=>{
      const rb=i%2===0?C.RED_LT:C.WHITE;
      wv(ws,r,0,d.name,'s',ss.c(true,C.BLACK,'left',rb));
      wv(ws,r,1,d.amount,'s',ss.c(false,C.BLACK,'right',rb));
      wv(ws,r,2,d.transactionDate,'s',ss.c(false,C.BLACK,'center',rb));
      wv(ws,r,3,d.occurrences,'n',ss.c(true,C.RED,'center',rb));
      r++;
    });
    if(!analysis.duplicates.length){wv(ws,r,0,'✔  None found','s',ss.c(true,C.GREEN));}

    ws['!ref']=`A1:E${r+1}`;
    ws['!cols']=[W(26),W(36),W(42),W(10),W(10)];
    ws['!merges']=[MG(0,0,0,4),MG(1,0,1,4),MG(2,0,2,4)];
    XLSX.utils.book_append_sheet(wb, ws, 'Flags');
  }

  XLSX.writeFile(wb, `${base(fileName)}_Financial_Analysis.xlsx`);
}

// ── PDF ───────────────────────────────────────────────────────────────────────
export function exportPdf(fileName: string, analysis: LedgerAnalysis, pl: ProfitAndLoss, bs: BalanceSheet, cf: CashFlowStatement): void {
  const tbl=(h:string[],rows:(string|number)[][])=>`<table><thead><tr>${h.map(x=>`<th>${x}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map(c=>`<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>${base(fileName)}</title><style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;font-size:12px;color:#111;padding:32px}h1{font-size:20px;margin-bottom:4px}.meta{color:#666;font-size:11px;margin-bottom:24px}.sec{margin-bottom:24px}h2{font-size:13px;font-weight:700;text-transform:uppercase;color:#1F3864;border-bottom:2px solid #1F3864;padding-bottom:4px;margin-bottom:8px}table{width:100%;border-collapse:collapse;font-size:11px}th{background:#1F3864;color:#fff;text-align:left;padding:5px 8px}td{padding:4px 8px;border-bottom:1px solid #e5e7eb}tr:nth-child(even)td{background:#f1f5f9}@media print{body{padding:16px}}</style></head><body><h1>${base(fileName)} — Financial Report</h1><p class="meta">Generated: ${new Date().toLocaleString()} · ${analysis.totalTransactions} transactions</p><div class="sec"><h2>P&L</h2>${tbl(['','Amount','% Rev'],[['Revenue',$f(pl.totalRevenue),'100%'],['COGS',$f(pl.totalCogs),pctStr(pl.totalRevenue?pl.totalCogs/pl.totalRevenue:0)],['Gross Profit',$f(pl.grossProfit),pctStr(pl.grossMargin)],['OpEx',$f(pl.totalExpenses),pctStr(pl.totalRevenue?pl.totalExpenses/pl.totalRevenue:0)],['Net Profit',$f(pl.netProfit),pctStr(pl.netMargin)]])}</div><div class="sec"><h2>Balance Sheet ${bs.isBalanced?'✓':'✗'}</h2>${tbl(['Cat','Account','Balance'],[...bs.assets.map(({account,value})=>['Asset',account,$f(value)]),...bs.liabilities.map(({account,value})=>['Liability',account,$f(value)]),...bs.equity.map(({account,value})=>['Equity',account,$f(value)])])}</div></body></html>`;
  const win=window.open('','_blank');if(!win)return;win.document.write(html);win.document.close();win.onload=()=>win.print();
}
