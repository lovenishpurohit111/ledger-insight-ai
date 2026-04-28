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
  return isNaN(n) ? 0 : (neg ? -n : n);
};

// ── Colours ──────────────────────────────────────────────────────────────────
const C = {
  NAVY:'1F3864', NAVY2:'2F5496', TEAL:'17375E',
  GREEN:'375623', GREEN2:'70AD47',
  RED:'C00000', AMBER:'ED7D31', GOLD:'BF8F00',
  WHITE:'FFFFFF', OFF:'F2F2F2', LGRAY:'D9D9D9', DGRAY:'595959', BLACK:'000000',
  BLUE_LT:'BDD7EE', GRN_LT:'E2EFDA', RED_LT:'FFDCDC', YLW_LT:'FFF2CC', NAVY_LT:'D6DCE4',
};

// ── Style helpers ─────────────────────────────────────────────────────────────
const thin  = (rgb: string) => ({ style: 'thin',   color: { rgb } });
const hair  = (rgb: string) => ({ style: 'hair',   color: { rgb } });
const med   = (rgb: string) => ({ style: 'medium', color: { rgb } });
const dbl   = (rgb: string) => ({ style: 'double', color: { rgb } });

const bdrAll  = (rgb = C.LGRAY): CS => ({ top: thin(rgb), bottom: thin(rgb), left: thin(rgb), right: thin(rgb) });
const bdrData = (): CS => ({ bottom: hair(C.LGRAY), right: hair(C.LGRAY) });
const bdrTotal= (): CS => ({ top: med(C.LGRAY), bottom: dbl(C.LGRAY) });

const font = (bold: boolean, sz: number, rgb: string, italic = false) => ({ bold, sz, name: 'Calibri', color: { rgb }, italic });
const fill = (rgb: string): CS => ({ fgColor: { rgb }, patternType: 'solid' });
const align = (h: string, v = 'center', wrap = false): CS => ({ horizontal: h, vertical: v, wrapText: wrap });
const numFmtMoney = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';
const numFmtPct   = '0.0%';
const numFmtPctDelta = '+0.0%;[Red]-0.0%;"-"';

const s = {
  hdr:  (bg: string, fg = C.WHITE, sz = 11): CS => ({ font: font(true, sz, fg),  fill: fill(bg), alignment: align('center'),       border: bdrAll(bg) }),
  hdrL: (bg: string, fg = C.WHITE, sz = 11): CS => ({ font: font(true, sz, fg),  fill: fill(bg), alignment: align('left', 'center'), border: bdrAll(bg) }),
  secH: (bg: string, fg = C.WHITE): CS =>          ({ font: font(true, 10, fg),   fill: fill(bg), alignment: align('left', 'center') }),
  cell: (bold = false, fg = C.BLACK, ha = 'left',  bg?: string): CS => ({ font: font(bold, 10, fg), ...(bg ? { fill: fill(bg) } : {}), alignment: align(ha, 'center'), border: bdrData() }),
  mon:  (bold = false, fg = C.BLACK, bg?: string):  CS => ({ font: font(bold, 10, fg), ...(bg ? { fill: fill(bg) } : {}), numFmt: numFmtMoney, alignment: align('right', 'center'), border: bdrData() }),
  pct:  (bold = false, fg = C.BLACK, bg?: string):  CS => ({ font: font(bold, 10, fg), ...(bg ? { fill: fill(bg) } : {}), numFmt: numFmtPct, alignment: align('right', 'center'), border: bdrData() }),
  totL: (bg: string, fg = C.WHITE): CS =>           ({ font: font(true, 11, fg),  fill: fill(bg), alignment: align('left', 'center'),  border: bdrTotal() }),
  tot:  (bg: string, fg = C.WHITE): CS =>           ({ font: font(true, 11, fg),  fill: fill(bg), numFmt: numFmtMoney, alignment: align('right', 'center'), border: bdrTotal() }),
  totP: (bg: string, fg = C.WHITE): CS =>           ({ font: font(true, 11, fg),  fill: fill(bg), numFmt: numFmtPct,   alignment: align('right', 'center'), border: bdrTotal() }),
  titl: (bg: string): CS =>                         ({ font: font(true, 18, C.WHITE), fill: fill(bg), alignment: align('left', 'center') }),
  sub:  (bg: string): CS =>                         ({ font: font(false, 9, C.NAVY_LT, true), fill: fill(bg), alignment: align('left', 'center') }),
  note: (): CS =>                                   ({ font: font(false, 9, C.DGRAY, true),   alignment: align('left', 'center') }),
};

const W = (n: number) => ({ wch: n });
const L = (n: number) => n < 26 ? String.fromCharCode(65 + n) : String.fromCharCode(64 + Math.floor(n / 26)) + String.fromCharCode(65 + (n % 26));
const MG = (r: number, c: number, r2: number, c2: number) => ({ s: { r, c }, e: { r: r2, c: c2 } });

// Cell writers
const wv = (ws: CS, addr: string, v: string | number, t: string, style: CS) => { ws[addr] = { v, t, s: style }; };
const wf = (ws: CS, addr: string, f: string, v: number, style: CS) => { ws[addr] = { t: 'n', f, v, s: style }; };
const wb_ = (ws: CS, r: number, c: number, v: string | number, t: string, style: CS) => wv(ws, `${L(c)}${r}`, v, t, style);
const wbf = (ws: CS, r: number, c: number, f: string, v: number, style: CS) => wf(ws, `${L(c)}${r}`, f, v, style);
const bg_ = (ws: CS, r: number, cols: number, rgb: string) => { for (let c = 0; c < cols; c++) { const a = `${L(c)}${r}`; if (!ws[a]) ws[a] = { v: '', t: 's' }; if (!ws[a].s) ws[a].s = {}; ws[a].s.fill = fill(rgb); } };

// ── Formula constants ─────────────────────────────────────────────────────────
const RL = 'Raw Ledger';
let DEND = 5000; // updated per export

// Type-matching SUMIF chain (one SUMIF per exact type string)
// Raw Ledger: col B = account type, col I = amount, col J = balance
const STYPES = {
  inc:  ['Income','income','Revenue','revenue','Sales','sales','Other Income','other income','Non-operating Income'],
  cogs: ['Cost of Goods Sold','cost of goods sold','COGS','cogs','Cost of Sales'],
  exp:  ['Expense','expense','Expenses','expenses','Other Expense','other expense','Non-operating Expense'],
  ast:  ['Asset','asset','Bank','bank','Accounts Receivable (A/R)','Other Current Assets','Fixed Assets','Other Assets','Inventory'],
  lia:  ['Liability','liability','Accounts Payable (A/P)','Credit Card','Other Current Liabilities','Long Term Liabilities','Other Liability'],
  eq:   ['Equity','equity','Retained Earnings','Opening Balance Equity'],
};

const sumCol = (types: string[], col: 'I' | 'J') =>
  types.map(t => `SUMIF('${RL}'!B$2:B$${DEND},"${t}",'${RL}'!${col}$2:${col}$${DEND})`).join('+');

const sumMonthTypes = (types: string[], mk: string) => {
  const parts = types.map(t => `('${RL}'!B$2:B$${DEND}="${t}")`).join(',');
  return `SUMPRODUCT((MMULT((EXACT(IF(ISERR('${RL}'!B$2:B$${DEND}),"",""),"")+0,IF({${types.map(()=>'1').join(',')}}=1,'${RL}'!B$2:B$${DEND},"Z")),'${RL}'!I$2:I$${DEND}))`;
};
// Simpler SUMPRODUCT for month+type — avoids MMULT complexity
const sumMT = (types: string[], mk: string) =>
  `SUMPRODUCT((${types.map(t => `('${RL}'!B$2:B$${DEND}="${t}")`).join('+')}>=1)*(TEXT('${RL}'!C$2:C$${DEND},"YYYY-MM")="${mk}")*('${RL}'!I$2:I$${DEND}))`;

const sumAcctMonth = (acct: string, mk: string) =>
  `SUMPRODUCT(('${RL}'!A$2:A$${DEND}="${acct.replace(/"/g, '""')}")*(TEXT('${RL}'!C$2:C$${DEND},"YYYY-MM")="${mk}")*('${RL}'!I$2:I$${DEND}))`;

const lastBal = (acct: string) =>
  `IFERROR(LOOKUP(2,1/('${RL}'!A$2:A$${DEND}="${acct.replace(/"/g, '""')}"),'${RL}'!J$2:J$${DEND}),0)`;

// ── CSV ───────────────────────────────────────────────────────────────────────
export function exportCsv(fileName: string, analysis: LedgerAnalysis, pl: ProfitAndLoss, bs: BalanceSheet, cf: CashFlowStatement): void {
  const e = (v: string | number) => { const s = String(v); return s.includes(',') ? `"${s}"` : s; };
  const sec = (t: string, h: string[], rows: (string | number)[][]) => [t, h.map(e).join(','), ...rows.map(r => r.map(e).join(',')), ''].join('\n');
  const out = [
    sec('P&L', ['Item','Amount'], [['Revenue',$f(pl.totalRevenue)],['COGS',$f(pl.totalCogs)],['Gross Profit',$f(pl.grossProfit)],['OpEx',$f(pl.totalExpenses)],['Net Profit',$f(pl.netProfit)]]),
    sec('BS',  ['Cat','Account','Balance'], [...bs.assets.map(({account,value})=>['Asset',account,$f(value)]),...bs.liabilities.map(({account,value})=>['Liability',account,$f(value)]),...bs.equity.map(({account,value})=>['Equity',account,$f(value)])]),
  ];
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([out.join('\n')], { type: 'text/csv' }));
  a.download = `${base(fileName)}.csv`; a.click();
}

// ── Excel ─────────────────────────────────────────────────────────────────────
export function exportExcel(
  fileName: string, analysis: LedgerAnalysis, pl: ProfitAndLoss,
  bs: BalanceSheet, cf: CashFlowStatement, mom?: MoMPL, rawRows?: LedgerRow[],
): void {
  const wb = XLSX.utils.book_new();
  const gd = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
  const co = base(fileName);
  const rows = rawRows ?? [];
  DEND = rows.length > 1 ? rows.length + 1 : 5000;

  const HDR10 = ['Distribution account','Distribution account type','Transaction date','Transaction type','Num','Name','Description','Split','Amount','Balance'];

  // ══ Raw Ledger ════════════════════════════════════════════════════════════
  {
    const ws: CS = {};
    // Write header
    HDR10.forEach((h, i) => wv(ws, `${L(i)}1`, h, 's', s.hdr(C.NAVY, C.WHITE, 10)));

    // Write data rows — CRITICAL: Amount and Balance must be numbers, not strings
    rows.forEach((row, ri) => {
      const rn = ri + 2;
      const bg = ri % 2 === 0 ? C.WHITE : C.OFF;
      HDR10.forEach((h, ci) => {
        const raw = row[h as keyof typeof row] ?? '';
        const isAmt = ci === 8 || ci === 9;
        if (isAmt) {
          const num = toNum(String(raw));
          ws[`${L(ci)}${rn}`] = { v: num, t: 'n', s: s.mon(false, C.BLACK, bg) };
        } else {
          ws[`${L(ci)}${rn}`] = { v: String(raw), t: 's', s: s.cell(false, C.BLACK, 'left', bg) };
        }
      });
    });

    ws['!ref'] = `A1:J${rows.length + 1}`;
    ws['!cols'] = [W(28),W(22),W(13),W(14),W(7),W(18),W(30),W(18),W(13),W(13)];
    ws['!rows'] = [{ hpt: 24 }];
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };
    ws['!autofilter'] = { ref: 'A1:J1' };
    XLSX.utils.book_append_sheet(wb, ws, RL);
  }

  // ══ P & L ═════════════════════════════════════════════════════════════════
  {
    const ws: CS = {};
    const npC = pl.netProfit >= 0 ? C.GREEN : C.RED;
    const months = Object.entries(pl.monthlyBreakdown).sort();
    const LR = 19 + months.length; // last row of monthly table

    // Title banner (rows 1-3, cols A-F)
    for (let c = 0; c < 6; c++) for (let r = 1; r <= 3; r++) { ws[`${L(c)}${r}`] = { v: '', t: 's', s: { fill: fill(C.NAVY) } }; }
    wb_(ws, 1, 0, `PROFIT & LOSS — ${co}`, 's', s.titl(C.NAVY));
    wb_(ws, 2, 0, `All values calculated live from '${RL}' sheet via SUMIF/SUMPRODUCT`, 's', s.sub(C.NAVY));
    wb_(ws, 3, 0, `Generated: ${gd}`, 's', s.sub(C.NAVY));

    // Summary table (rows 5-15)
    wb_(ws, 5, 0, 'LINE ITEM',    's', s.hdrL(C.TEAL));
    wb_(ws, 5, 1, 'Amount ($)',   's', s.hdr(C.TEAL));
    wb_(ws, 5, 2, '% of Revenue', 's', s.hdr(C.TEAL));
    wb_(ws, 5, 3, 'Notes',        's', s.hdr(C.TEAL));

    // Revenue row 6
    bg_(ws, 6, 4, C.NAVY2);
    wb_(ws, 6, 0, 'REVENUE', 's', s.secH(C.NAVY2));
    wb_(ws, 7, 0, '  Total Revenue', 's', s.cell(false, C.BLACK, 'left', C.GRN_LT));
    wbf(ws, 7, 1, sumCol(STYPES.inc, 'I'), pl.totalRevenue, s.mon(true, C.GREEN2, C.GRN_LT));
    wbf(ws, 7, 2, 'IF(B7=0,0,B7/B7)', 1, s.pct(false, C.BLACK, C.GRN_LT));
    wb_(ws, 7, 3, 'From Raw Ledger Income/Revenue accounts', 's', s.note());

    // COGS rows 8-10
    bg_(ws, 8, 4, C.TEAL);
    wb_(ws, 8, 0, 'COST OF GOODS SOLD', 's', s.secH(C.TEAL));
    wb_(ws, 9, 0, '  Total COGS', 's', s.cell());
    wbf(ws, 9, 1, sumCol(STYPES.cogs, 'I'), pl.totalCogs, s.mon());
    wbf(ws, 9, 2, 'IF(B7=0,0,B9/B7)', pl.totalRevenue ? pl.totalCogs / pl.totalRevenue : 0, s.pct());
    wb_(ws, 9, 3, 'From Raw Ledger COGS accounts', 's', s.note());

    wb_(ws, 10, 0, 'GROSS PROFIT', 's', s.totL(C.GREEN));
    wbf(ws, 10, 1, 'B7-B9', pl.grossProfit, s.tot(C.GREEN));
    wbf(ws, 10, 2, 'IF(B7=0,0,B10/B7)', pl.grossMargin, s.totP(C.GREEN));
    wb_(ws, 10, 3, '= Revenue − COGS', 's', { ...s.note(), fill: fill(C.GRN_LT) });

    // Expenses rows 11-13
    bg_(ws, 11, 4, C.TEAL);
    wb_(ws, 11, 0, 'OPERATING EXPENSES', 's', s.secH(C.TEAL));
    wb_(ws, 12, 0, '  Total OpEx', 's', s.cell());
    wbf(ws, 12, 1, sumCol(STYPES.exp, 'I'), pl.totalExpenses, s.mon());
    wbf(ws, 12, 2, 'IF(B7=0,0,B12/B7)', pl.totalRevenue ? pl.totalExpenses / pl.totalRevenue : 0, s.pct());
    wb_(ws, 12, 3, 'From Raw Ledger Expense accounts', 's', s.note());

    wb_(ws, 13, 0, 'NET PROFIT / (LOSS)', 's', s.totL(npC));
    wbf(ws, 13, 1, 'B10-B12', pl.netProfit, s.tot(npC));
    wbf(ws, 13, 2, 'IF(B7=0,0,B13/B7)', pl.netMargin, s.totP(npC));
    wb_(ws, 13, 3, '= Gross Profit − OpEx', 's', { ...s.note(), fill: fill(pl.netProfit >= 0 ? C.GRN_LT : C.RED_LT) });

    // Spacer row 14
    wb_(ws, 14, 0, '', 's', {});

    // Monthly breakdown header (row 15)
    bg_(ws, 15, 6, C.NAVY);
    wb_(ws, 15, 0, 'MONTHLY BREAKDOWN (Live formulas from Raw Ledger)', 's', s.hdrL(C.NAVY, C.WHITE, 10));
    ['Month','Revenue','COGS','OpEx','Net Profit','Margin %'].forEach((h, i) => wb_(ws, 16, i, h, 's', s.hdr(C.NAVY2, C.WHITE, 10)));

    months.forEach(([mk, { revenue, cogs, expenses }], i) => {
      const r = 17 + i;
      const bg = i % 2 === 0 ? C.WHITE : C.OFF;
      const net = revenue - (cogs ?? 0) - expenses;
      wb_(ws, r, 0, mk, 's', s.cell(false, C.BLACK, 'left', bg));
      wbf(ws, r, 1, sumMT(STYPES.inc, mk),  revenue,      s.mon(false, C.BLACK, bg));
      wbf(ws, r, 2, sumMT(STYPES.cogs, mk), cogs ?? 0,    s.mon(false, C.BLACK, bg));
      wbf(ws, r, 3, sumMT(STYPES.exp, mk),  expenses,     s.mon(false, C.BLACK, bg));
      wbf(ws, r, 4, `${L(1)}${r}-${L(2)}${r}-${L(3)}${r}`, net, s.mon(false, net >= 0 ? C.GREEN : C.RED, bg));
      wbf(ws, r, 5, `IF(${L(1)}${r}=0,0,${L(4)}${r}/${L(1)}${r})`, revenue ? net / revenue : 0, s.pct(false, net >= 0 ? C.GREEN : C.RED, bg));
    });

    const totR = 17 + months.length;
    wb_(ws, totR, 0, 'TOTAL', 's', s.totL(C.NAVY));
    wbf(ws, totR, 1, `SUM(B17:B${totR - 1})`, pl.totalRevenue, s.tot(C.NAVY));
    wbf(ws, totR, 2, `SUM(C17:C${totR - 1})`, pl.totalCogs, s.tot(C.NAVY));
    wbf(ws, totR, 3, `SUM(D17:D${totR - 1})`, pl.totalExpenses, s.tot(C.NAVY));
    wbf(ws, totR, 4, `SUM(E17:E${totR - 1})`, pl.netProfit, s.tot(npC));
    wbf(ws, totR, 5, `IF(B${totR}=0,0,E${totR}/B${totR})`, pl.netMargin, s.totP(npC));

    ws['!ref'] = `A1:F${totR}`;
    ws['!cols'] = [W(32), W(18), W(14), W(16), W(16), W(12)];
    ws['!rows'] = [{ hpt: 36 }, { hpt: 20 }, { hpt: 16 }, {}, { hpt: 22 }, { hpt: 22 }, { hpt: 22 }];
    ws['!merges'] = [
      MG(0,0,0,5), MG(1,0,1,5), MG(2,0,2,5),   // title banner
      MG(4,0,4,3),                                // summary header
      MG(5,0,5,0), MG(7,0,7,0), MG(8,0,8,0), MG(10,0,10,0), MG(11,0,11,0), // section headers
      MG(14,0,14,5),                              // monthly section header
    ];
    ws['!freeze'] = { xSplit: 0, ySplit: 5 };
    XLSX.utils.book_append_sheet(wb, ws, 'P & L');
  }

  // ══ Balance Sheet ══════════════════════════════════════════════════════════
  {
    const ws: CS = {};
    let r = 1;

    // Title
    for (let c = 0; c < 4; c++) for (let rr = 1; rr <= 3; rr++) { ws[`${L(c)}${rr}`] = { v: '', t: 's', s: { fill: fill(C.NAVY) } }; }
    wb_(ws, 1, 0, `BALANCE SHEET — ${co}`, 's', s.titl(C.NAVY));
    wb_(ws, 2, 0, `Last balance per account via LOOKUP from '${RL}' sheet`, 's', s.sub(C.NAVY));
    wb_(ws, 3, 0, `Generated: ${gd}`, 's', s.sub(C.NAVY));

    r = 5;
    wb_(ws, r, 0, 'ACCOUNT', 's', s.hdrL(C.TEAL)); wb_(ws, r, 1, 'Balance ($)', 's', s.hdr(C.TEAL)); wb_(ws, r, 2, '% of Total', 's', s.hdr(C.TEAL)); wb_(ws, r, 3, 'Source', 's', s.hdr(C.TEAL));
    r++;

    const writeSection = (title: string, entries: BalanceSheet['assets'], total: number, bg: string) => {
      bg_(ws, r, 4, bg); wb_(ws, r, 0, title, 's', s.secH(bg)); r++;
      const startR = r;
      entries.forEach((e, i) => {
        const rb = i % 2 === 0 ? C.WHITE : C.OFF;
        wb_(ws, r, 0, `  ${e.account}`, 's', s.cell(false, C.BLACK, 'left', rb));
        wbf(ws, r, 1, lastBal(e.account), e.value, s.mon(false, C.BLACK, rb));
        wbf(ws, r, 2, `IF(SUM(B${startR}:B${startR + entries.length})=0,0,B${r}/SUM(B${startR}:B${startR + entries.length - 1}))`, total ? e.value / total : 0, s.pct(false, C.BLACK, rb));
        wb_(ws, r, 3, `=LOOKUP from '${RL}' col J`, 's', s.note());
        r++;
      });
      wb_(ws, r, 0, 'Total', 's', s.totL(bg)); wbf(ws, r, 1, `SUM(B${startR}:B${r - 1})`, total, s.tot(bg)); wbf(ws, r, 2, '1', 1, s.totP(bg));
      r += 2;
      return r;
    };

    writeSection('ASSETS',      bs.assets,      bs.totals.assetsTotal,      C.NAVY2);
    writeSection('LIABILITIES', bs.liabilities, bs.totals.liabilitiesTotal, C.TEAL);
    writeSection('EQUITY',      bs.equity,      bs.totals.equityTotal,      C.GREEN);

    // Reconciliation
    bg_(ws, r, 4, C.GOLD); wb_(ws, r, 0, 'BALANCE SHEET CHECK', 's', s.secH(C.GOLD)); r++;
    wb_(ws, r, 0, 'Total Assets', 's', s.cell(true)); wbf(ws, r, 1, `B${6 + bs.assets.length}`, bs.totals.assetsTotal, s.mon(true)); r++;
    const liabTotR = 6 + bs.assets.length + bs.liabilities.length + 3;
    const eqTotR   = liabTotR + bs.equity.length + 2;
    wb_(ws, r, 0, 'Liabilities + Equity', 's', s.cell(true));
    wbf(ws, r, 1, `B${liabTotR}+B${eqTotR}`, bs.totals.liabilitiesTotal + bs.totals.equityTotal, s.mon(true)); r++;
    const varC = bs.isBalanced ? C.GREEN : C.RED;
    wb_(ws, r, 0, 'Variance (A − L − E)', 's', s.cell(true, varC));
    wbf(ws, r, 1, `B${r - 2}-B${r - 1}`, bs.variance, { ...s.mon(true, varC), ...(bs.isBalanced ? { fill: fill(C.GRN_LT) } : { fill: fill(C.RED_LT) }) }); r++;
    wb_(ws, r, 0, 'Status', 's', s.cell(true));
    wb_(ws, r, 1, bs.isBalanced ? 'BALANCED ✓' : 'NOT BALANCED ✗', 's', {
      font: font(true, 12, bs.isBalanced ? C.GREEN : C.RED),
      fill: fill(bs.isBalanced ? C.GRN_LT : C.RED_LT),
      alignment: align('center', 'center'),
    });

    ws['!ref'] = `A1:D${r}`;
    ws['!cols'] = [W(34), W(18), W(13), W(30)];
    ws['!rows'] = [{ hpt: 36 }, { hpt: 20 }, { hpt: 16 }];
    ws['!merges'] = [MG(0,0,0,3), MG(1,0,1,3), MG(2,0,2,3), MG(4,0,4,3)];
    ws['!freeze'] = { xSplit: 0, ySplit: 5 };
    XLSX.utils.book_append_sheet(wb, ws, 'Balance Sheet');
  }

  // ══ Month-over-Month ══════════════════════════════════════════════════════
  if (mom && mom.months.length > 0) {
    const ws: CS = {};
    const months = mom.months;
    const nC = months.length + 2; // Account + months + Total

    for (let c = 0; c < nC; c++) for (let r = 1; r <= 3; r++) { ws[`${L(c)}${r}`] = { v: '', t: 's', s: { fill: fill(C.NAVY) } }; }
    wb_(ws, 1, 0, `MONTH-OVER-MONTH P&L — ${co}`, 's', s.titl(C.NAVY));
    wb_(ws, 2, 0, `SUMPRODUCT formulas link live to '${RL}' · ${months.length} months`, 's', s.sub(C.NAVY));
    wb_(ws, 3, 0, `Generated: ${gd}`, 's', s.sub(C.NAVY));

    // Header row 5
    wb_(ws, 5, 0, 'Account', 's', s.hdrL(C.NAVY, C.WHITE, 10));
    months.forEach((m, i) => wb_(ws, 5, i + 1, monthLabel(m), 's', s.hdr(C.NAVY, C.WHITE, 10)));
    wb_(ws, 5, months.length + 1, 'Total', 's', s.hdr(C.NAVY, C.WHITE, 10));

    // MoM % row 6
    wb_(ws, 6, 0, 'MoM Rev %', 's', s.hdrL(C.TEAL, C.WHITE, 9));
    months.forEach((m, i) => {
      if (i === 0) { wb_(ws, 6, 1, '—', 's', s.hdr(C.TEAL, C.WHITE, 9)); return; }
      const cur = mom.monthlyRevenue[m] ?? 0;
      const prv = mom.monthlyRevenue[months[i - 1]] ?? 0;
      const pct = prv !== 0 ? (cur - prv) / Math.abs(prv) : 0;
      wbf(ws, 6, i + 1, `IF(${L(i)}${8 + mom.incomeCategories.length}=0,0,(${L(i+1)}${8 + mom.incomeCategories.length}-${L(i)}${8 + mom.incomeCategories.length})/${L(i)}${8 + mom.incomeCategories.length})`,
        pct, { ...s.pct(false, cur >= prv ? C.GREEN : C.RED, C.YLW_LT), numFmt: numFmtPctDelta });
    });
    wb_(ws, 6, months.length + 1, '', 's', s.hdr(C.TEAL, C.WHITE, 9));

    let r = 7;

    // Income section
    bg_(ws, r, nC, C.GREEN2); wb_(ws, r, 0, '▸  INCOME', 's', s.secH(C.GREEN2)); r++;
    const incStart = r;
    mom.incomeCategories.forEach((cat, ri) => {
      const rb = ri % 2 === 0 ? C.WHITE : C.OFF;
      wb_(ws, r, 0, `  ${cat.name}`, 's', s.cell(false, C.BLACK, 'left', rb));
      months.forEach((m, i) => { const v = cat.months[m] ?? 0; wbf(ws, r, i + 1, sumAcctMonth(cat.name, m), v, s.mon(false, v !== 0 ? C.BLACK : C.LGRAY, rb)); });
      wbf(ws, r, months.length + 1, `SUM(${L(1)}${r}:${L(months.length)}${r})`, cat.total, s.mon(true, C.BLACK, rb));
      r++;
    });
    const revTotR = r;
    wb_(ws, r, 0, 'Total Revenue', 's', s.totL(C.GREEN));
    months.forEach((_, i) => wbf(ws, r, i + 1, `SUM(${L(i+1)}${incStart}:${L(i+1)}${r - 1})`, mom.monthlyRevenue[months[i]] ?? 0, s.tot(C.GREEN)));
    wbf(ws, r, months.length + 1, `SUM(${L(1)}${r}:${L(months.length)}${r})`, mom.totalRevenue, s.tot(C.GREEN));
    r += 2;

    // Expense section
    bg_(ws, r, nC, C.RED); wb_(ws, r, 0, '▸  EXPENSES', 's', s.secH(C.RED)); r++;
    const expStart = r;
    mom.expenseCategories.forEach((cat, ri) => {
      const rb = ri % 2 === 0 ? C.WHITE : C.OFF;
      wb_(ws, r, 0, `  ${cat.name}`, 's', s.cell(false, C.BLACK, 'left', rb));
      months.forEach((m, i) => { const v = cat.months[m] ?? 0; wbf(ws, r, i + 1, sumAcctMonth(cat.name, m), v, s.mon(false, v !== 0 ? C.BLACK : C.LGRAY, rb)); });
      wbf(ws, r, months.length + 1, `SUM(${L(1)}${r}:${L(months.length)}${r})`, cat.total, s.mon(true, C.BLACK, rb));
      r++;
    });
    const expTotR = r;
    wb_(ws, r, 0, 'Total Expenses', 's', s.totL(C.RED));
    months.forEach((_, i) => wbf(ws, r, i + 1, `SUM(${L(i+1)}${expStart}:${L(i+1)}${r - 1})`, mom.monthlyExpenses[months[i]] ?? 0, s.tot(C.RED)));
    wbf(ws, r, months.length + 1, `SUM(${L(1)}${r}:${L(months.length)}${r})`, mom.totalExpenses, s.tot(C.RED));
    r += 2;

    // Net Profit
    const npC2 = mom.totalNetProfit >= 0 ? C.GREEN : C.RED;
    wb_(ws, r, 0, 'NET PROFIT', 's', s.totL(npC2));
    months.forEach((_, i) => {
      const v = mom.monthlyNetProfit[months[i]] ?? 0;
      wbf(ws, r, i + 1, `${L(i+1)}${revTotR}-${L(i+1)}${expTotR}`, v, s.tot(v >= 0 ? C.GREEN : C.RED));
    });
    wbf(ws, r, months.length + 1, `${L(months.length+1)}${revTotR}-${L(months.length+1)}${expTotR}`, mom.totalNetProfit, s.tot(npC2));

    ws['!ref'] = `A1:${L(nC - 1)}${r}`;
    ws['!cols'] = [W(30), ...months.map(() => W(13)), W(14)];
    ws['!rows'] = [{ hpt: 36 }, { hpt: 20 }, { hpt: 16 }];
    ws['!merges'] = [MG(0,0,0,nC-1), MG(1,0,1,nC-1), MG(2,0,2,nC-1)];
    ws['!freeze'] = { xSplit: 1, ySplit: 5 };
    XLSX.utils.book_append_sheet(wb, ws, 'Month-over-Month');
  }

  // ══ Cash Flow ══════════════════════════════════════════════════════════════
  {
    const ws: CS = {};
    const ocfC = cf.operatingCashFlow >= 0 ? C.GREEN : C.RED;

    for (let c = 0; c < 2; c++) for (let r = 1; r <= 3; r++) { ws[`${L(c)}${r}`] = { v: '', t: 's', s: { fill: fill(C.NAVY) } }; }
    wb_(ws, 1, 0, `CASH FLOW STATEMENT — ${co}`, 's', s.titl(C.NAVY));
    wb_(ws, 2, 0, `Net Profit referenced from 'P & L' sheet cell B13`, 's', s.sub(C.NAVY));
    wb_(ws, 3, 0, `Generated: ${gd}`, 's', s.sub(C.NAVY));

    bg_(ws, 5, 2, C.NAVY2); wb_(ws, 5, 0, 'OPERATING ACTIVITIES', 's', s.secH(C.NAVY2)); wb_(ws, 5, 1, 'Amount ($)', 's', s.hdr(C.NAVY2));
    wb_(ws, 6, 0, '  Net Profit  (from P&L)', 's', s.cell(false, C.BLACK, 'left', C.BLUE_LT));
    wbf(ws, 6, 1, `'P & L'!B13`, pl.netProfit, s.mon(true, C.NAVY2, C.BLUE_LT));
    wb_(ws, 7, 0, '  Working capital adjustments:', 's', s.cell(false, C.DGRAY));

    let r = 8;
    cf.adjustments.forEach((adj, i) => {
      const rb = i % 2 === 0 ? C.WHITE : C.OFF;
      wb_(ws, r, 0, `    ${adj.account}`, 's', s.cell(false, C.BLACK, 'left', rb));
      wbf(ws, r, 1, String(adj.impact), adj.impact, s.mon(false, adj.impact >= 0 ? C.GREEN : C.RED, rb));
      r++;
    });

    wb_(ws, r, 0, 'NET OPERATING CASH FLOW', 's', s.totL(ocfC));
    wbf(ws, r, 1, `B6+SUM(B8:B${r - 1})`, cf.operatingCashFlow, s.tot(ocfC));
    r++;
    wb_(ws, r + 1, 0, `📌 Net Profit linked from 'P & L'!B13  ·  Transactions in '${RL}' tab`, 's', s.note());

    ws['!ref'] = `A1:B${r + 2}`;
    ws['!cols'] = [W(40), W(20)];
    ws['!rows'] = [{ hpt: 36 }, { hpt: 20 }, { hpt: 16 }];
    ws['!merges'] = [MG(0,0,0,1), MG(1,0,1,1), MG(2,0,2,1)];
    ws['!freeze'] = { xSplit: 0, ySplit: 4 };
    XLSX.utils.book_append_sheet(wb, ws, 'Cash Flow');
  }

  // ══ Flags ══════════════════════════════════════════════════════════════════
  {
    const ws: CS = {};
    let r = 1;
    for (let c = 0; c < 4; c++) for (let rr = 1; rr <= 3; rr++) { ws[`${L(c)}${rr}`] = { v: '', t: 's', s: { fill: fill(C.RED) } }; }
    wb_(ws, 1, 0, `FLAGS & AUDIT — ${co}`, 's', s.titl(C.RED));
    wb_(ws, 2, 0, `${analysis.inconsistentVendors.length} inconsistent vendors · ${analysis.duplicates.length} duplicate transactions`, 's', s.sub(C.RED));

    r = 5;
    bg_(ws, r, 4, C.GOLD); ['Vendor','Reason','Accounts Affected',''].forEach((h, i) => wb_(ws, r, i, h, 's', s.hdr(i === 0 ? C.GOLD : C.NAVY2))); r++;
    analysis.inconsistentVendors.forEach((v, i) => {
      const rb = i % 2 === 0 ? C.YLW_LT : C.WHITE;
      wb_(ws, r, 0, v.vendor, 's', s.cell(true, C.BLACK, 'left', rb));
      wb_(ws, r, 1, v.reason, 's', s.cell(false, C.BLACK, 'left', rb));
      wb_(ws, r, 2, v.accounts.join(', '), 's', s.cell(false, C.BLACK, 'left', rb));
      r++;
    });
    if (!analysis.inconsistentVendors.length) { wb_(ws, r, 0, 'None found ✓', 's', s.cell(false, C.GREEN)); r++; }

    r++;
    bg_(ws, r, 4, C.RED); ['Vendor','Amount','Date','Count'].forEach((h, i) => wb_(ws, r, i, h, 's', s.hdr(i === 0 ? C.RED : C.NAVY2))); r++;
    analysis.duplicates.forEach((d, i) => {
      const rb = i % 2 === 0 ? C.RED_LT : C.WHITE;
      wb_(ws, r, 0, d.name, 's', s.cell(true, C.BLACK, 'left', rb));
      wb_(ws, r, 1, d.amount, 's', s.cell(false, C.BLACK, 'right', rb));
      wb_(ws, r, 2, d.transactionDate, 's', s.cell(false, C.BLACK, 'center', rb));
      wb_(ws, r, 3, d.occurrences, 'n', s.cell(true, C.RED, 'center', rb));
      r++;
    });
    if (!analysis.duplicates.length) { wb_(ws, r, 0, 'None found ✓', 's', s.cell(false, C.GREEN)); }

    ws['!ref'] = `A1:D${r + 1}`;
    ws['!cols'] = [W(26), W(36), W(42), W(10)];
    ws['!merges'] = [MG(0,0,0,3), MG(1,0,1,3)];
    XLSX.utils.book_append_sheet(wb, ws, 'Flags');
  }

  XLSX.writeFile(wb, `${base(fileName)}_Financial_Analysis.xlsx`);
}

// ── PDF ───────────────────────────────────────────────────────────────────────
export function exportPdf(fileName: string, analysis: LedgerAnalysis, pl: ProfitAndLoss, bs: BalanceSheet, cf: CashFlowStatement): void {
  const tbl = (h: string[], rows: (string|number)[][]) =>
    `<table><thead><tr>${h.map(x=>`<th>${x}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map(c=>`<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/><title>${base(fileName)}</title>
<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;font-size:12px;color:#111;padding:32px}
h1{font-size:20px;margin-bottom:4px}.meta{color:#666;font-size:11px;margin-bottom:24px}
.section{margin-bottom:24px}h2{font-size:13px;font-weight:700;text-transform:uppercase;color:#1F3864;border-bottom:2px solid #1F3864;padding-bottom:4px;margin-bottom:8px}
table{width:100%;border-collapse:collapse;font-size:11px}th{background:#1F3864;color:#fff;text-align:left;padding:5px 8px}
td{padding:4px 8px;border-bottom:1px solid #e5e7eb}tr:nth-child(even) td{background:#f1f5f9}@media print{body{padding:16px}}</style></head><body>
<h1>${base(fileName)} — Financial Report</h1><p class="meta">Generated: ${new Date().toLocaleString()} · ${analysis.totalTransactions} transactions</p>
<div class="section"><h2>P&L</h2>${tbl(['','Amount','% Rev'],[['Revenue',$f(pl.totalRevenue),'100%'],['COGS',$f(pl.totalCogs),`${(pl.totalRevenue?pl.totalCogs/pl.totalRevenue*100:0).toFixed(1)}%`],['Gross Profit',$f(pl.grossProfit),`${(pl.grossMargin*100).toFixed(1)}%`],['OpEx',$f(pl.totalExpenses),`${(pl.totalRevenue?pl.totalExpenses/pl.totalRevenue*100:0).toFixed(1)}%`],['Net Profit',$f(pl.netProfit),`${(pl.netMargin*100).toFixed(1)}%`]])}</div>
<div class="section"><h2>Balance Sheet ${bs.isBalanced?'✓ Balanced':'✗ Not Balanced'}</h2>${tbl(['Category','Account','Balance'],[...bs.assets.map(({account,value})=>['Asset',account,$f(value)]),...bs.liabilities.map(({account,value})=>['Liability',account,$f(value)]),...bs.equity.map(({account,value})=>['Equity',account,$f(value)])])}</div>
<div class="section"><h2>Cash Flow</h2>${tbl(['','Amount'],[['Net Profit',$f(cf.netProfit)],['Operating CF',$f(cf.operatingCashFlow)]])}</div>
</body></html>`;
  const win = window.open('', '_blank'); if (!win) return;
  win.document.write(html); win.document.close(); win.onload = () => win.print();
}

function base(f: string) { return f.replace(/\.[^.]+$/, '') || 'ledger'; }
