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

const $ = (n: number) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(n);

// ─── Colour palette (professional navy/green/gold) ───────────────────────────
const C = {
  NAVY:    '1F3864',
  NAVY2:   '2F5496',
  TEAL:    '17375E',
  GREEN:   '375623',
  GREEN2:  '70AD47',
  RED:     'C00000',
  RED2:    'FF0000',
  GOLD:    'C09000',
  AMBER:   'ED7D31',
  WHITE:   'FFFFFF',
  OFF_WHT: 'F2F2F2',
  LGRAY:   'D9D9D9',
  DGRAY:   'A6A6A6',
  BLACK:   '000000',
  BLUE_LT: 'BDD7EE',
  GRN_LT:  'E2EFDA',
  RED_LT:  'FFDCDC',
  YLW_LT:  'FFF2CC',
  NAVY_LT: 'D6DCE4',
};

// ─── Style helpers ────────────────────────────────────────────────────────────
const s = {
  hdr: (bg: string, fg = C.WHITE, sz = 11, bold = true): CS => ({
    font: { bold, sz, name: 'Calibri', color: { rgb: fg } },
    fill: { fgColor: { rgb: bg }, patternType: 'solid' },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: {
      top:    { style: 'thin', color: { rgb: C.LGRAY } },
      bottom: { style: 'thin', color: { rgb: C.LGRAY } },
      left:   { style: 'thin', color: { rgb: C.LGRAY } },
      right:  { style: 'thin', color: { rgb: C.LGRAY } },
    },
  }),
  cell: (bold = false, fg = C.BLACK, align: 'left'|'right'|'center' = 'left', bg?: string): CS => ({
    font: { bold, sz: 10, name: 'Calibri', color: { rgb: fg } },
    ...(bg ? { fill: { fgColor: { rgb: bg }, patternType: 'solid' } } : {}),
    alignment: { horizontal: align, vertical: 'center' },
    border: {
      bottom: { style: 'hair', color: { rgb: C.LGRAY } },
      right:  { style: 'hair', color: { rgb: C.LGRAY } },
    },
  }),
  money: (bold = false, fg = C.BLACK, bg?: string): CS => ({
    font: { bold, sz: 10, name: 'Calibri', color: { rgb: fg } },
    ...(bg ? { fill: { fgColor: { rgb: bg }, patternType: 'solid' } } : {}),
    numFmt: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
    alignment: { horizontal: 'right', vertical: 'center' },
    border: {
      bottom: { style: 'hair', color: { rgb: C.LGRAY } },
      right:  { style: 'hair', color: { rgb: C.LGRAY } },
    },
  }),
  pct: (bold = false, fg = C.BLACK, bg?: string): CS => ({
    font: { bold, sz: 10, name: 'Calibri', color: { rgb: fg } },
    ...(bg ? { fill: { fgColor: { rgb: bg }, patternType: 'solid' } } : {}),
    numFmt: '0.0%',
    alignment: { horizontal: 'right', vertical: 'center' },
    border: { bottom: { style: 'hair', color: { rgb: C.LGRAY } }, right: { style: 'hair', color: { rgb: C.LGRAY } } },
  }),
  title: (bg: string): CS => ({
    font: { bold: true, sz: 18, name: 'Calibri', color: { rgb: C.WHITE } },
    fill: { fgColor: { rgb: bg }, patternType: 'solid' },
    alignment: { horizontal: 'left', vertical: 'center' },
  }),
  subtitle: (bg: string): CS => ({
    font: { sz: 10, name: 'Calibri', color: { rgb: C.NAVY_LT }, italic: true },
    fill: { fgColor: { rgb: bg }, patternType: 'solid' },
    alignment: { horizontal: 'left', vertical: 'center' },
  }),
  kpiVal: (fg = C.NAVY): CS => ({
    font: { bold: true, sz: 22, name: 'Calibri', color: { rgb: fg } },
    alignment: { horizontal: 'center', vertical: 'center' },
  }),
  kpiLbl: (): CS => ({
    font: { sz: 9, name: 'Calibri', color: { rgb: C.DGRAY }, italic: true },
    alignment: { horizontal: 'center', vertical: 'center' },
    numFmt: '@',
  }),
  sectionHdr: (bg: string): CS => ({
    font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: C.WHITE } },
    fill: { fgColor: { rgb: bg }, patternType: 'solid' },
    alignment: { horizontal: 'left', vertical: 'center' },
    border: { bottom: { style: 'medium', color: { rgb: bg } } },
  }),
  total: (bg: string, fg = C.WHITE): CS => ({
    font: { bold: true, sz: 11, name: 'Calibri', color: { rgb: fg } },
    fill: { fgColor: { rgb: bg }, patternType: 'solid' },
    numFmt: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
    alignment: { horizontal: 'right', vertical: 'center' },
    border: { top: { style: 'thin', color: { rgb: fg } }, bottom: { style: 'double', color: { rgb: fg } } },
  }),
  totalLabel: (bg: string, fg = C.WHITE): CS => ({
    font: { bold: true, sz: 11, name: 'Calibri', color: { rgb: fg } },
    fill: { fgColor: { rgb: bg }, patternType: 'solid' },
    alignment: { horizontal: 'left', vertical: 'center' },
    border: { top: { style: 'thin', color: { rgb: fg } }, bottom: { style: 'double', color: { rgb: fg } } },
  }),
};

// ─── Utilities ────────────────────────────────────────────────────────────────
const W = (n: number) => ({ wch: n });
const cell = (ws: CS, addr: string, v: string|number, t: string, style: CS) => {
  ws[addr] = { v, t, s: style };
};
const setS = (ws: CS, addr: string, style: CS) => {
  if (!ws[addr]) ws[addr] = { t: 'z', v: '' };
  ws[addr].s = style;
};
const mergeRange = (r: number, c: number, r2: number, c2: number) => ({ s: { r, c }, e: { r: r2, c: c2 } });
const colLetter = (n: number) => n < 26 ? String.fromCharCode(65 + n) : String.fromCharCode(64 + Math.floor(n/26)) + String.fromCharCode(65 + (n % 26));

function baseRef(sheetName: string) {
  // e.g. COUNTA('Raw Ledger'!A:A)-1 for row count
  return `'${sheetName}'`;
}

// ─── CSV export ───────────────────────────────────────────────────────────────
export function exportCsv(fileName: string, analysis: LedgerAnalysis, pl: ProfitAndLoss, bs: BalanceSheet, cf: CashFlowStatement): void {
  const esc = (v: string|number) => { const s = String(v); return s.includes(',') || s.includes('"') ? `"${s.replace(/"/g,'""')}"` : s; };
  const sec = (title: string, hdrs: string[], rows: (string|number)[][]) =>
    [title, hdrs.map(esc).join(','), ...rows.map(r => r.map(esc).join(',')), ''].join('\n');

  const sections = [
    sec('LEDGER ANALYSIS',['Metric','Value'],[['Total Transactions',analysis.totalTransactions],['Inconsistent Vendors',analysis.inconsistentVendors.length],['Duplicates',analysis.duplicates.length]]),
    sec('P&L',['Item','Amount'],[['Revenue',$(pl.totalRevenue)],['COGS',$(pl.totalCogs)],['Gross Profit',$(pl.grossProfit)],['OpEx',$(pl.totalExpenses)],['Net Profit',$(pl.netProfit)]]),
    sec('BALANCE SHEET',['Category','Account','Balance'],[
      ...bs.assets.map(({account,value})=>['Asset',account,$(value)]),
      ...bs.liabilities.map(({account,value})=>['Liability',account,$(value)]),
      ...bs.equity.map(({account,value})=>['Equity',account,$(value)]),
    ]),
    sec('CASH FLOW',['Item','Amount'],[['Net Profit',$(cf.netProfit)],['Operating CF',$(cf.operatingCashFlow)]]),
  ];
  const blob = new Blob([sections.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = `${base(fileName)}_analysis.csv`; a.click();
  URL.revokeObjectURL(url);
}

// ─── Excel export ─────────────────────────────────────────────────────────────
export function exportExcel(
  fileName: string,
  analysis: LedgerAnalysis,
  pl: ProfitAndLoss,
  bs: BalanceSheet,
  cf: CashFlowStatement,
  mom?: MoMPL,
  rawRows?: LedgerRow[],
): void {
  const wb = XLSX.utils.book_new();
  const genDate = new Date().toLocaleDateString('en-US', { year:'numeric', month:'long', day:'numeric' });
  const companyName = base(fileName);
  const RAW_SHEET = 'Raw Ledger';

  // ── Sheet 1: Cover ──────────────────────────────────────────────────────────
  {
    const ws: CS = { '!ref': 'A1:G20' };
    const M = (r:number, c:number, r2:number, c2:number) => mergeRange(r,c,r2,c2);

    // Background banner
    for (let r = 0; r < 8; r++) for (let c = 0; c < 7; c++) {
      const addr = `${colLetter(c)}${r+1}`;
      ws[addr] = { v: '', t: 's', s: { fill: { fgColor: { rgb: C.NAVY }, patternType: 'solid' } } };
    }

    ws['A1'] = { v: '📊 FINANCIAL ANALYSIS REPORT', t: 's', s: s.title(C.NAVY) };
    ws['A2'] = { v: companyName, t: 's', s: { font: { bold: true, sz: 14, name: 'Calibri', color: { rgb: C.BLUE_LT } }, fill: { fgColor: { rgb: C.NAVY }, patternType: 'solid' }, alignment: { horizontal: 'left', vertical: 'center' } } };
    ws['A3'] = { v: `Generated: ${genDate}`, t: 's', s: s.subtitle(C.NAVY) };
    ws['A4'] = { v: `Source: ${fileName}`, t: 's', s: s.subtitle(C.NAVY) };
    ws['A5'] = { v: `Transactions: ${analysis.totalTransactions.toLocaleString()}`, t: 's', s: s.subtitle(C.NAVY) };

    // KPI row (rows 10–13)
    const kpis = [
      { label: 'Total Revenue', val: pl.totalRevenue, fmt: true, color: C.GREEN2 },
      { label: 'Total Expenses', val: pl.totalAllExpenses, fmt: true, color: C.AMBER },
      { label: 'Net Profit', val: pl.netProfit, fmt: true, color: pl.netProfit >= 0 ? C.GREEN2 : C.RED },
      { label: 'Gross Margin', val: pl.grossMargin, fmt: false, color: C.NAVY2 },
      { label: 'Net Margin', val: pl.netMargin, fmt: false, color: C.NAVY2 },
      { label: 'BS Balanced', val: bs.isBalanced ? 'YES ✓' : 'NO ✗', fmt: false, color: bs.isBalanced ? C.GREEN : C.RED },
    ];

    kpis.forEach(({ label, val, fmt: isMoney, color }, i) => {
      const col = colLetter(i);
      ws[`${col}10`] = { v: label, t: 's', s: { ...s.hdr(C.NAVY2, C.WHITE, 9), border: undefined } };
      const dispVal = typeof val === 'string' ? val : isMoney ? val : `${(val * 100).toFixed(1)}%`;
      const numVal = typeof val === 'number' ? val : 0;
      ws[`${col}11`] = {
        v: typeof val === 'string' ? val : numVal,
        t: typeof val === 'string' ? 's' : 'n',
        s: {
          font: { bold: true, sz: 18, name: 'Calibri', color: { rgb: color } },
          alignment: { horizontal: 'center', vertical: 'center' },
          numFmt: isMoney ? '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)' : typeof val === 'number' ? '0.0%' : '@',
          fill: { fgColor: { rgb: C.OFF_WHT }, patternType: 'solid' },
          border: { bottom: { style: 'medium', color: { rgb: color } } },
        },
      };
    });

    // Sheet index
    const sheets = ['Raw Ledger','P & L','Balance Sheet','Cash Flow','Month-over-Month','Flags'];
    ws['A14'] = { v: 'CONTENTS', t: 's', s: s.hdr(C.TEAL, C.WHITE, 10) };
    sheets.forEach((sh, i) => {
      ws[`A${15+i}`] = { v: `  ${i+1}.  ${sh}`, t: 's', s: s.cell(false, C.NAVY2) };
    });

    ws['!cols'] = Array(7).fill(W(18));
    ws['!rows'] = [{ hpt: 36 }, { hpt: 28 }, { hpt: 20 }, { hpt: 20 }, { hpt: 20 }, {}, {}, {}, { hpt: 8 }, { hpt: 22 }, { hpt: 40 }, {}, {}, { hpt: 22 }];
    ws['!merges'] = [
      M(0,0,0,6), M(1,0,1,6), M(2,0,2,6), M(3,0,3,6), M(4,0,4,6),
      M(9,0,9,6), M(10,0,10,0), M(10,1,10,1), M(10,2,10,2), M(10,3,10,3), M(10,4,10,4), M(10,5,10,5),
      M(13,0,13,6),
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Cover');
  }

  // ── Sheet 2: Raw Ledger ─────────────────────────────────────────────────────
  if (rawRows && rawRows.length > 0) {
    const headers = ['Distribution account','Distribution account type','Transaction date','Transaction type','Num','Name','Description','Split','Amount','Balance'];
    const aoa = [headers, ...rawRows.map(r => headers.map(h => r[h as keyof typeof r] ?? ''))];
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Style header
    headers.forEach((_, i) => {
      const addr = `${colLetter(i)}1`;
      ws[addr].s = s.hdr(C.NAVY, C.WHITE, 10);
    });

    // Alternating rows
    rawRows.forEach((_, ri) => {
      const rowNum = ri + 2;
      const bg = ri % 2 === 0 ? C.WHITE : C.OFF_WHT;
      headers.forEach((_, ci) => {
        const addr = `${colLetter(ci)}${rowNum}`;
        if (!ws[addr]) ws[addr] = { t: 's', v: '' };
        const isAmt = ci === 8 || ci === 9;
        ws[addr].s = isAmt ? s.money(false, C.BLACK, bg) : s.cell(false, C.BLACK, 'left', bg);
      });
    });

    ws['!cols'] = [W(28),W(24),W(14),W(14),W(8),W(18),W(32),W(18),W(14),W(14)];
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };
    ws['!autofilter'] = { ref: `A1:J1` };
    XLSX.utils.book_append_sheet(wb, ws, RAW_SHEET);
  }

  // ── Sheet 3: P & L ─────────────────────────────────────────────────────────
  {
    const ws: CS = {};
    let r = 0;

    const put = (row: number, col: number, v: string|number, t: string, style: CS) => {
      ws[`${colLetter(col)}${row}`] = { v, t, s: style };
    };

    // Title
    put(1,0,`PROFIT & LOSS STATEMENT — ${companyName}`,'s',s.title(C.NAVY)); r=1;
    put(2,0,`For the period ended ${genDate}`,'s',s.subtitle(C.NAVY)); r=2;
    put(3,0,`Source: ${fileName} (see '${RAW_SHEET}' tab for full data)`,'s',s.subtitle(C.NAVY));

    // Section: Revenue
    r=5;
    put(r,0,'REVENUE','s',s.sectionHdr(C.NAVY2));
    put(r,1,'Amount','s',s.hdr(C.NAVY2,C.WHITE,10));
    put(r,2,'% of Revenue','s',s.hdr(C.NAVY2,C.WHITE,10));

    r++;
    put(r,0,'  Total Revenue','s',s.cell(false,C.BLACK,'left',C.GRN_LT));
    put(r,1,pl.totalRevenue,'n',s.money(false,C.BLACK,C.GRN_LT));
    put(r,2,1,'n',s.pct(false,C.BLACK,C.GRN_LT));

    // COGS
    r+=2;
    put(r,0,'COST OF GOODS SOLD','s',s.sectionHdr(C.TEAL));
    put(r,1,'','s',s.hdr(C.TEAL));
    put(r,2,'','s',s.hdr(C.TEAL));

    r++;
    put(r,0,'  Total COGS','s',s.cell());
    put(r,1,pl.totalCogs,'n',s.money());
    put(r,2,pl.totalRevenue ? pl.totalCogs/pl.totalRevenue : 0,'n',s.pct());

    r++;
    put(r,0,'GROSS PROFIT','s',s.totalLabel(C.GREEN,C.WHITE));
    put(r,1,pl.grossProfit,'n',s.total(C.GREEN,C.WHITE));
    put(r,2,pl.grossMargin,'n',{...s.total(C.GREEN,C.WHITE),numFmt:'0.0%'});

    // Operating expenses
    r+=2;
    put(r,0,'OPERATING EXPENSES','s',s.sectionHdr(C.TEAL));
    put(r,1,'','s',s.hdr(C.TEAL));
    put(r,2,'','s',s.hdr(C.TEAL));

    r++;
    put(r,0,'  Total Operating Expenses','s',s.cell());
    put(r,1,pl.totalExpenses,'n',s.money());
    put(r,2,pl.totalRevenue ? pl.totalExpenses/pl.totalRevenue : 0,'n',s.pct());

    r++;
    put(r,0,'NET PROFIT / (LOSS)','s',s.totalLabel(pl.netProfit>=0?C.GREEN:C.RED));
    put(r,1,pl.netProfit,'n',s.total(pl.netProfit>=0?C.GREEN:C.RED));
    put(r,2,pl.netMargin,'n',{...s.total(pl.netProfit>=0?C.GREEN:C.RED),numFmt:'0.0%'});

    // Monthly table
    r+=3;
    put(r,0,'MONTHLY BREAKDOWN','s',s.sectionHdr(C.NAVY));
    put(r,1,'','s',s.hdr(C.NAVY));
    put(r,2,'','s',s.hdr(C.NAVY));
    put(r,3,'','s',s.hdr(C.NAVY));
    put(r,4,'','s',s.hdr(C.NAVY));

    r++;
    ['Month','Revenue','COGS','Op Expenses','Net Profit','Margin %'].forEach((h,i) => {
      put(r,i,h,'s',s.hdr(C.NAVY2,C.WHITE,10));
    });

    const months = Object.entries(pl.monthlyBreakdown).sort();
    months.forEach(([m, { revenue, cogs, expenses }], i) => {
      r++;
      const net = revenue - (cogs??0) - expenses;
      const bg = i % 2 === 0 ? C.WHITE : C.OFF_WHT;
      put(r,0,m,'s',s.cell(false,C.BLACK,'left',bg));
      put(r,1,revenue,'n',s.money(false,C.BLACK,bg));
      put(r,2,cogs??0,'n',s.money(false,C.BLACK,bg));
      put(r,3,expenses,'n',s.money(false,C.BLACK,bg));
      put(r,4,net,'n',s.money(false,net>=0?C.GREEN:C.RED,bg));
      put(r,5,revenue ? net/revenue : 0,'n',s.pct(false,net>=0?C.GREEN:C.RED,bg));
    });

    // Totals row
    r++;
    const totNet = pl.grossProfit - pl.totalExpenses;
    put(r,0,'TOTAL','s',s.totalLabel(C.NAVY));
    put(r,1,pl.totalRevenue,'n',s.total(C.NAVY));
    put(r,2,pl.totalCogs,'n',s.total(C.NAVY));
    put(r,3,pl.totalExpenses,'n',s.total(C.NAVY));
    put(r,4,pl.netProfit,'n',s.total(pl.netProfit>=0?C.GREEN:C.RED));
    put(r,5,pl.netMargin,'n',{...s.total(C.NAVY),numFmt:'0.0%'});

    ws['!ref'] = `A1:F${r}`;
    ws['!cols'] = [W(34),W(18),W(16),W(18),W(18),W(12)];
    ws['!rows'] = [{ hpt: 36 },{hpt:20},{hpt:16}];
    ws['!merges'] = [mergeRange(0,0,0,5),mergeRange(1,0,1,5),mergeRange(2,0,2,5)];
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };
    XLSX.utils.book_append_sheet(wb, ws, 'P & L');
  }

  // ── Sheet 4: Balance Sheet ──────────────────────────────────────────────────
  {
    const ws: CS = {};
    let r = 0;

    const put = (row: number, col: number, v: string|number, t: string, style: CS) => {
      ws[`${colLetter(col)}${row}`] = { v, t, s: style };
    };

    put(1,0,`BALANCE SHEET — ${companyName}`,'s',s.title(C.NAVY));
    put(2,0,genDate,'s',s.subtitle(C.NAVY));
    put(3,0,`Variance (A − L − E): ${$(bs.variance)} — ${bs.isBalanced ? 'BALANCED ✓' : 'NOT BALANCED ✗'}`,'s',{
      ...s.subtitle(C.NAVY),
      font: { sz:10, name:'Calibri', color:{ rgb: bs.isBalanced?C.GREEN2:C.RED }, bold:true },
    });

    const writeSection = (title: string, entries: BalanceSheet['assets'], totalLabel: string, total: number, bg: string) => {
      r += 2;
      put(r,0,title,'s',s.sectionHdr(bg));
      put(r,1,'Balance','s',s.hdr(bg));
      put(r,2,'% of Total','s',s.hdr(bg));

      entries.forEach((e, i) => {
        r++;
        const rowBg = i%2===0 ? C.WHITE : C.OFF_WHT;
        put(r,0,`  ${e.account}${e.isCurrent !== undefined ? (e.isCurrent ? ' ◦' : '') : ''}`,'s',s.cell(false,C.BLACK,'left',rowBg));
        put(r,1,e.value,'n',s.money(false,C.BLACK,rowBg));
        put(r,2,total ? e.value/total : 0,'n',s.pct(false,C.BLACK,rowBg));
      });

      r++;
      put(r,0,`  ${totalLabel}`,'s',s.totalLabel(bg));
      put(r,1,total,'n',s.total(bg));
      put(r,2,1,'n',{...s.total(bg),numFmt:'0.0%'});
    };

    writeSection('ASSETS (◦ = Current)', bs.assets, 'Total Assets', bs.totals.assetsTotal, C.NAVY2);
    writeSection('LIABILITIES (◦ = Current)', bs.liabilities, 'Total Liabilities', bs.totals.liabilitiesTotal, C.TEAL);
    writeSection('EQUITY', bs.equity, 'Total Equity', bs.totals.equityTotal, C.GREEN);

    // Reconciliation
    r += 2;
    put(r,0,'BALANCE SHEET CHECK','s',s.sectionHdr(bs.isBalanced?C.GREEN:C.RED));
    put(r,1,'','s',{fill:{fgColor:{rgb:bs.isBalanced?C.GREEN:C.RED},patternType:'solid'}});
    put(r,2,'','s',{fill:{fgColor:{rgb:bs.isBalanced?C.GREEN:C.RED},patternType:'solid'}});

    const checks = [
      ['Total Assets', bs.totals.assetsTotal],
      ['Total Liabilities + Equity', bs.totals.liabilitiesTotal + bs.totals.equityTotal],
      ['Variance (A − L − E)', bs.variance],
    ];
    checks.forEach(([lbl, val]) => {
      r++;
      const isVar = String(lbl).includes('Variance');
      const vNum = Number(val);
      const fg = isVar ? (Math.abs(vNum) <= 1 ? C.GREEN : C.RED) : C.BLACK;
      put(r,0,String(lbl),'s',s.cell(true,fg));
      put(r,1,vNum,'n',s.money(true,fg,isVar?(Math.abs(vNum)<=1?C.GRN_LT:C.RED_LT):undefined));
    });

    r++;
    put(r,0,'Status','s',s.cell(true));
    put(r,1,bs.isBalanced?'BALANCED ✓':'NOT BALANCED ✗','s',{
      ...s.cell(true,bs.isBalanced?C.GREEN:C.RED,'left',bs.isBalanced?C.GRN_LT:C.RED_LT),
      font:{bold:true,sz:12,name:'Calibri',color:{rgb:bs.isBalanced?C.GREEN:C.RED}},
    });

    // Financial ratios
    r += 2;
    put(r,0,'FINANCIAL RATIOS','s',s.sectionHdr(C.NAVY2));
    put(r,1,'Value','s',s.hdr(C.NAVY2));
    put(r,2,'Benchmark','s',s.hdr(C.NAVY2));
    put(r,3,'Formula','s',s.hdr(C.NAVY2));

    const ratios = [
      { label:'Current Ratio', val:bs.ratios.currentRatio, benchmark:'≥ 1.5', formula:'Current Assets ÷ Current Liabilities', good: (v:number) => v>=1.5 },
      { label:'Quick Ratio', val:bs.ratios.quickRatio, benchmark:'≥ 1.0', formula:'(Current Assets − Inventory) ÷ Current Liabilities', good: (v:number) => v>=1.0 },
      { label:'Debt-to-Equity', val:bs.ratios.debtToEquity, benchmark:'≤ 2.0', formula:'Total Liabilities ÷ Total Equity', good: (v:number) => v<=2.0 },
      { label:'Debt Ratio', val:bs.ratios.debtRatio, benchmark:'≤ 0.50', formula:'Total Liabilities ÷ Total Assets', good: (v:number) => v<=0.5 },
    ];

    ratios.forEach((ratio, i) => {
      r++;
      const bg = i%2===0 ? C.WHITE : C.OFF_WHT;
      const isGood = ratio.val !== null && ratio.good(ratio.val);
      const fg = ratio.val === null ? C.BLACK : isGood ? C.GREEN : C.RED;
      put(r,0,`  ${ratio.label}`,'s',s.cell(false,C.BLACK,'left',bg));
      put(r,1,ratio.val ?? 'N/A', ratio.val !== null ? 'n' : 's', {
        ...s.cell(true,fg,'right',bg), numFmt:'0.00',
      });
      put(r,2,ratio.benchmark,'s',s.cell(false,C.DGRAY,'center',bg));
      put(r,3,ratio.formula,'s',s.cell(false,C.DGRAY,'left',bg));
    });

    ws['!ref'] = `A1:D${r}`;
    ws['!cols'] = [W(36),W(18),W(16),W(48)];
    ws['!merges'] = [mergeRange(0,0,0,3),mergeRange(1,0,1,3),mergeRange(2,0,2,3)];
    ws['!rows'] = [{hpt:36},{hpt:20},{hpt:20}];
    ws['!freeze'] = { xSplit: 0, ySplit: 3 };
    XLSX.utils.book_append_sheet(wb, ws, 'Balance Sheet');
  }

  // ── Sheet 5: Cash Flow ──────────────────────────────────────────────────────
  {
    const ws: CS = {};
    let r = 0;

    const put = (row: number, col: number, v: string|number, t: string, style: CS) => {
      ws[`${colLetter(col)}${row}`] = { v, t, s: style };
    };

    put(1,0,`CASH FLOW STATEMENT — ${companyName}`,'s',s.title(C.NAVY));
    put(2,0,genDate,'s',s.subtitle(C.NAVY));
    put(3,0,'Indirect Method — Reconciles Net Profit to Operating Cash Flow','s',s.subtitle(C.NAVY));

    r=5;
    put(r,0,'OPERATING ACTIVITIES','s',s.sectionHdr(C.NAVY2));
    put(r,1,'Amount','s',s.hdr(C.NAVY2));

    r++;
    put(r,0,'  Net Profit (from P&L)','s',s.cell(false,C.BLACK,'left',C.BLUE_LT));
    put(r,1,cf.netProfit,'n',s.money(true,C.NAVY,C.BLUE_LT));

    r++;
    put(r,0,'  Adjustments to reconcile Net Profit:','s',s.cell(false,C.DGRAY));

    cf.adjustments.forEach((adj, i) => {
      r++;
      const bg = i%2===0 ? C.WHITE : C.OFF_WHT;
      put(r,0,`    ${adj.account}`,'s',s.cell(false,C.BLACK,'left',bg));
      put(r,1,adj.impact,'n',s.money(false,adj.impact>=0?C.GREEN:C.RED,bg));
    });

    r++;
    const ocfColor = cf.operatingCashFlow >= 0 ? C.GREEN : C.RED;
    put(r,0,'NET OPERATING CASH FLOW','s',s.totalLabel(ocfColor));
    put(r,1,cf.operatingCashFlow,'n',s.total(ocfColor));

    // Note about linkage
    r+=2;
    put(r,0,'📌 Note: Net Profit figure sourced from P&L tab. Full transactions in Raw Ledger tab.','s',{
      font:{sz:9,name:'Calibri',color:{rgb:C.DGRAY},italic:true},
      alignment:{horizontal:'left'},
    });

    ws['!ref'] = `A1:B${r}`;
    ws['!cols'] = [W(40),W(20)];
    ws['!merges'] = [mergeRange(0,0,0,1),mergeRange(1,0,1,1),mergeRange(2,0,2,1)];
    ws['!rows'] = [{hpt:36},{hpt:20},{hpt:16}];
    XLSX.utils.book_append_sheet(wb, ws, 'Cash Flow');
  }

  // ── Sheet 6: Month-over-Month ───────────────────────────────────────────────
  if (mom && mom.months.length > 0) {
    const ws: CS = {};
    const months = mom.months;
    const numCols = months.length + 2; // Account + months + Total

    const put = (row: number, col: number, v: string|number, t: string, style: CS) => {
      ws[`${colLetter(col)}${row}`] = { v, t, s: style };
    };

    // Title
    put(1,0,`MONTH-OVER-MONTH P&L — ${companyName}`,'s',s.title(C.NAVY));
    put(2,0,`${monthLabel(months[0])} → ${monthLabel(months[months.length-1])}`,'s',s.subtitle(C.NAVY));
    put(3,0,`${months.length} months · All figures in USD · Source: ${RAW_SHEET} tab`,'s',s.subtitle(C.NAVY));

    // Column headers
    const HDR_ROW = 5;
    put(HDR_ROW,0,'Account','s',s.hdr(C.NAVY,C.WHITE,10));
    months.forEach((m,i) => put(HDR_ROW,i+1,monthLabel(m),'s',s.hdr(C.NAVY,C.WHITE,10)));
    put(HDR_ROW,months.length+1,'Total','s',s.hdr(C.NAVY,C.WHITE,10));

    // MoM % row
    put(HDR_ROW+1,0,'MoM Change %','s',s.hdr(C.TEAL,C.WHITE,9));
    months.forEach((m,i) => {
      if (i === 0) { put(HDR_ROW+1,1,'Baseline','s',s.hdr(C.TEAL,C.WHITE,9)); return; }
      const cur = mom.monthlyRevenue[m]??0;
      const prev = mom.monthlyRevenue[months[i-1]]??0;
      const pct = prev!==0?(cur-prev)/Math.abs(prev):0;
      put(HDR_ROW+1,i+1,pct,'n',{...s.pct(false,pct>=0?C.GREEN:C.RED,C.YLW_LT),numFmt:'+0.0%;-0.0%;"-"'});
    });
    put(HDR_ROW+1,months.length+1,'','s',s.hdr(C.TEAL,C.WHITE,9));

    let r = HDR_ROW + 2;

    // Income section
    put(r,0,'▸  INCOME','s',{...s.sectionHdr(C.GREEN2),font:{bold:true,sz:10,name:'Calibri',color:{rgb:C.WHITE}}});
    for(let c=1;c<numCols;c++) ws[`${colLetter(c)}${r}`]={v:'',t:'s',s:{fill:{fgColor:{rgb:C.GREEN2},patternType:'solid'}}};
    r++;

    mom.incomeCategories.forEach((cat, ri) => {
      const bg = ri%2===0 ? C.WHITE : C.OFF_WHT;
      put(r,0,`  ${cat.name}`,'s',s.cell(false,C.BLACK,'left',bg));
      months.forEach((m,i) => { const v=cat.months[m]??0; put(r,i+1,v,'n',s.money(false,v>0?C.BLACK:C.LGRAY,bg)); });
      put(r,months.length+1,cat.total,'n',s.money(true,C.BLACK,bg));
      r++;
    });

    // Revenue total
    put(r,0,'Total Revenue','s',s.totalLabel(C.GREEN));
    months.forEach((m,i) => put(r,i+1,mom.monthlyRevenue[m]??0,'n',s.total(C.GREEN)));
    put(r,months.length+1,mom.totalRevenue,'n',s.total(C.GREEN));
    r+=2;

    // Expense section
    put(r,0,'▸  EXPENSES','s',{...s.sectionHdr(C.RED),font:{bold:true,sz:10,name:'Calibri',color:{rgb:C.WHITE}}});
    for(let c=1;c<numCols;c++) ws[`${colLetter(c)}${r}`]={v:'',t:'s',s:{fill:{fgColor:{rgb:C.RED},patternType:'solid'}}};
    r++;

    mom.expenseCategories.forEach((cat, ri) => {
      const bg = ri%2===0 ? C.WHITE : C.OFF_WHT;
      put(r,0,`  ${cat.name}`,'s',s.cell(false,C.BLACK,'left',bg));
      months.forEach((m,i) => { const v=cat.months[m]??0; put(r,i+1,v,'n',s.money(false,v>0?C.BLACK:C.LGRAY,bg)); });
      put(r,months.length+1,cat.total,'n',s.money(true,C.BLACK,bg));
      r++;
    });

    // Expense total
    put(r,0,'Total Expenses','s',s.totalLabel(C.RED));
    months.forEach((m,i) => put(r,i+1,mom.monthlyExpenses[m]??0,'n',s.total(C.RED)));
    put(r,months.length+1,mom.totalExpenses,'n',s.total(C.RED));
    r+=2;

    // Net Profit
    const npColor = mom.totalNetProfit>=0?C.GREEN:C.RED;
    put(r,0,'NET PROFIT','s',s.totalLabel(npColor));
    months.forEach((m,i) => {
      const v=mom.monthlyNetProfit[m]??0;
      put(r,i+1,v,'n',s.total(v>=0?C.GREEN:C.RED));
    });
    put(r,months.length+1,mom.totalNetProfit,'n',s.total(npColor));

    ws['!ref'] = `A1:${colLetter(numCols-1)}${r}`;
    ws['!cols'] = [W(32), ...months.map(()=>W(13)), W(15)];
    ws['!merges'] = [mergeRange(0,0,0,numCols-1),mergeRange(1,0,1,numCols-1),mergeRange(2,0,2,numCols-1)];
    ws['!freeze'] = { xSplit: 1, ySplit: HDR_ROW };
    ws['!rows'] = [{hpt:36},{hpt:20},{hpt:16}];
    XLSX.utils.book_append_sheet(wb, ws, 'Month-over-Month');
  }

  // ── Sheet 7: Flags ──────────────────────────────────────────────────────────
  {
    const ws: CS = {};
    let r = 0;

    const put = (row: number, col: number, v: string|number, t: string, style: CS) => {
      ws[`${colLetter(col)}${row}`] = { v, t, s: style };
    };

    put(1,0,`FLAGS & AUDIT NOTES — ${companyName}`,'s',s.title(C.RED));
    put(2,0,`${analysis.inconsistentVendors.length} inconsistent vendor(s) · ${analysis.duplicates.length} duplicate transaction(s)`,'s',s.subtitle(C.RED));

    r=4;
    put(r,0,'INCONSISTENT VENDORS','s',s.sectionHdr(C.GOLD));
    ['Vendor','Reason','Accounts Affected'].forEach((h,i) => put(r,i+0,h,'s',s.hdr(i===0?C.GOLD:C.NAVY2,C.WHITE,10)));

    // re-draw first col
    put(r,0,'INCONSISTENT VENDORS','s',s.hdr(C.GOLD,C.WHITE,10));

    analysis.inconsistentVendors.forEach((v, i) => {
      r++;
      const bg = i%2===0 ? C.YLW_LT : C.WHITE;
      put(r,0,v.vendor,'s',s.cell(true,C.BLACK,'left',bg));
      put(r,1,v.reason,'s',s.cell(false,C.BLACK,'left',bg));
      put(r,2,v.accounts.join(', '),'s',s.cell(false,C.BLACK,'left',bg));
    });
    if (analysis.inconsistentVendors.length === 0) { r++; put(r,0,'No inconsistent vendors found ✓','s',s.cell(false,C.GREEN)); }

    r+=2;
    put(r,0,'DUPLICATE TRANSACTIONS','s',s.hdr(C.RED,C.WHITE,10));
    ['Vendor','Amount','Date','Occurrences'].forEach((h,i) => put(r,i,h,'s',s.hdr(i===0?C.RED:C.NAVY2,C.WHITE,10)));
    put(r,0,'DUPLICATE TRANSACTIONS','s',s.hdr(C.RED,C.WHITE,10));

    analysis.duplicates.forEach((d, i) => {
      r++;
      const bg = i%2===0 ? C.RED_LT : C.WHITE;
      put(r,0,d.name,'s',s.cell(true,C.BLACK,'left',bg));
      put(r,1,d.amount,'s',s.cell(false,C.BLACK,'right',bg));
      put(r,2,d.transactionDate,'s',s.cell(false,C.BLACK,'center',bg));
      put(r,3,d.occurrences,'n',s.cell(true,C.RED,'center',bg));
    });
    if (analysis.duplicates.length === 0) { r++; put(r,0,'No duplicate transactions found ✓','s',s.cell(false,C.GREEN)); }

    ws['!ref'] = `A1:D${r}`;
    ws['!cols'] = [W(28),W(36),W(44),W(14)];
    ws['!merges'] = [mergeRange(0,0,0,3),mergeRange(1,0,1,3)];
    ws['!rows'] = [{hpt:36},{hpt:16}];
    XLSX.utils.book_append_sheet(wb, ws, 'Flags');
  }

  XLSX.writeFile(wb, `${base(fileName)}_Financial_Analysis.xlsx`);
}

// ─── PDF export ───────────────────────────────────────────────────────────────
export function exportPdf(fileName: string, analysis: LedgerAnalysis, pl: ProfitAndLoss, bs: BalanceSheet, cf: CashFlowStatement): void {
  const tbl = (hdrs: string[], rows: (string|number)[][]) =>
    `<table><thead><tr>${hdrs.map(h=>`<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map(c=>`<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table>`;

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"/><title>${base(fileName)}</title>
<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;font-size:12px;color:#111;padding:32px}
h1{font-size:20px;margin-bottom:4px}.meta{color:#666;font-size:11px;margin-bottom:24px}
.section{margin-bottom:28px;page-break-inside:avoid}h2{font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#1F3864;border-bottom:2px solid #1F3864;padding-bottom:4px;margin-bottom:10px}
table{width:100%;border-collapse:collapse;font-size:11px}th{background:#1F3864;color:#fff;text-align:left;padding:5px 8px;font-weight:600}
td{padding:4px 8px;border-bottom:1px solid #e5e7eb}tr:nth-child(even) td{background:#f1f5f9}
.kpi{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:12px}.kpi-card{border:1px solid #e5e7eb;border-radius:6px;padding:10px 16px;min-width:140px}
.kpi-label{font-size:10px;color:#666;text-transform:uppercase;letter-spacing:.06em}.kpi-value{font-size:16px;font-weight:700;margin-top:2px}
.balanced{background:#d1fae5;color:#065f46}.not-balanced{background:#fee2e2;color:#991b1b}.pill{display:inline-block;padding:2px 10px;border-radius:999px;font-size:11px;font-weight:600;margin-left:8px}
@media print{body{padding:16px}}</style></head><body>
<h1>${base(fileName)} — Financial Analysis Report</h1>
<p class="meta">Generated: ${new Date().toLocaleString()} · Transactions: ${analysis.totalTransactions}</p>
<div class="section"><h2>Profit & Loss</h2><div class="kpi">
<div class="kpi-card"><div class="kpi-label">Revenue</div><div class="kpi-value">${$(pl.totalRevenue)}</div></div>
<div class="kpi-card"><div class="kpi-label">Gross Profit</div><div class="kpi-value">${$(pl.grossProfit)} (${(pl.grossMargin*100).toFixed(1)}%)</div></div>
<div class="kpi-card"><div class="kpi-label">Net Profit</div><div class="kpi-value">${$(pl.netProfit)} (${(pl.netMargin*100).toFixed(1)}%)</div></div>
</div>${tbl(['Month','Revenue','Gross Profit','Net'],Object.entries(pl.monthlyBreakdown).map(([m,{revenue,cogs,expenses}])=>[m,$(revenue),$(revenue-(cogs??0)),$(revenue-(cogs??0)-expenses)]))}</div>
<div class="section"><h2>Balance Sheet <span class="pill ${bs.isBalanced?'balanced':'not-balanced'}">${bs.isBalanced?'Balanced ✓':'Not Balanced ✗'}</span></h2>
${tbl(['Category','Account','Balance'],[...bs.assets.map(({account,value})=>['Asset',account,$(value)]),...bs.liabilities.map(({account,value})=>['Liability',account,$(value)]),...bs.equity.map(({account,value})=>['Equity',account,$(value)])])}</div>
<div class="section"><h2>Cash Flow</h2>${tbl(['Item','Amount'],[['Net Profit',$(cf.netProfit)],['Operating CF',$(cf.operatingCashFlow)]])}</div>
</body></html>`;

  const win = window.open('','_blank');
  if (!win) return;
  win.document.write(html); win.document.close();
  win.onload = () => win.print();
}

function base(f: string) { return f.replace(/\.[^.]+$/,'') || 'ledger'; }
