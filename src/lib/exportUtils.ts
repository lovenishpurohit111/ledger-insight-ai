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

// ── Formatters ────────────────────────────────────────────────────────────────
const $f  = (n: number) => new Intl.NumberFormat('en-US',{style:'currency',currency:'USD'}).format(n);
const pf  = (n: number) => `${(n*100).toFixed(1)}%`;
const base = (f: string) => f.replace(/\.[^.]+$/,'') || 'ledger';

const toNum = (s: string): number => {
  if (!s) return 0;
  const neg = s.trim().startsWith('(') || s.trim().startsWith('-');
  const n = parseFloat(s.replace(/[^0-9.]/g,''));
  return isNaN(n) ? 0 : neg ? -n : n;
};

const toIsoDate = (s: string): string => {
  if (!s.trim()) return s;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s.trim())) return s.trim();
  const us = s.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (us) return `${us[3]}-${us[1].padStart(2,'0')}-${us[2].padStart(2,'0')}`;
  try { const d=new Date(s); if(!isNaN(d.getTime())) return d.toISOString().slice(0,10); } catch{/**/}
  return s;
};

// ── Modern Minimal Palette ────────────────────────────────────────────────────
const P = {
  // Charcoal / dark blue - primary
  INK:    '1A1F2E',  // Darkest - titles
  DARK:   '2C3E50',  // Dark - section headers
  MID:    '34495E',  // Mid - sub headers
  // Grays
  RULE:   'DDE1E7',  // Very light rule lines
  RULE2:  'BFC6D0',  // Slightly stronger rule
  SLATE:  'F0F2F5',  // Background rows
  WHITE:  'FFFFFF',
  OFFWHT: 'FAFBFC',  // Near-white
  GHOST:  'E8EAED',  // Disabled / empty cells
  TEXT:   '2D3748',  // Body text
  MUTED:  '718096',  // Secondary text
  // Accent colors - soft & professional
  GRN:    '276749',  // Dark green - profit labels
  GRN_LT: 'C6F6D5',  // Light green - profit cells
  GRN_MD: '48BB78',  // Mid green - positive values
  RED:    '9B2C2C',  // Dark red - loss labels
  RED_LT: 'FED7D7',  // Light red - loss cells
  RED_MD: 'FC8181',  // Mid red - negative values
  AMBER:  '7B341E',  // Dark amber
  AMB_LT: 'FEEBC8',  // Light amber
  BLUE:   '2B6CB0',  // Link blue
  BLU_LT: 'EBF4FF',  // Light blue bg
  // Chart colors (data bars)
  BAR1:   '3182CE',  // Revenue bar
  BAR2:   '48BB78',  // Profit bar
  BAR3:   'E53E3E',  // Expense bar
};

// ── Font presets ──────────────────────────────────────────────────────────────
const F = (bold: boolean, sz: number, rgb: string, italic=false) =>
  ({bold, sz, name:'Calibri', color:{rgb}, italic});

// ── Fill ──────────────────────────────────────────────────────────────────────
const Fill = (rgb: string): CS => ({fgColor:{rgb}, patternType:'solid'});
const noFill = (): CS => ({patternType:'none'});

// ── Alignment ─────────────────────────────────────────────────────────────────
const AL = (h: string, v='center', wrap=false): CS => ({horizontal:h, vertical:v, wrapText:wrap});

// ── Borders – minimal approach ────────────────────────────────────────────────
const bNone   = {};
const bBot    = (c=P.RULE)  => ({bottom:{style:'thin',  color:{rgb:c}}});
const bBotMed = (c=P.RULE2) => ({bottom:{style:'medium',color:{rgb:c}}});
const bBox    = (c=P.RULE)  => ({top:{style:'thin',color:{rgb:c}},bottom:{style:'thin',color:{rgb:c}},left:{style:'thin',color:{rgb:c}},right:{style:'thin',color:{rgb:c}}});
const bLeft   = (c=P.RULE)  => ({left:{style:'thin',color:{rgb:c}}});

// ── Number formats ────────────────────────────────────────────────────────────
const FMT_MONEY  = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"_);_(@_)';
const FMT_MONEY0 = '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)';
const FMT_PCT    = '0.0%';
const FMT_NUM    = '#,##0';

// ── Style presets ─────────────────────────────────────────────────────────────
const S = {
  // Page title — ink on white, large
  pageTitle: (): CS => ({
    font: F(true,16,P.INK), fill: Fill(P.WHITE),
    alignment: AL('left','center'),
  }),
  pageSub: (): CS => ({
    font: F(false,9,P.MUTED,true), fill: Fill(P.WHITE),
    alignment: AL('left','center'),
  }),
  // Section label — white on dark
  secHdr: (bg=P.DARK): CS => ({
    font: F(true,10,P.WHITE), fill: Fill(bg),
    alignment: AL('left','center'),
    border: bNone,
  }),
  // Column header — dark text on light bg
  colHdr: (): CS => ({
    font: F(true,9,P.MID), fill: Fill(P.SLATE),
    alignment: AL('right','center'),
    border: bBotMed(P.MID),
  }),
  colHdrL: (): CS => ({
    font: F(true,9,P.MID), fill: Fill(P.SLATE),
    alignment: AL('left','center'),
    border: bBotMed(P.MID),
  }),
  // Data row — alternating
  row: (bg=P.WHITE, fg=P.TEXT, bold=false): CS => ({
    font: F(bold,10,fg), fill: Fill(bg),
    alignment: AL('left','center'),
    border: bBot(),
  }),
  rowR: (bg=P.WHITE, fg=P.TEXT, bold=false): CS => ({
    font: F(bold,10,fg), fill: Fill(bg),
    numFmt: FMT_MONEY, alignment: AL('right','center'),
    border: bBot(),
  }),
  rowPct: (bg=P.WHITE, fg=P.MUTED): CS => ({
    font: F(false,9,fg), fill: Fill(bg),
    numFmt: FMT_PCT, alignment: AL('right','center'),
    border: bBot(),
  }),
  // Total row — bold with top rule
  total: (bg=P.DARK, fg=P.WHITE): CS => ({
    font: F(true,10,fg), fill: Fill(bg),
    numFmt: FMT_MONEY, alignment: AL('right','center'),
    border: {top:{style:'medium',color:{rgb:bg}},bottom:{style:'medium',color:{rgb:bg}}},
  }),
  totalL: (bg=P.DARK, fg=P.WHITE): CS => ({
    font: F(true,10,fg), fill: Fill(bg),
    alignment: AL('left','center'),
    border: {top:{style:'medium',color:{rgb:bg}},bottom:{style:'medium',color:{rgb:bg}}},
  }),
  // KPI card
  kpiLabel: (accent=P.BLUE): CS => ({
    font: F(true,8,P.MUTED), fill: Fill(P.WHITE),
    alignment: AL('left','bottom'),
    border: {top:{style:'medium',color:{rgb:accent}},left:{style:'thin',color:{rgb:P.RULE}},right:{style:'thin',color:{rgb:P.RULE}}},
  }),
  kpiValue: (fg=P.INK, accent=P.BLUE): CS => ({
    font: F(true,22,fg), fill: Fill(P.WHITE),
    numFmt: FMT_MONEY0, alignment: AL('left','top'),
    border: {bottom:{style:'medium',color:{rgb:accent}},left:{style:'thin',color:{rgb:P.RULE}},right:{style:'thin',color:{rgb:P.RULE}}},
  }),
  // Note / helper text
  note: (): CS => ({
    font: F(false,8,P.MUTED,true), fill: Fill(P.OFFWHT),
    alignment: AL('left','center'),
    border: bBot(P.GHOST),
  }),
};

// ── Helpers ───────────────────────────────────────────────────────────────────
const L  = (n: number) => n<26 ? String.fromCharCode(65+n) : String.fromCharCode(64+Math.floor(n/26))+String.fromCharCode(65+n%26);
const MG = (r:number,c:number,r2:number,c2:number) => ({s:{r,c},e:{r:r2,c:c2}});
const W  = (n: number) => ({wch:n});
const H  = (pt: number) => ({hpt:pt});

// Write helpers
const wv = (ws:CS,r:number,c:number,v:string|number,t:string,style:CS) => { ws[`${L(c)}${r}`]={v,t,s:style}; };
const wf = (ws:CS,r:number,c:number,f:string,v:number,style:CS) => { ws[`${L(c)}${r}`]={t:'n',f,v,s:style}; };
const fillRow = (ws:CS,r:number,from:number,to:number,bg:string) => {
  for(let c=from;c<=to;c++){const a=`${L(c)}${r}`;ws[a]=ws[a]??{v:'',t:'s'};ws[a].s={fill:Fill(bg)};}
};
const emptyRow = (ws:CS,r:number,from:number,to:number) => {
  for(let c=from;c<=to;c++){const a=`${L(c)}${r}`;ws[a]=ws[a]??{v:'',t:'s',s:{fill:Fill(P.WHITE)}};}
};

// ASCII bar (visual substitute for chart — character-based data bar)
const bar = (val: number, max: number, width=20): string => {
  if(max<=0) return '';
  const filled = Math.round((val/max)*width);
  return '█'.repeat(Math.max(0,filled)) + '░'.repeat(Math.max(0,width-filled));
};

// ── Formula builders ──────────────────────────────────────────────────────────
const RL = 'Raw Ledger';
let DEND = 5000;

const TINC  = ['Income','income','Revenue','revenue','Sales','sales','Other Income','other income'];
const TCOGS = ['Cost of Goods Sold','cost of goods sold','COGS','cogs','Cost of Sales'];
const TEXP  = ['Expense','expense','Expenses','expenses','Other Expense','other expense'];
const TAST  = ['Asset','asset','Bank','bank','Accounts Receivable (A/R)','Other Current Assets','Fixed Assets','Other Assets','Inventory'];
const TLIA  = ['Liability','liability','Accounts Payable (A/P)','Credit Card','Other Current Liabilities','Long Term Liabilities','Other Liability'];
const TEQ   = ['Equity','equity','Retained Earnings','Opening Balance Equity'];

const SUMIF_chain = (types:string[],col:'I'|'J') =>
  `IFERROR(${types.map(t=>`SUMIF('${RL}'!B$2:B$${DEND},"${t}",'${RL}'!${col}$2:${col}$${DEND})`).join('+')},0)`;

const SUMPRODUCT_month = (types:string[],mk:string) => {
  const tf = types.map(t=>`('${RL}'!B$2:B$${DEND}="${t}")`).join('+');
  return `IFERROR(SUMPRODUCT(((${tf})>0)*(LEFT('${RL}'!C$2:C$${DEND},7)="${mk}")*ISNUMBER('${RL}'!I$2:I$${DEND})*('${RL}'!I$2:I$${DEND})),0)`;
};

const SUMPRODUCT_acct = (acct:string,mk:string) =>
  `IFERROR(SUMPRODUCT(('${RL}'!A$2:A$${DEND}="${acct.replace(/"/g,'""')}")*(LEFT('${RL}'!C$2:C$${DEND},7)="${mk}")*ISNUMBER('${RL}'!I$2:I$${DEND})*('${RL}'!I$2:I$${DEND})),0)`;

const LOOKUP_bal = (acct:string) =>
  `IFERROR(LOOKUP(2,1/('${RL}'!A$2:A$${DEND}="${acct.replace(/"/g,'""')}"),'${RL}'!J$2:J$${DEND}),0)`;

// ═════════════════════════════════════════════════════════════════════════════
// CSV
// ═════════════════════════════════════════════════════════════════════════════
export function exportCsv(fileName:string,analysis:LedgerAnalysis,pl:ProfitAndLoss,bs:BalanceSheet,cf:CashFlowStatement):void{
  const e=(v:string|number)=>{const s=String(v);return s.includes(',')? `"${s}"`:s;};
  const sec=(t:string,h:string[],rows:(string|number)[][])=>[t,h.map(e).join(','),...rows.map(r=>r.map(e).join(',')),'']. join('\n');
  const out=[
    sec('P&L',['Item','Amount','Margin'],[['Revenue',$f(pl.totalRevenue),'100%'],['COGS',$f(pl.totalCogs),pf(pl.totalRevenue?pl.totalCogs/pl.totalRevenue:0)],['Gross Profit',$f(pl.grossProfit),pf(pl.grossMargin)],['OpEx',$f(pl.totalExpenses),pf(pl.totalRevenue?pl.totalExpenses/pl.totalRevenue:0)],['Net Profit',$f(pl.netProfit),pf(pl.netMargin)]]),
    '\n',
    sec('BS',['Cat','Account','Balance'],[...bs.assets.map(({account,value})=>['Asset',account,$f(value)]),...bs.liabilities.map(({account,value})=>['Liability',account,$f(value)]),...bs.equity.map(({account,value})=>['Equity',account,$f(value)])]),
  ].join('\n');
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([out],{type:'text/csv'}));a.download=`${base(fileName)}.csv`;a.click();
}

// ═════════════════════════════════════════════════════════════════════════════
// EXCEL — Modern Minimal Design
// ═════════════════════════════════════════════════════════════════════════════
export function exportExcel(
  fileName:string,analysis:LedgerAnalysis,pl:ProfitAndLoss,
  bs:BalanceSheet,cf:CashFlowStatement,mom?:MoMPL,rawRows?:LedgerRow[],
):void{
  const wb = XLSX.utils.book_new();
  const gd = new Date().toLocaleDateString('en-US',{year:'numeric',month:'long',day:'numeric'});
  const co = base(fileName);
  const rows = rawRows??[];
  DEND = rows.length>1 ? rows.length+1 : 5000;
  const HDRS=['Distribution account','Distribution account type','Transaction date','Transaction type','Num','Name','Description','Split','Amount','Balance'];

  // ══ 1. DASHBOARD ═══════════════════════════════════════════════════════════
  {
    const ws:CS={};
    const bsOK=bs.isBalanced;

    // Row 1: White space / top padding
    emptyRow(ws,1,0,9);
    ws['!rows']=[H(8)]; // row 1 padding

    // Row 2: Company name
    wv(ws,2,0,co,'s',{font:F(true,18,P.INK),fill:Fill(P.WHITE),alignment:AL('left','center')});
    wv(ws,2,6,`Generated ${gd}`,'s',{font:F(false,9,P.MUTED,true),fill:Fill(P.WHITE),alignment:AL('right','center')});
    emptyRow(ws,2,1,5); emptyRow(ws,2,7,9);

    // Row 3: Subtitle rule
    for(let c=0;c<=9;c++) ws[`${L(c)}3`]={v:'',t:'s',s:{fill:Fill(P.INK)}};

    // Row 4: spacer
    emptyRow(ws,4,0,9);

    // Rows 5-6: KPI CARDS (no background, just top-accent border)
    const kpis = [
      {label:'TOTAL REVENUE',     val:pl.totalRevenue,   fmt:FMT_MONEY0, fg:P.INK,    accent:P.BAR1},
      {label:'GROSS PROFIT',      val:pl.grossProfit,    fmt:FMT_MONEY0, fg:pl.grossProfit>=0?P.GRN:P.RED, accent:pl.grossProfit>=0?P.GRN:P.RED_MD},
      {label:'NET PROFIT',        val:pl.netProfit,      fmt:FMT_MONEY0, fg:pl.netProfit>=0?P.GRN:P.RED,  accent:pl.netProfit>=0?P.GRN:P.RED_MD},
      {label:'GROSS MARGIN',      val:pl.grossMargin,    fmt:FMT_PCT,    fg:P.INK,    accent:P.BAR1},
      {label:'NET MARGIN',        val:pl.netMargin,      fmt:FMT_PCT,    fg:P.INK,    accent:P.BAR1},
    ];
    kpis.forEach(({label,val,fmt,fg,accent},i)=>{
      const c=i*2;
      // Label row 5
      ws[`${L(c)}5`]={v:label,t:'s',s:{font:F(true,8,P.MUTED),fill:Fill(P.WHITE),alignment:AL('left','bottom'),border:{top:{style:'medium',color:{rgb:accent}},left:{style:'thin',color:{rgb:P.RULE}}}}};
      ws[`${L(c+1)}5`]={v:'',t:'s',s:{fill:Fill(P.WHITE),border:{top:{style:'medium',color:{rgb:accent}},right:{style:'thin',color:{rgb:P.RULE}}}}};
      // Value row 6
      ws[`${L(c)}6`]={v:val,t:'n',s:{font:F(true,22,fg),fill:Fill(P.WHITE),numFmt:fmt,alignment:AL('left','center'),border:{bottom:{style:'medium',color:{rgb:accent}},left:{style:'thin',color:{rgb:P.RULE}}}}};
      ws[`${L(c+1)}6`]={v:'',t:'s',s:{fill:Fill(P.WHITE),border:{bottom:{style:'medium',color:{rgb:accent}},right:{style:'thin',color:{rgb:P.RULE}}}}};
    });
    // BS Status card (cols 9-9 only, since we have 5 KPIs × 2 = cols 0-9)
    // Put BS status at col 9 only
    ws['J5']={v:'BALANCE SHEET',t:'s',s:{font:F(true,8,bsOK?P.GRN:P.RED),fill:Fill(P.WHITE),alignment:AL('left','bottom'),border:{top:{style:'medium',color:{rgb:bsOK?P.GRN:P.RED_MD}},left:{style:'thin',color:{rgb:P.RULE}},right:{style:'thin',color:{rgb:P.RULE}}}}};
    ws['J6']={v:bsOK?'✔  Balanced':'✘  Off',t:'s',s:{font:F(true,13,bsOK?P.GRN:P.RED),fill:Fill(bsOK?P.GRN_LT:P.RED_LT),alignment:AL('center','center'),border:{bottom:{style:'medium',color:{rgb:bsOK?P.GRN:P.RED_MD}},left:{style:'thin',color:{rgb:P.RULE}},right:{style:'thin',color:{rgb:P.RULE}}}}};

    // Row 7: spacer
    emptyRow(ws,7,0,9);

    // ── P&L Summary (rows 8-14) ──────────────────────────────────────────────
    wv(ws,8,0,'Profit & Loss','s',{font:F(true,11,P.WHITE),fill:Fill(P.DARK),alignment:AL('left','center')});
    wv(ws,8,1,'Amount','s',{font:F(true,9,P.WHITE),fill:Fill(P.DARK),alignment:AL('right','center')});
    wv(ws,8,2,'Margin','s',{font:F(true,9,P.WHITE),fill:Fill(P.DARK),alignment:AL('right','center')});
    wv(ws,8,3,'vs Total','s',{font:F(true,9,P.WHITE),fill:Fill(P.DARK),alignment:AL('left','center')});
    emptyRow(ws,8,4,4);

    const plRows:[string,number,boolean,string][]=[
      ['Revenue',     pl.totalRevenue,  false, ''],
      ['Cost of Goods',pl.totalCogs,   false, ''],
      ['Gross Profit',pl.grossProfit,  true,  ''],
      ['Op Expenses', pl.totalExpenses,false, ''],
      ['Net Profit',  pl.netProfit,    true,  ''],
    ];
    const maxPL = Math.max(pl.totalRevenue, pl.grossProfit, pl.totalExpenses, Math.abs(pl.netProfit));
    plRows.forEach(([label,val,bold],i)=>{
      const r=9+i;
      const rb=i%2===0?P.WHITE:P.OFFWHT;
      const isProfit=label==='Gross Profit'||label==='Net Profit';
      const fg=isProfit?(val>=0?P.GRN:P.RED):P.TEXT;
      const bg=isProfit?(val>=0?P.GRN_LT:P.RED_LT):rb;
      wv(ws,r,0,`  ${label}`,'s',S.row(bg,fg,bold));
      ws[`B${r}`]={v:val,t:'n',s:{...S.rowR(bg,fg,bold),numFmt:FMT_MONEY,font:F(bold,10,fg),fill:Fill(bg)}};
      ws[`C${r}`]={v:pl.totalRevenue?val/pl.totalRevenue:0,t:'n',s:S.rowPct(bg,isProfit?fg:P.MUTED)};
      // Mini bar chart
      const barStr=bar(Math.abs(val),maxPL,16);
      wv(ws,r,3,barStr,'s',{font:{name:'Calibri',sz:8,color:{rgb:isProfit?(val>=0?P.GRN_MD:P.RED_MD):P.BAR1}},fill:Fill(bg),alignment:AL('left','center')});
    });

    // P&L Totals line
    const totR=14;
    wv(ws,totR,0,'','s',{fill:Fill(P.RULE)});wv(ws,totR,1,'','s',{fill:Fill(P.RULE)});wv(ws,totR,2,'','s',{fill:Fill(P.RULE)});wv(ws,totR,3,'','s',{fill:Fill(P.RULE)});

    // ── BS Snapshot (rows 8-14, cols 5-9) ─────────────────────────────────────
    wv(ws,8,5,'Balance Sheet','s',{font:F(true,11,P.WHITE),fill:Fill(P.MID),alignment:AL('left','center')});
    wv(ws,8,6,'Balance','s',{font:F(true,9,P.WHITE),fill:Fill(P.MID),alignment:AL('right','center')});
    emptyRow(ws,8,7,9);

    [
      {label:'Assets',      val:bs.totals.assetsTotal,      fg:P.TEXT},
      {label:'Liabilities', val:bs.totals.liabilitiesTotal, fg:P.AMBER},
      {label:'Equity',      val:bs.totals.equityTotal,      fg:P.GRN},
      {label:'Variance',    val:bs.variance,                fg:bsOK?P.GRN:P.RED},
    ].forEach(({label,val,fg},i)=>{
      const r=9+i; const rb=i%2===0?P.WHITE:P.OFFWHT;
      const isVar=i===3; const vbg=isVar?(bsOK?P.GRN_LT:P.RED_LT):rb;
      wv(ws,r,5,label,'s',S.row(vbg,fg,isVar));
      ws[`G${r}`]={v:val,t:'n',s:{...S.rowR(vbg,fg,isVar),numFmt:FMT_MONEY}};
      emptyRow(ws,r,7,9);
    });

    // ── Flags (rows 16-20) ────────────────────────────────────────────────────
    emptyRow(ws,15,0,9);
    wv(ws,16,0,'Audit Summary','s',{font:F(true,11,P.WHITE),fill:Fill(P.DARK),alignment:AL('left','center')});
    emptyRow(ws,16,1,9);

    const flags=[
      {label:'Inconsistent Vendors',  val:analysis.inconsistentVendors.length, bad:analysis.inconsistentVendors.length>0},
      {label:'Duplicate Transactions',val:analysis.duplicates.length,           bad:analysis.duplicates.length>0},
      {label:'Total Transactions',    val:analysis.totalTransactions,           bad:false},
      {label:'Months Analysed',       val:mom?.months?.length??0,              bad:false},
    ];
    flags.forEach(({label,val,bad},i)=>{
      const r=17+i; const rb=i%2===0?P.WHITE:P.OFFWHT; const vbg=bad?P.AMB_LT:rb;
      wv(ws,r,0,label,'s',S.row(vbg,bad?P.AMBER:P.TEXT));
      ws[`B${r}`]={v:val,t:'n',s:{...S.row(vbg,bad?P.AMBER:P.TEXT),numFmt:FMT_NUM,alignment:AL('right','center')}};
      wv(ws,r,2,bad?'⚠ Review':'✔ OK','s',{font:F(true,9,bad?P.AMBER:P.GRN),fill:Fill(vbg),alignment:AL('left','center'),border:bBot()});
      emptyRow(ws,r,3,9);
    });

    ws['!ref']='A1:J21';
    ws['!cols']=[W(20),W(16),W(10),W(20),W(2),W(16),W(16),W(4),W(4),W(16)];
    ws['!rows']=[H(6),H(36),H(4),H(6),H(18),H(40),H(6),H(22),H(20),H(20),H(20),H(20),H(20),H(4),H(4),H(6),H(22),H(20),H(20),H(20),H(20)];
    ws['!merges']=[
      MG(1,0,1,9), // title
      MG(1,0,1,5), MG(1,6,1,9), // subtitle split
      MG(2,0,2,9), // separator
      MG(3,0,3,9), // spacer
      // KPI label rows (row 5 = index 4)
      MG(4,0,4,1),MG(4,2,4,3),MG(4,4,4,5),MG(4,6,4,7),MG(4,8,4,8),
      // KPI value rows (row 6 = index 5)
      MG(5,0,5,1),MG(5,2,5,3),MG(5,4,5,5),MG(5,6,5,7),MG(5,8,5,8),
      MG(6,0,6,9), // spacer
      // P&L section
      MG(7,0,7,4), MG(7,5,7,9),
      MG(15,0,15,9), // spacer
      MG(16,0,16,9), // audit header
      MG(17,2,17,9),MG(18,2,18,9),MG(19,2,19,9),MG(20,2,20,9),
    ];
    ws['!freeze']={xSplit:0,ySplit:4};
    XLSX.utils.book_append_sheet(wb,ws,'Dashboard');
  }

  // ══ 2. RAW LEDGER ══════════════════════════════════════════════════════════
  {
    const ws:CS={};
    // Thin header bar
    HDRS.forEach((_,i)=>ws[`${L(i)}1`]={v:'',t:'s',s:{fill:Fill(P.INK)}});
    wv(ws,1,0,`${co}  ·  ${rows.length.toLocaleString()} transactions`,'s',{font:F(true,11,P.WHITE),fill:Fill(P.INK),alignment:AL('left','center')});

    // Column headers
    const hdrColors=['Distribution Account','Type','Date','Txn Type','#','Name','Description','Split','Amount','Balance'];
    hdrColors.forEach((h,i)=>{
      ws[`${L(i)}2`]={v:h,t:'s',s:{font:F(true,9,P.MID),fill:Fill(P.SLATE),alignment:AL(i>=8?'right':'left','center'),border:bBotMed(P.RULE2)}};
    });

    // Data
    rows.forEach((row,ri)=>{
      const rn=ri+3; const rb=ri%2===0?P.WHITE:P.OFFWHT;
      HDRS.forEach((h,ci)=>{
        const raw=row[h as keyof typeof row]??'';
        const isAmt=ci===8||ci===9; const isDate=ci===2;
        if(isAmt){
          const num=toNum(String(raw));
          const fg=num<0?P.RED:num>0?P.TEXT:P.MUTED;
          ws[`${L(ci)}${rn}`]={v:num,t:'n',s:{font:F(false,9,fg),fill:Fill(rb),numFmt:FMT_MONEY,alignment:AL('right','center'),border:bBot()}};
        } else if(isDate){
          wv(ws,rn,ci,toIsoDate(String(raw)),'s',{font:F(false,9,P.MUTED),fill:Fill(rb),alignment:AL('center','center'),border:bBot()});
        } else {
          // Color-code type column
          const cellBg = ci===1 ? (()=>{
            const t=(String(raw)).toLowerCase();
            if(t.includes('income')||t.includes('revenue')) return P.GRN_LT;
            if(t.includes('expense')||t.includes('cost')) return P.RED_LT;
            if(t.includes('asset')||t.includes('bank')) return P.BLU_LT;
            if(t.includes('liabilit')) return P.AMB_LT;
            return rb;
          })() : rb;
          wv(ws,rn,ci,String(raw),'s',{font:F(false,9,P.TEXT),fill:Fill(cellBg),alignment:AL('left','center'),border:bBot()});
        }
      });
    });

    ws['!ref']=`A1:J${rows.length+2}`;
    ws['!cols']=[W(26),W(20),W(12),W(14),W(7),W(16),W(28),W(16),W(13),W(13)];
    ws['!rows']=[H(24),H(20)];
    ws['!merges']=[MG(0,0,0,9)];
    ws['!freeze']={xSplit:2,ySplit:2};
    ws['!autofilter']={ref:'A2:J2'};
    XLSX.utils.book_append_sheet(wb,ws,RL);
  }

  // ══ 3. P & L ═══════════════════════════════════════════════════════════════
  {
    const ws:CS={};
    const npC=pl.netProfit>=0?P.GRN:P.RED;
    const months=Object.entries(pl.monthlyBreakdown).sort();

    // Header
    emptyRow(ws,1,0,5);
    wv(ws,2,0,'Profit & Loss Statement','s',{font:F(true,16,P.INK),fill:Fill(P.WHITE),alignment:AL('left','center')});
    wv(ws,3,0,`${co}  ·  ${gd}`,'s',{font:F(false,9,P.MUTED,true),fill:Fill(P.WHITE),alignment:AL('left','center')});
    wv(ws,3,4,`← All values via SUMIF from '${RL}' tab`,'s',{font:F(false,8,P.MUTED,true),fill:Fill(P.WHITE),alignment:AL('right','center')});
    // Ink rule under title
    for(let c=0;c<=5;c++) ws[`${L(c)}4`]={v:'',t:'s',s:{fill:Fill(P.INK)}};
    emptyRow(ws,5,0,5);

    // ── Summary block (rows 6-12) ──
    wv(ws,6,0,'REVENUE','s',S.secHdr());
    wv(ws,7,0,'  Total Revenue','s',S.row(P.WHITE,P.TEXT,false));
    wf(ws,7,1,SUMIF_chain(TINC,'I'),pl.totalRevenue,{...S.rowR(),font:F(false,11,P.INK),numFmt:FMT_MONEY});
    wf(ws,7,2,'IF(B7=0,0,B7/B7)',1,S.rowPct());
    wv(ws,7,3,`← SUMIF Income/Revenue`,'s',S.note());

    emptyRow(ws,8,0,5);
    wv(ws,9,0,'COST OF GOODS','s',S.secHdr(P.MID));
    wv(ws,10,0,'  Cost of Goods Sold','s',S.row());
    wf(ws,10,1,SUMIF_chain(TCOGS,'I'),pl.totalCogs,{...S.rowR(),numFmt:FMT_MONEY});
    wf(ws,10,2,'IF(B7=0,0,B10/B7)',pl.totalRevenue?pl.totalCogs/pl.totalRevenue:0,S.rowPct());

    // Gross Profit
    wv(ws,11,0,'GROSS PROFIT','s',S.totalL(P.GRN_LT,P.GRN));
    wf(ws,11,1,'B7-B10',pl.grossProfit,{...S.total(P.GRN_LT,P.GRN),border:{top:{style:'thin',color:{rgb:P.GRN}},bottom:{style:'thin',color:{rgb:P.GRN}}}});
    wf(ws,11,2,'IF(B7=0,0,B11/B7)',pl.grossMargin,{...S.rowPct(P.GRN_LT,P.GRN),border:{top:{style:'thin',color:{rgb:P.GRN}},bottom:{style:'thin',color:{rgb:P.GRN}}}});
    wv(ws,11,3,`= Revenue − COGS`,'s',{...S.note(),fill:Fill(P.GRN_LT)});

    emptyRow(ws,12,0,5);
    wv(ws,13,0,'OPERATING EXPENSES','s',S.secHdr(P.MID));
    wv(ws,14,0,'  Total Operating Expenses','s',S.row());
    wf(ws,14,1,SUMIF_chain(TEXP,'I'),pl.totalExpenses,{...S.rowR(),numFmt:FMT_MONEY});
    wf(ws,14,2,'IF(B7=0,0,B14/B7)',pl.totalRevenue?pl.totalExpenses/pl.totalRevenue:0,S.rowPct());

    // Net Profit
    const npBg=pl.netProfit>=0?P.GRN_LT:P.RED_LT;
    wv(ws,15,0,'NET PROFIT / (LOSS)','s',S.totalL(npBg,npC));
    wf(ws,15,1,'B11-B14',pl.netProfit,{...S.total(npBg,npC),border:{top:{style:'thin',color:{rgb:npC}},bottom:{style:'thin',color:{rgb:npC}}}});
    wf(ws,15,2,'IF(B7=0,0,B15/B7)',pl.netMargin,{...S.rowPct(npBg,npC),border:{top:{style:'thin',color:{rgb:npC}},bottom:{style:'thin',color:{rgb:npC}}}});

    // Spacer
    emptyRow(ws,16,0,5);

    // ── Monthly breakdown ──────────────────────────────────────────────────
    wv(ws,17,0,'MONTHLY BREAKDOWN','s',S.secHdr());
    ['Month','Revenue','COGS','OpEx','Net Profit','Margin %'].forEach((h,i)=>{
      ws[`${L(i)}18`]={v:h,t:'s',s:{font:F(true,9,P.MID),fill:Fill(P.SLATE),alignment:AL(i===0?'left':'right','center'),border:bBotMed(P.RULE2)}};
    });

    const maxRev = Math.max(...months.map(([,{revenue}])=>revenue),1);
    months.forEach(([mk,{revenue,cogs,expenses}],i)=>{
      const r=19+i; const rb=i%2===0?P.WHITE:P.OFFWHT;
      const net=revenue-(cogs??0)-expenses;
      const netBg=net>=0?P.WHITE:P.RED_LT; const netFg=net>=0?P.GRN:P.RED;
      wv(ws,r,0,mk,'s',{...S.row(rb),font:F(false,9,P.MUTED)});
      wf(ws,r,1,SUMPRODUCT_month(TINC,mk),revenue,{...S.rowR(rb),numFmt:FMT_MONEY});
      wf(ws,r,2,SUMPRODUCT_month(TCOGS,mk),cogs??0,{...S.rowR(rb),numFmt:FMT_MONEY});
      wf(ws,r,3,SUMPRODUCT_month(TEXP,mk),expenses,{...S.rowR(rb),numFmt:FMT_MONEY});
      wf(ws,r,4,`${L(1)}${r}-${L(2)}${r}-${L(3)}${r}`,net,{...S.rowR(net>=0?rb:P.RED_LT,netFg),numFmt:FMT_MONEY});
      wf(ws,r,5,`IF(${L(1)}${r}=0,0,${L(4)}${r}/${L(1)}${r})`,revenue?net/revenue:0,S.rowPct(rb,net>=0?P.GRN_MD:P.RED_MD));
    });

    const totR=19+months.length;
    wv(ws,totR,0,'TOTAL','s',S.totalL());
    wf(ws,totR,1,`SUM(B19:B${totR-1})`,pl.totalRevenue,S.total());
    wf(ws,totR,2,`SUM(C19:C${totR-1})`,pl.totalCogs,S.total());
    wf(ws,totR,3,`SUM(D19:D${totR-1})`,pl.totalExpenses,S.total());
    wf(ws,totR,4,`SUM(E19:E${totR-1})`,pl.netProfit,{...S.total(pl.netProfit>=0?P.GRN_LT:P.RED_LT,npC),border:{top:{style:'medium',color:{rgb:npC}},bottom:{style:'medium',color:{rgb:npC}}}});
    wf(ws,totR,5,`IF(B${totR}=0,0,E${totR}/B${totR})`,pl.netMargin,S.rowPct(P.SLATE,pl.netProfit>=0?P.GRN:P.RED));

    ws['!ref']=`A1:F${totR}`;
    ws['!cols']=[W(30),W(16),W(12),W(16),W(16),W(10)];
    ws['!rows']=[H(6),H(28),H(18),H(4),H(6)];
    ws['!merges']=[MG(1,0,1,5),MG(2,0,2,3),MG(2,4,2,5),MG(3,0,3,5),MG(4,0,4,5),MG(5,0,5,5),MG(6,0,6,5),MG(8,0,8,5),MG(9,0,9,5),MG(12,0,12,5),MG(13,0,13,5),MG(16,0,16,5),MG(17,0,17,5)];
    ws['!freeze']={xSplit:0,ySplit:5};
    XLSX.utils.book_append_sheet(wb,ws,'P & L');
  }

  // ══ 4. BALANCE SHEET ═══════════════════════════════════════════════════════
  {
    const ws:CS={};
    let r=1;
    emptyRow(ws,r,0,4);
    r++;
    wv(ws,r,0,'Balance Sheet','s',{font:F(true,16,P.INK),fill:Fill(P.WHITE),alignment:AL('left','center')});
    r++;
    wv(ws,r,0,`${co}  ·  ${gd}`,'s',{font:F(false,9,P.MUTED,true),fill:Fill(P.WHITE),alignment:AL('left','center')});
    r++;
    for(let c=0;c<=4;c++) ws[`${L(c)}${r}`]={v:'',t:'s',s:{fill:Fill(P.INK)}};
    r++;
    emptyRow(ws,r,0,4);
    r++;

    // Column headers
    ['Account','Balance','% of Total','Type','Source'].forEach((h,i)=>{
      ws[`${L(i)}${r}`]={v:h,t:'s',s:{font:F(true,9,P.MID),fill:Fill(P.SLATE),alignment:AL(i===1?'right':i===2?'right':'left','center'),border:bBotMed(P.RULE2)}};
    });
    r++;

    let assetTotR=0,liabTotR=0,eqTotR=0;

    const writeBS=(title:string,entries:BalanceSheet['assets'],total:number,accent:string):number=>{
      for(let c=0;c<=4;c++) ws[`${L(c)}${r}`]={v:c===0?title:'',t:'s',s:{font:F(true,10,P.WHITE),fill:Fill(accent),alignment:AL('left','center')}};
      r++;
      const sr=r;
      entries.forEach((e,i)=>{
        const rb=i%2===0?P.WHITE:P.OFFWHT;
        const isCPE=e.account==='Current Period Earnings';
        const formula=isCPE?`'P & L'!B15`:LOOKUP_bal(e.account);
        const note=isCPE?`← Net Profit (P & L B15)`:e.account.slice(0,20);
        wv(ws,r,0,`  ${e.account}`,'s',S.row(rb));
        wf(ws,r,1,formula,e.value,{...S.rowR(rb),numFmt:FMT_MONEY});
        wf(ws,r,2,`IF(SUM(B${sr}:B${sr+entries.length-1})=0,0,B${r}/SUM(B${sr}:B${sr+entries.length-1}))`,total?e.value/total:0,S.rowPct(rb));
        wv(ws,r,3,e.isCurrent!==undefined?(e.isCurrent?'Current':'Non-Current'):'','s',{font:F(false,8,P.MUTED),fill:Fill(rb),alignment:AL('center','center'),border:bBot()});
        wv(ws,r,4,note,'s',S.note());
        r++;
      });
      // Total
      wv(ws,r,0,`  Total ${title}`,'s',{...S.totalL(P.SLATE,accent),font:F(true,10,accent)});
      wf(ws,r,1,`SUM(B${sr}:B${r-1})`,total,{...S.total(P.SLATE,accent),border:{top:{style:'thin',color:{rgb:accent}},bottom:{style:'thin',color:{rgb:accent}}}});
      wf(ws,r,2,'1',1,{...S.rowPct(P.SLATE,accent),border:{top:{style:'thin',color:{rgb:accent}},bottom:{style:'thin',color:{rgb:accent}}}});
      const totRow=r; r+=2;
      return totRow;
    };

    assetTotR = writeBS('Assets',      bs.assets,      bs.totals.assetsTotal,      P.BLUE);
    liabTotR  = writeBS('Liabilities', bs.liabilities, bs.totals.liabilitiesTotal, P.AMBER);
    eqTotR    = writeBS('Equity',      bs.equity,      bs.totals.equityTotal,      P.GRN);

    // Reconciliation
    for(let c=0;c<=4;c++) ws[`${L(c)}${r}`]={v:c===0?'RECONCILIATION':'',t:'s',s:{font:F(true,10,P.WHITE),fill:Fill(P.DARK),alignment:AL('left','center')}};
    r++;
    const varRow=r+2;
    [['Total Assets',bs.totals.assetsTotal,`B${assetTotR}`],
     ['Liabilities + Equity',bs.totals.liabilitiesTotal+bs.totals.equityTotal,`B${liabTotR}+B${eqTotR}`],
     ['Variance (A − L − E)',bs.variance,`B${assetTotR}-(B${liabTotR}+B${eqTotR})`],
    ].forEach(([lbl,val,fml],i)=>{
      const rb=i%2===0?P.WHITE:P.OFFWHT;
      const isVar=i===2; const vbg=isVar?(bsOK?P.GRN_LT:P.RED_LT):rb; const fg=isVar?(bsOK?P.GRN:P.RED):P.TEXT;
      wv(ws,r,0,String(lbl),'s',S.row(vbg,fg,isVar));
      wf(ws,r,1,String(fml),Number(val),{...S.rowR(vbg,fg,isVar),numFmt:FMT_MONEY});
      r++;
    });
    // Status formula-driven
    const bsOK=bs.isBalanced;
    ws[`A${r}`]={v:'Status',t:'s',s:S.row(P.WHITE,P.TEXT,true)};
    ws[`B${r}`]={t:'s',f:`IF(ABS(B${varRow})<=1,"✔  BALANCED","✘  Off by "&TEXT(ABS(B${varRow}),"$#,##0.00"))`,v:bsOK?'✔  BALANCED':`✘  Off by ${$f(Math.abs(bs.variance))}`,s:{font:F(true,11,bsOK?P.GRN:P.RED),fill:Fill(bsOK?P.GRN_LT:P.RED_LT),alignment:AL('left','center'),border:bBox(bsOK?P.GRN_MD:P.RED_MD)}};

    ws['!ref']=`A1:E${r+1}`;
    ws['!cols']=[W(32),W(16),W(10),W(14),W(28)];
    ws['!rows']=[H(6),H(28),H(16),H(4)];
    ws['!merges']=[MG(1,0,1,4),MG(2,0,2,4),MG(3,0,3,4),MG(4,0,4,4)];
    ws['!freeze']={xSplit:0,ySplit:7};
    XLSX.utils.book_append_sheet(wb,ws,'Balance Sheet');
  }

  // ══ 5. MONTH-OVER-MONTH ═══════════════════════════════════════════════════
  if(mom&&mom.months.length>0){
    const ws:CS={};
    const months=mom.months; const nC=months.length+2;

    emptyRow(ws,1,0,nC-1);
    wv(ws,2,0,'Month-over-Month P&L','s',{font:F(true,16,P.INK),fill:Fill(P.WHITE),alignment:AL('left','center')});
    wv(ws,3,0,`${co}  ·  ${monthLabel(months[0])} → ${monthLabel(months[months.length-1])}  ·  ${months.length} months`,'s',{font:F(false,9,P.MUTED,true),fill:Fill(P.WHITE),alignment:AL('left','center')});
    for(let c=0;c<=nC-1;c++) ws[`${L(c)}4`]={v:'',t:'s',s:{fill:Fill(P.INK)}};
    emptyRow(ws,5,0,nC-1);

    // Headers row 6
    ws['A6']={v:'Account',t:'s',s:S.colHdrL()};
    months.forEach((m,i)=>ws[`${L(i+1)}6`]={v:monthLabel(m),t:'s',s:S.colHdr()});
    ws[`${L(months.length+1)}6`]={v:'Total',t:'s',s:{...S.colHdr(),fill:Fill(P.SLATE)}};

    // MoM % row 7
    ws['A7']={v:'MoM Revenue Δ%',t:'s',s:{font:F(true,8,P.MUTED),fill:Fill(P.OFFWHT),alignment:AL('left','center'),border:bBotMed()}};
    months.forEach((_,i)=>{
      if(i===0){ws[`B7`]={v:'Baseline',t:'s',s:{font:F(false,8,P.MUTED),fill:Fill(P.OFFWHT),alignment:AL('center','center'),border:bBotMed()}};return;}
      const cur=mom.monthlyRevenue[months[i]]??0; const prv=mom.monthlyRevenue[months[i-1]]??0;
      const pct=prv!==0?(cur-prv)/Math.abs(prv):0;
      const revTotRow=8+mom.incomeCategories.length+1;
      ws[`${L(i+1)}7`]={t:'n',f:`IF(${L(i)}${revTotRow}=0,0,(${L(i+1)}${revTotRow}-${L(i)}${revTotRow})/ABS(${L(i)}${revTotRow}))`,v:pct,s:{font:F(true,9,cur>=prv?P.GRN:P.RED),fill:Fill(cur>=prv?P.GRN_LT:P.RED_LT),numFmt:'+0.0%;[Red]-0.0%;"—"',alignment:AL('right','center'),border:bBotMed()}};
    });
    ws[`${L(months.length+1)}7`]={v:'',t:'s',s:{fill:Fill(P.OFFWHT),border:bBotMed()}};

    let r=8;

    // Income
    for(let c=0;c<nC;c++) ws[`${L(c)}${r}`]={v:c===0?'  INCOME':'',t:'s',s:{font:F(true,9,P.WHITE),fill:Fill(P.GRN),alignment:AL('left','center')}};
    r++;
    const incS=r;
    mom.incomeCategories.forEach((cat,ri)=>{
      const rb=ri%2===0?P.WHITE:P.OFFWHT;
      wv(ws,r,0,`  ${cat.name}`,'s',S.row(rb));
      months.forEach((m,i)=>{const v=cat.months[m]??0;wf(ws,r,i+1,SUMPRODUCT_acct(cat.name,m),v,{...S.rowR(rb,v>0?P.TEXT:P.MUTED),numFmt:FMT_MONEY});});
      wf(ws,r,months.length+1,`SUM(${L(1)}${r}:${L(months.length)}${r})`,cat.total,{...S.rowR(rb),numFmt:FMT_MONEY,font:F(true,10,P.TEXT)});
      r++;
    });
    const revTR=r;
    ws[`A${r}`]={v:'Total Revenue',t:'s',s:S.totalL(P.GRN_LT,P.GRN)};
    months.forEach((_,i)=>wf(ws,r,i+1,`SUM(${L(i+1)}${incS}:${L(i+1)}${r-1})`,mom.monthlyRevenue[months[i]]??0,{...S.total(P.GRN_LT,P.GRN)}));
    wf(ws,r,months.length+1,`SUM(${L(1)}${r}:${L(months.length)}${r})`,mom.totalRevenue,S.total(P.GRN_LT,P.GRN));
    r+=2;

    // Expenses
    for(let c=0;c<nC;c++) ws[`${L(c)}${r}`]={v:c===0?'  EXPENSES':'',t:'s',s:{font:F(true,9,P.WHITE),fill:Fill(P.RED_MD),alignment:AL('left','center')}};
    r++;
    const expS=r;
    mom.expenseCategories.forEach((cat,ri)=>{
      const rb=ri%2===0?P.WHITE:P.OFFWHT;
      wv(ws,r,0,`  ${cat.name}`,'s',S.row(rb));
      months.forEach((m,i)=>{const v=cat.months[m]??0;wf(ws,r,i+1,SUMPRODUCT_acct(cat.name,m),v,{...S.rowR(rb,v>0?P.TEXT:P.MUTED),numFmt:FMT_MONEY});});
      wf(ws,r,months.length+1,`SUM(${L(1)}${r}:${L(months.length)}${r})`,cat.total,{...S.rowR(rb),numFmt:FMT_MONEY,font:F(true,10,P.TEXT)});
      r++;
    });
    const expTR=r;
    ws[`A${r}`]={v:'Total Expenses',t:'s',s:S.totalL(P.RED_LT,P.RED)};
    months.forEach((_,i)=>wf(ws,r,i+1,`SUM(${L(i+1)}${expS}:${L(i+1)}${r-1})`,mom.monthlyExpenses[months[i]]??0,S.total(P.RED_LT,P.RED)));
    wf(ws,r,months.length+1,`SUM(${L(1)}${r}:${L(months.length)}${r})`,mom.totalExpenses,S.total(P.RED_LT,P.RED));
    r+=2;

    const npC2=mom.totalNetProfit>=0?P.GRN:P.RED; const npBg=mom.totalNetProfit>=0?P.GRN_LT:P.RED_LT;
    ws[`A${r}`]={v:'NET PROFIT',t:'s',s:S.totalL(npBg,npC2)};
    months.forEach((_,i)=>{
      const v=mom.monthlyNetProfit[months[i]]??0; const vc=v>=0?P.GRN:P.RED; const vbg=v>=0?P.GRN_LT:P.RED_LT;
      wf(ws,r,i+1,`${L(i+1)}${revTR}-${L(i+1)}${expTR}`,v,{...S.total(vbg,vc)});
    });
    wf(ws,r,months.length+1,`${L(months.length+1)}${revTR}-${L(months.length+1)}${expTR}`,mom.totalNetProfit,S.total(npBg,npC2));

    ws['!ref']=`A1:${L(nC-1)}${r}`;
    ws['!cols']=[W(28),...months.map(()=>W(12)),W(14)];
    ws['!rows']=[H(6),H(28),H(16),H(4),H(6)];
    ws['!merges']=[MG(1,0,1,nC-1),MG(2,0,2,nC-1),MG(3,0,3,nC-1),MG(4,0,4,nC-1)];
    ws['!freeze']={xSplit:1,ySplit:6};
    XLSX.utils.book_append_sheet(wb,ws,'Month-over-Month');
  }

  // ══ 6. CASH FLOW ══════════════════════════════════════════════════════════
  {
    const ws:CS={};
    const ocfC=cf.operatingCashFlow>=0?P.GRN:P.RED;
    const ocfBg=cf.operatingCashFlow>=0?P.GRN_LT:P.RED_LT;

    emptyRow(ws,1,0,2);
    wv(ws,2,0,'Cash Flow Statement','s',{font:F(true,16,P.INK),fill:Fill(P.WHITE),alignment:AL('left','center')});
    wv(ws,3,0,`${co}  ·  ${gd}`,'s',{font:F(false,9,P.MUTED,true),fill:Fill(P.WHITE),alignment:AL('left','center')});
    for(let c=0;c<=2;c++) ws[`${L(c)}4`]={v:'',t:'s',s:{fill:Fill(P.INK)}};
    emptyRow(ws,5,0,2);

    for(let c=0;c<=2;c++) ws[`${L(c)}6`]={v:c===0?'OPERATING ACTIVITIES':'',t:'s',s:S.secHdr()};

    wv(ws,7,0,'  Net Profit','s',S.row(P.BLU_LT,P.BLUE,true));
    ws['B7']={t:'n',f:`'P & L'!B15`,v:pl.netProfit,s:{...S.rowR(P.BLU_LT,P.BLUE,true),numFmt:FMT_MONEY}};
    wv(ws,7,2,`← Linked from P & L tab, cell B15`,'s',S.note());

    wv(ws,8,0,'  Adjustments:','s',{font:F(false,9,P.MUTED,true),fill:Fill(P.WHITE),alignment:AL('left','center'),border:bBot()});
    emptyRow(ws,8,1,2);

    let r=9;
    cf.adjustments.forEach((adj,i)=>{
      const rb=i%2===0?P.WHITE:P.OFFWHT;
      wv(ws,r,0,`    ${adj.account}`,'s',S.row(rb));
      ws[`B${r}`]={v:adj.impact,t:'n',s:{...S.rowR(rb,adj.impact>=0?P.GRN:P.RED),numFmt:FMT_MONEY}};
      emptyRow(ws,r,2,2);
      r++;
    });

    emptyRow(ws,r,0,2);r++;
    ws[`A${r}`]={v:'NET OPERATING CASH FLOW',t:'s',s:S.totalL(ocfBg,ocfC)};
    wf(ws,r,1,`B7+SUM(B9:B${r-2})`,cf.operatingCashFlow,{...S.total(ocfBg,ocfC),border:{top:{style:'thin',color:{rgb:ocfC}},bottom:{style:'thin',color:{rgb:ocfC}}}});
    r+=2;
    wv(ws,r,0,`All data sourced from '${RL}' tab  ·  Net Profit = P & L!B15`,'s',S.note());

    ws['!ref']=`A1:C${r}`;
    ws['!cols']=[W(38),W(18),W(32)];
    ws['!rows']=[H(6),H(28),H(16),H(4)];
    ws['!merges']=[MG(1,0,1,2),MG(2,0,2,2),MG(3,0,3,2),MG(4,0,4,2),MG(5,0,5,2),MG(6,0,6,2)];
    ws['!freeze']={xSplit:0,ySplit:5};
    XLSX.utils.book_append_sheet(wb,ws,'Cash Flow');
  }

  // ══ 7. FLAGS ══════════════════════════════════════════════════════════════
  {
    const ws:CS={};
    let r=1;
    emptyRow(ws,r,0,4);r++;
    wv(ws,r,0,'Audit & Flags','s',{font:F(true,16,P.INK),fill:Fill(P.WHITE),alignment:AL('left','center')});r++;
    wv(ws,r,0,`${co}  ·  ${gd}`,'s',{font:F(false,9,P.MUTED,true),fill:Fill(P.WHITE),alignment:AL('left','center')});r++;
    for(let c=0;c<=4;c++) ws[`${L(c)}${r}`]={v:'',t:'s',s:{fill:Fill(P.INK)}};r++;
    emptyRow(ws,r,0,4);r++;

    // Inconsistent Vendors
    for(let c=0;c<=4;c++) ws[`${L(c)}${r}`]={v:c===0?'INCONSISTENT VENDORS':'',t:'s',s:S.secHdr()};r++;
    ['Vendor','Reason','Accounts','',''].forEach((h,i)=>ws[`${L(i)}${r}`]={v:h,t:'s',s:S.colHdrL()});r++;
    if(analysis.inconsistentVendors.length===0){
      wv(ws,r,0,'✔  No inconsistent vendors found','s',{font:F(false,10,P.GRN),fill:Fill(P.GRN_LT),alignment:AL('left','center'),border:bBot(P.GRN)});r++;
    } else {
      analysis.inconsistentVendors.forEach((v,i)=>{
        const rb=i%2===0?P.WHITE:P.OFFWHT;
        wv(ws,r,0,v.vendor,'s',S.row(rb,P.TEXT,true));
        wv(ws,r,1,v.reason,'s',S.row(rb,P.MUTED));
        wv(ws,r,2,v.accounts.join(', '),'s',S.row(rb,P.TEXT));
        emptyRow(ws,r,3,4);r++;
      });
    }

    emptyRow(ws,r,0,4);r++;
    for(let c=0;c<=4;c++) ws[`${L(c)}${r}`]={v:c===0?'DUPLICATE TRANSACTIONS':'',t:'s',s:S.secHdr(P.MID)};r++;
    ['Vendor','Amount','Date','Count',''].forEach((h,i)=>ws[`${L(i)}${r}`]={v:h,t:'s',s:S.colHdrL()});r++;
    if(analysis.duplicates.length===0){
      wv(ws,r,0,'✔  No duplicate transactions found','s',{font:F(false,10,P.GRN),fill:Fill(P.GRN_LT),alignment:AL('left','center'),border:bBot(P.GRN)});r++;
    } else {
      analysis.duplicates.forEach((d,i)=>{
        const rb=i%2===0?P.WHITE:P.OFFWHT;
        wv(ws,r,0,d.name,'s',S.row(rb,P.TEXT,true));
        ws[`B${r}`]={v:d.amount,t:'s',s:{...S.row(rb),alignment:AL('right','center')}};
        wv(ws,r,2,d.transactionDate,'s',{...S.row(rb,P.MUTED),alignment:AL('center','center')});
        ws[`D${r}`]={v:d.occurrences,t:'n',s:{...S.row(rb,P.RED),numFmt:FMT_NUM,alignment:AL('center','center')}};
        emptyRow(ws,r,4,4);r++;
      });
    }

    ws['!ref']=`A1:E${r}`;
    ws['!cols']=[W(24),W(34),W(40),W(10),W(10)];
    ws['!rows']=[H(6),H(28),H(16),H(4),H(6)];
    ws['!merges']=[MG(1,0,1,4),MG(2,0,2,4),MG(3,0,3,4),MG(4,0,4,4)];
    XLSX.utils.book_append_sheet(wb,ws,'Flags');
  }

  XLSX.writeFile(wb,`${base(fileName)}_Financial_Analysis.xlsx`);
}

// ═════════════════════════════════════════════════════════════════════════════
// PDF
// ═════════════════════════════════════════════════════════════════════════════
export function exportPdf(fileName:string,analysis:LedgerAnalysis,pl:ProfitAndLoss,bs:BalanceSheet,cf:CashFlowStatement):void{
  const tbl=(h:string[],rows:(string|number)[][])=>`<table><thead><tr>${h.map(x=>`<th>${x}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map(c=>`<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>${base(fileName)}</title><style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Segoe UI',Calibri,sans-serif;font-size:12px;color:#2D3748;padding:40px;background:#fff}h1{font-size:22px;color:#1A1F2E;margin-bottom:4px;font-weight:700}h2{font-size:13px;font-weight:600;color:#2C3E50;border-bottom:2px solid #1A1F2E;padding-bottom:4px;margin:20px 0 8px}.meta{color:#718096;font-size:10px;margin-bottom:24px}.sec{margin-bottom:20px;page-break-inside:avoid}table{width:100%;border-collapse:collapse;font-size:11px}th{background:#2C3E50;color:#fff;text-align:left;padding:6px 10px;font-weight:600;font-size:10px;letter-spacing:.04em}td{padding:5px 10px;border-bottom:1px solid #E8EAED}tr:nth-child(even)td{background:#F0F2F5}.pos{color:#276749}.neg{color:#9B2C2C}@media print{body{padding:20px}}</style></head><body>
<h1>${base(fileName)}</h1><p class="meta">Financial Analysis Report  ·  Generated ${new Date().toLocaleString()}  ·  ${analysis.totalTransactions} transactions</p>
<div class="sec"><h2>Profit & Loss</h2>${tbl(['','Amount','Margin'],[['Revenue',$f(pl.totalRevenue),'100%'],['Cost of Goods',$f(pl.totalCogs),pf(pl.totalRevenue?pl.totalCogs/pl.totalRevenue:0)],['Gross Profit',$f(pl.grossProfit),pf(pl.grossMargin)],['Op Expenses',$f(pl.totalExpenses),pf(pl.totalRevenue?pl.totalExpenses/pl.totalRevenue:0)],['Net Profit',$f(pl.netProfit),pf(pl.netMargin)]])}</div>
<div class="sec"><h2>Balance Sheet  ${bs.isBalanced?'<span class="pos">✔ Balanced</span>':'<span class="neg">✘ Not Balanced</span>'}</h2>${tbl(['Category','Account','Balance'],[...bs.assets.map(({account,value})=>['Asset',account,$f(value)]),...bs.liabilities.map(({account,value})=>['Liability',account,$f(value)]),...bs.equity.map(({account,value})=>['Equity',account,$f(value)])])}</div>
<div class="sec"><h2>Cash Flow</h2>${tbl(['','Amount'],[['Net Profit',$f(cf.netProfit)],['Operating Cash Flow',$f(cf.operatingCashFlow)]])}</div>
</body></html>`;
  const win=window.open('','_blank');if(!win)return;win.document.write(html);win.document.close();win.onload=()=>win.print();
}
