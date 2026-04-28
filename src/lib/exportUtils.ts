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

const $$ = (n: number) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(n);

// ── Colours ───────────────────────────────────────────────────────────────────
const C = {
  NAVY:'1F3864', NAVY2:'2F5496', TEAL:'17375E',
  GREEN:'375623', GREEN2:'70AD47',
  RED:'C00000', AMBER:'ED7D31', GOLD:'C09000',
  WHITE:'FFFFFF', OFF:'F2F2F2', LGRAY:'D9D9D9', DGRAY:'A6A6A6', BLACK:'000000',
  BLUE_LT:'BDD7EE', GRN_LT:'E2EFDA', RED_LT:'FFDCDC', YLW_LT:'FFF2CC', NAVY_LT:'D6DCE4',
};

// ── Style helpers ─────────────────────────────────────────────────────────────
const bdr = { bottom:{ style:'hair', color:{ rgb:C.LGRAY } }, right:{ style:'hair', color:{ rgb:C.LGRAY } } };
const bdrMed = { top:{ style:'thin', color:{ rgb:C.LGRAY } }, bottom:{ style:'double', color:{ rgb:C.LGRAY } } };

const ss = {
  hdr:(bg:string,fg=C.WHITE,sz=11):CS=>({ font:{bold:true,sz,name:'Calibri',color:{rgb:fg}}, fill:{fgColor:{rgb:bg},patternType:'solid'}, alignment:{horizontal:'center',vertical:'center',wrapText:true}, border:{top:{style:'thin',color:{rgb:C.LGRAY}},bottom:{style:'thin',color:{rgb:C.LGRAY}},left:{style:'thin',color:{rgb:C.LGRAY}},right:{style:'thin',color:{rgb:C.LGRAY}}} }),
  cell:(bold=false,fg=C.BLACK,align:'left'|'right'|'center'='left',bg?:string):CS=>({ font:{bold,sz:10,name:'Calibri',color:{rgb:fg}}, ...(bg?{fill:{fgColor:{rgb:bg},patternType:'solid'}}:{}), alignment:{horizontal:align,vertical:'center'}, border:bdr }),
  money:(bold=false,fg=C.BLACK,bg?:string):CS=>({ font:{bold,sz:10,name:'Calibri',color:{rgb:fg}}, ...(bg?{fill:{fgColor:{rgb:bg},patternType:'solid'}}:{}), numFmt:'_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)', alignment:{horizontal:'right',vertical:'center'}, border:bdr }),
  pct:(bold=false,fg=C.BLACK,bg?:string):CS=>({ font:{bold,sz:10,name:'Calibri',color:{rgb:fg}}, ...(bg?{fill:{fgColor:{rgb:bg},patternType:'solid'}}:{}), numFmt:'0.0%', alignment:{horizontal:'right',vertical:'center'}, border:bdr }),
  total:(bg:string,fg=C.WHITE):CS=>({ font:{bold:true,sz:11,name:'Calibri',color:{rgb:fg}}, fill:{fgColor:{rgb:bg},patternType:'solid'}, numFmt:'_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)', alignment:{horizontal:'right',vertical:'center'}, border:bdrMed }),
  totalLbl:(bg:string,fg=C.WHITE):CS=>({ font:{bold:true,sz:11,name:'Calibri',color:{rgb:fg}}, fill:{fgColor:{rgb:bg},patternType:'solid'}, alignment:{horizontal:'left',vertical:'center'}, border:bdrMed }),
  totalPct:(bg:string,fg=C.WHITE):CS=>({ font:{bold:true,sz:11,name:'Calibri',color:{rgb:fg}}, fill:{fgColor:{rgb:bg},patternType:'solid'}, numFmt:'0.0%', alignment:{horizontal:'right',vertical:'center'}, border:bdrMed }),
  secHdr:(bg:string,fg=C.WHITE):CS=>({ font:{bold:true,sz:10,name:'Calibri',color:{rgb:fg}}, fill:{fgColor:{rgb:bg},patternType:'solid'}, alignment:{horizontal:'left',vertical:'center'} }),
  title:(bg:string):CS=>({ font:{bold:true,sz:18,name:'Calibri',color:{rgb:C.WHITE}}, fill:{fgColor:{rgb:bg},patternType:'solid'}, alignment:{horizontal:'left',vertical:'center'} }),
  sub:(bg:string):CS=>({ font:{sz:9,name:'Calibri',color:{rgb:C.NAVY_LT},italic:true}, fill:{fgColor:{rgb:bg},patternType:'solid'}, alignment:{horizontal:'left',vertical:'center'} }),
};

const W = (n:number) => ({wch:n});
const L = (n:number) => n < 26 ? String.fromCharCode(65+n) : String.fromCharCode(64+Math.floor(n/26))+String.fromCharCode(65+(n%26));
const MG = (r:number,c:number,r2:number,c2:number) => ({s:{r,c},e:{r:r2,c:c2}});

// Formula cell helper — value is fallback, formula drives live linking
const fc = (f:string, v:number, style:CS):CS => ({t:'f', f, v, s:style});
const vc = (v:string|number, t:string, style:CS):CS => ({v, t, s:style});

// Raw Ledger column positions (1-indexed for Excel)
const RL = 'Raw Ledger';
const COL = { acct:'A', type:'B', date:'C', txType:'D', num:'E', name:'F', desc:'G', split:'H', amount:'I', balance:'J' };
const MAX_ROW = 1000000; // Excel handles this fine with SUMIF

// QBO account type buckets — used in SUMIF criteria arrays
const INC_TYPES  = '"Income","income","Revenue","revenue","Sales","sales","Other Income","other income","Other Revenue"';
const COGS_TYPES = '"Cost of Goods Sold","cost of goods sold","COGS","cogs","Cost of Sales"';
const EXP_TYPES  = '"Expense","expense","Expenses","expenses","Other Expense","other expense","Other Expenses"';
const ASSET_TYPES= '"Asset","asset","Assets","Bank","bank","Accounts Receivable (A/R)","Other Current Assets","Other Current Asset","Fixed Assets","Fixed Asset","Other Assets","Inventory"';
const LIAB_TYPES = '"Liability","liability","Liabilities","Accounts Payable (A/P)","Credit Card","Other Current Liabilities","Other Current Liability","Long Term Liabilities","Long Term Liability"';
const EQUITY_TYPES='"Equity","equity","Retained Earnings","Opening Balance Equity"';

// SUMIF across multiple type criteria — sum all matching types
const sumTypes = (types: string, amtOrBal: string, extraCriteria = '') => {
  const arr = types.split(',').map(t => t.trim());
  return arr.map(t => `SUMIF('${RL}'!${COL.type}:${COL.type},${t},'${RL}'!${amtOrBal}:${amtOrBal})`).join('+');
};

// SUMPRODUCT for filtered sum (type + month)
const sumMonth = (typeArr: string, monthKey: string, amtCol: string) =>
  `SUMPRODUCT((ISNUMBER(MATCH('${RL}'!${COL.type}$2:${COL.type}$${MAX_ROW},{${typeArr}},0)))*(TEXT('${RL}'!${COL.date}$2:${COL.date}$${MAX_ROW},"YYYY-MM")="${monthKey}")*('${RL}'!${amtCol}$2:${amtCol}$${MAX_ROW}))`;

// SUMPRODUCT for account + month
const sumAcctMonth = (acctName: string, monthKey: string) =>
  `SUMPRODUCT(('${RL}'!${COL.acct}$2:${COL.acct}$${MAX_ROW}="${acctName}")*(TEXT('${RL}'!${COL.date}$2:${COL.date}$${MAX_ROW},"YYYY-MM")="${monthKey}")*('${RL}'!${COL.amount}$2:${COL.amount}$${MAX_ROW}))`;

// Last balance for account (LOOKUP trick)
const lastBalance = (acctRef: string) =>
  `IFERROR(LOOKUP(2,1/('${RL}'!${COL.acct}$2:${COL.acct}$${MAX_ROW}=${acctRef}),'${RL}'!${COL.balance}$2:${COL.balance}$${MAX_ROW}),0)`;

// ── CSV export ────────────────────────────────────────────────────────────────
export function exportCsv(fileName: string, analysis: LedgerAnalysis, pl: ProfitAndLoss, bs: BalanceSheet, cf: CashFlowStatement): void {
  const esc=(v:string|number)=>{const s=String(v);return s.includes(',')||s.includes('"')?`"${s.replace(/"/g,'""')}"`:`${s}`;};
  const sec=(t:string,h:string[],rows:(string|number)[][])=>[t,h.map(esc).join(','),...rows.map(r=>r.map(esc).join(',')),'']. join('\n');
  const sections=[
    sec('P&L',['Item','Amount'],[['Revenue',$$(pl.totalRevenue)],['COGS',$$(pl.totalCogs)],['Gross Profit',$$(pl.grossProfit)],['OpEx',$$(pl.totalExpenses)],['Net Profit',$$(pl.netProfit)]]),
    sec('BALANCE SHEET',['Category','Account','Balance'],[...bs.assets.map(({account,value})=>['Asset',account,$$(value)]),...bs.liabilities.map(({account,value})=>['Liability',account,$$(value)]),...bs.equity.map(({account,value})=>['Equity',account,$$(value)])]),
  ];
  const blob=new Blob([sections.join('\n')],{type:'text/csv;charset=utf-8;'});
  const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=`${base(fileName)}.csv`;a.click();URL.revokeObjectURL(url);
}

// ── Excel export (formula-linked) ─────────────────────────────────────────────
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
  const genDate = new Date().toLocaleDateString('en-US',{year:'numeric',month:'long',day:'numeric'});
  const co = base(fileName);
  const HEADERS = ['Distribution account','Distribution account type','Transaction date','Transaction type','Num','Name','Description','Split','Amount','Balance'];

  // ══ SHEET 1: Raw Ledger ═══════════════════════════════════════════════════
  if (rawRows && rawRows.length > 0) {
    const aoa = [HEADERS, ...rawRows.map(r => HEADERS.map(h => r[h as keyof typeof r] ?? ''))];
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    HEADERS.forEach((_,i) => { ws[`${L(i)}1`].s = ss.hdr(C.NAVY,C.WHITE,10); });

    rawRows.forEach((_, ri) => {
      const rn = ri+2; const bg = ri%2===0 ? C.WHITE : C.OFF;
      HEADERS.forEach((_,ci) => {
        const addr=`${L(ci)}${rn}`; if(!ws[addr]) ws[addr]={t:'s',v:''};
        ws[addr].s = (ci===8||ci===9) ? ss.money(false,C.BLACK,bg) : ss.cell(false,C.BLACK,'left',bg);
      });
    });

    ws['!cols']=[W(28),W(22),W(14),W(14),W(8),W(18),W(32),W(18),W(14),W(14)];
    ws['!freeze']={xSplit:0,ySplit:1};
    ws['!autofilter']={ref:`A1:J1`};
    XLSX.utils.book_append_sheet(wb, ws, RL);
  }

  // ══ SHEET 2: P & L (formula-linked) ══════════════════════════════════════
  {
    const ws: CS = {};
    const put=(r:number,c:number,cell:CS)=>{ ws[`${L(c)}${r}`]=cell; };

    // Title banner
    for(let c=0;c<5;c++) for(let r=1;r<=3;r++) { if(!ws[`${L(c)}${r}`]) ws[`${L(c)}${r}`]={v:'',t:'s',s:{fill:{fgColor:{rgb:C.NAVY},patternType:'solid'}}}; }
    put(1,0,vc(`PROFIT & LOSS — ${co}`,'s',ss.title(C.NAVY)));
    put(2,0,vc(`All figures sourced live from '${RL}' sheet via SUMIF formulas`,'s',ss.sub(C.NAVY)));
    put(3,0,vc(`Generated: ${genDate}`,'s',ss.sub(C.NAVY)));

    // Headers row 5
    put(5,0,vc('LINE ITEM','s',ss.hdr(C.TEAL)));
    put(5,1,vc('Amount','s',ss.hdr(C.TEAL)));
    put(5,2,vc('% of Revenue','s',ss.hdr(C.TEAL)));

    // Revenue
    put(6,0,vc('REVENUE','s',ss.secHdr(C.NAVY2)));
    put(6,1,vc('','s',{fill:{fgColor:{rgb:C.NAVY2},patternType:'solid'}}));
    put(6,2,vc('','s',{fill:{fgColor:{rgb:C.NAVY2},patternType:'solid'}}));

    put(7,0,vc('  Total Revenue','s',ss.cell(false,C.BLACK,'left',C.GRN_LT)));
    put(7,1,fc(`${sumTypes(INC_TYPES,COL.amount)}`, pl.totalRevenue, ss.money(true,C.GREEN2,C.GRN_LT)));
    put(7,2,fc(`IF(B7=0,0,B7/B7)`, 1, ss.pct(false,C.BLACK,C.GRN_LT)));

    // COGS
    put(9,0,vc('COST OF GOODS SOLD','s',ss.secHdr(C.TEAL)));
    put(9,1,vc('','s',{fill:{fgColor:{rgb:C.TEAL},patternType:'solid'}}));
    put(9,2,vc('','s',{fill:{fgColor:{rgb:C.TEAL},patternType:'solid'}}));

    put(10,0,vc('  Total COGS','s',ss.cell()));
    put(10,1,fc(`${sumTypes(COGS_TYPES,COL.amount)}`, pl.totalCogs, ss.money()));
    put(10,2,fc(`IF(B7=0,0,B10/B7)`, pl.totalRevenue ? pl.totalCogs/pl.totalRevenue : 0, ss.pct()));

    put(11,0,vc('GROSS PROFIT','s',ss.totalLbl(C.GREEN)));
    put(11,1,fc(`B7-B10`, pl.grossProfit, ss.total(C.GREEN)));
    put(11,2,fc(`IF(B7=0,0,B11/B7)`, pl.grossMargin, ss.totalPct(C.GREEN)));

    // Expenses
    put(13,0,vc('OPERATING EXPENSES','s',ss.secHdr(C.TEAL)));
    put(13,1,vc('','s',{fill:{fgColor:{rgb:C.TEAL},patternType:'solid'}}));
    put(13,2,vc('','s',{fill:{fgColor:{rgb:C.TEAL},patternType:'solid'}}));

    put(14,0,vc('  Total Operating Expenses','s',ss.cell()));
    put(14,1,fc(`${sumTypes(EXP_TYPES,COL.amount)}`, pl.totalExpenses, ss.money()));
    put(14,2,fc(`IF(B7=0,0,B14/B7)`, pl.totalRevenue ? pl.totalExpenses/pl.totalRevenue : 0, ss.pct()));

    const npColor = pl.netProfit >= 0 ? C.GREEN : C.RED;
    put(15,0,vc('NET PROFIT / (LOSS)','s',ss.totalLbl(npColor)));
    put(15,1,fc(`B11-B14`, pl.netProfit, ss.total(npColor)));
    put(15,2,fc(`IF(B7=0,0,B15/B7)`, pl.netMargin, ss.totalPct(npColor)));

    // Monthly breakdown
    put(17,0,vc('MONTHLY BREAKDOWN','s',ss.secHdr(C.NAVY)));
    ['','','','',''].forEach((_,i)=>{ ws[`${L(i)}17`]={v:'',t:'s',s:{fill:{fgColor:{rgb:C.NAVY},patternType:'solid'}}}; });
    put(17,0,vc('MONTHLY BREAKDOWN','s',ss.secHdr(C.NAVY)));

    ['Month','Revenue','COGS','Op Expenses','Net Profit','Margin %'].forEach((h,i) => put(18,i,vc(h,'s',ss.hdr(C.NAVY2,C.WHITE,10))));

    const months = Object.entries(pl.monthlyBreakdown).sort();
    months.forEach(([mk, {revenue, cogs, expenses}], i) => {
      const r = 19+i; const bg = i%2===0?C.WHITE:C.OFF;
      const net = revenue-(cogs??0)-expenses;
      put(r,0,vc(mk,'s',ss.cell(false,C.BLACK,'left',bg)));
      put(r,1,fc(sumMonth(INC_TYPES,mk,COL.amount), revenue, ss.money(false,C.BLACK,bg)));
      put(r,2,fc(sumMonth(COGS_TYPES,mk,COL.amount), cogs??0, ss.money(false,C.BLACK,bg)));
      put(r,3,fc(sumMonth(EXP_TYPES,mk,COL.amount), expenses, ss.money(false,C.BLACK,bg)));
      put(r,4,fc(`${L(1)}${r}-${L(2)}${r}-${L(3)}${r}`, net, ss.money(false,net>=0?C.GREEN:C.RED,bg)));
      put(r,5,fc(`IF(${L(1)}${r}=0,0,${L(4)}${r}/${L(1)}${r})`, revenue?net/revenue:0, ss.pct(false,net>=0?C.GREEN:C.RED,bg)));
    });

    const lr = 19+months.length;
    put(lr,0,vc('TOTAL','s',ss.totalLbl(C.NAVY)));
    put(lr,1,fc(`SUM(B19:B${lr-1})`, pl.totalRevenue, ss.total(C.NAVY)));
    put(lr,2,fc(`SUM(C19:C${lr-1})`, pl.totalCogs, ss.total(C.NAVY)));
    put(lr,3,fc(`SUM(D19:D${lr-1})`, pl.totalExpenses, ss.total(C.NAVY)));
    put(lr,4,fc(`SUM(E19:E${lr-1})`, pl.netProfit, ss.total(npColor)));
    put(lr,5,fc(`IF(B${lr}=0,0,E${lr}/B${lr})`, pl.netMargin, ss.totalPct(C.NAVY)));

    ws['!ref']=`A1:F${lr}`;
    ws['!cols']=[W(34),W(18),W(16),W(18),W(18),W(12)];
    ws['!merges']=[MG(0,0,0,5),MG(1,0,1,5),MG(2,0,2,5),MG(5,0,5,0),MG(8,0,8,2),MG(12,0,12,2),MG(16,0,16,5)];
    ws['!freeze']={xSplit:0,ySplit:5};
    XLSX.utils.book_append_sheet(wb, ws, 'P & L');
  }

  // ══ SHEET 3: Balance Sheet (formula-linked) ════════════════════════════════
  {
    const ws: CS = {};
    const put=(r:number,c:number,cell:CS)=>{ ws[`${L(c)}${r}`]=cell; };

    for(let c=0;c<4;c++) for(let r=1;r<=3;r++) { ws[`${L(c)}${r}`]={v:'',t:'s',s:{fill:{fgColor:{rgb:C.NAVY},patternType:'solid'}}}; }
    put(1,0,vc(`BALANCE SHEET — ${co}`,'s',ss.title(C.NAVY)));
    put(2,0,vc(`Last balance per account via LOOKUP formula from '${RL}' sheet`,'s',ss.sub(C.NAVY)));
    put(3,0,vc(`Generated: ${genDate}`,'s',ss.sub(C.NAVY)));

    let r = 5;

    const writeSection = (
      title: string, entries: BalanceSheet['assets'],
      sumFormula: string, fallbackTotal: number, bg: string
    ) => {
      put(r,0,vc(title,'s',ss.secHdr(bg)));
      put(r,1,vc('Balance','s',ss.hdr(bg)));
      put(r,2,vc('% of Total','s',ss.hdr(bg)));
      r++;

      const startR = r;
      entries.forEach((e,i) => {
        const rowBg = i%2===0?C.WHITE:C.OFF;
        put(r,0,vc(`  ${e.account}`,'s',ss.cell(false,C.BLACK,'left',rowBg)));
        put(r,1,fc(lastBalance(`A${r}`), e.value, ss.money(false,C.BLACK,rowBg)));
        put(r,2,fc(`IF(SUM(B${startR}:B${startR+entries.length})=0,0,B${r}/SUM(B${startR}:B${startR+entries.length-1}))`, fallbackTotal?e.value/fallbackTotal:0, ss.pct(false,C.BLACK,rowBg)));
        r++;
      });

      put(r,0,vc('Total','s',ss.totalLbl(bg)));
      put(r,1,fc(`SUM(B${startR}:B${r-1})`, fallbackTotal, ss.total(bg)));
      put(r,2,fc('1',1,ss.totalPct(bg)));
      r+=2;
    };

    writeSection('ASSETS',  bs.assets,  sumTypes(ASSET_TYPES,COL.balance),  bs.totals.assetsTotal,      C.NAVY2);
    writeSection('LIABILITIES', bs.liabilities, sumTypes(LIAB_TYPES,COL.balance), bs.totals.liabilitiesTotal, C.TEAL);
    writeSection('EQUITY',  bs.equity,  sumTypes(EQUITY_TYPES,COL.balance), bs.totals.equityTotal,      C.GREEN);

    // Reconciliation — formulas reference Total rows above
    put(r,0,vc('BALANCE SHEET CHECK','s',ss.secHdr(C.GOLD)));
    put(r,1,vc('Value','s',ss.hdr(C.GOLD)));
    put(r,2,vc('Expected','s',ss.hdr(C.GOLD)));
    r++;

    // Find total rows by scanning for them — use fallback pre-computed values with formula refs
    // Assets total is always at a fixed offset, so we reference by formula
    const assetTotalR  = 6 + bs.assets.length;
    const liabTotalR   = assetTotalR + 3 + bs.liabilities.length;
    const equityTotalR = liabTotalR  + 3 + bs.equity.length;

    put(r,0,vc('Assets','s',ss.cell(true)));
    put(r,1,fc(`B${assetTotalR}`, bs.totals.assetsTotal, ss.money(true)));
    r++;
    put(r,0,vc('Liabilities + Equity','s',ss.cell(true)));
    put(r,1,fc(`B${liabTotalR}+B${equityTotalR}`, bs.totals.liabilitiesTotal+bs.totals.equityTotal, ss.money(true)));
    r++;
    put(r,0,vc('Variance (A − L − E)','s',ss.cell(true,bs.isBalanced?C.GREEN:C.RED)));
    put(r,1,fc(`B${assetTotalR}-(B${liabTotalR}+B${equityTotalR})`, bs.variance, ss.money(true,bs.isBalanced?C.GREEN:C.RED,bs.isBalanced?C.GRN_LT:C.RED_LT)));
    r++;
    put(r,0,vc('Status','s',ss.cell(true)));
    put(r,1,fc(`IF(ABS(B${r-1})<=1,"BALANCED ✓","NOT BALANCED ✗")`, bs.isBalanced?1:0, {
      ...ss.cell(true,bs.isBalanced?C.GREEN:C.RED,'left',bs.isBalanced?C.GRN_LT:C.RED_LT),
      numFmt:'@',
    }));

    ws['!ref']=`A1:C${r}`;
    ws['!cols']=[W(36),W(18),W(12)];
    ws['!merges']=[MG(0,0,0,3),MG(1,0,1,3),MG(2,0,2,3)];
    ws['!freeze']={xSplit:0,ySplit:4};
    XLSX.utils.book_append_sheet(wb, ws, 'Balance Sheet');
  }

  // ══ SHEET 4: Month-over-Month (formula-linked) ════════════════════════════
  if (mom && mom.months.length > 0) {
    const ws: CS = {};
    const months = mom.months;
    const nCols = months.length + 2;
    const put=(r:number,c:number,cell:CS)=>{ ws[`${L(c)}${r}`]=cell; };

    for(let c=0;c<nCols;c++) for(let r=1;r<=3;r++) { ws[`${L(c)}${r}`]={v:'',t:'s',s:{fill:{fgColor:{rgb:C.NAVY},patternType:'solid'}}}; }
    put(1,0,vc(`MONTH-OVER-MONTH P&L — ${co}`,'s',ss.title(C.NAVY)));
    put(2,0,vc(`All figures via SUMPRODUCT formulas from '${RL}' sheet · ${months.length} months`,'s',ss.sub(C.NAVY)));
    put(3,0,vc(`Generated: ${genDate}`,'s',ss.sub(C.NAVY)));

    const HDR = 5;
    put(HDR,0,vc('Account','s',ss.hdr(C.NAVY,C.WHITE,10)));
    months.forEach((m,i) => put(HDR,i+1,vc(monthLabel(m),'s',ss.hdr(C.NAVY,C.WHITE,10))));
    put(HDR,months.length+1,vc('Total','s',ss.hdr(C.NAVY,C.WHITE,10)));

    // MoM % change row
    put(HDR+1,0,vc('MoM Revenue Δ%','s',ss.hdr(C.TEAL,C.WHITE,9)));
    months.forEach((m,i) => {
      if(i===0){put(HDR+1,1,vc('Baseline','s',ss.hdr(C.TEAL,C.WHITE,9)));return;}
      const cur=`${L(i+1)}${HDR+2+mom.incomeCategories.length+1}`;// will ref total revenue row
      const prev=`${L(i)}${HDR+2+mom.incomeCategories.length+1}`;
      // Use approximate % for now, proper ref after we know row numbers
      const curV=mom.monthlyRevenue[m]??0; const prevV=mom.monthlyRevenue[months[i-1]]??0;
      put(HDR+1,i+1,fc(`IF(${prev}=0,0,(${cur}-${prev})/${prev})`, prevV?((curV-prevV)/prevV):0, ss.pct(false,curV>=prevV?C.GREEN:C.RED,C.YLW_LT)));
    });
    put(HDR+1,months.length+1,vc('','s',ss.hdr(C.TEAL,C.WHITE,9)));

    let r = HDR+2;

    // Income
    put(r,0,vc('▸  INCOME','s',ss.secHdr(C.GREEN2)));
    for(let c=1;c<nCols;c++) ws[`${L(c)}${r}`]={v:'',t:'s',s:{fill:{fgColor:{rgb:C.GREEN2},patternType:'solid'}}};
    r++;

    const incStartR = r;
    mom.incomeCategories.forEach((cat,ri) => {
      const bg=ri%2===0?C.WHITE:C.OFF;
      put(r,0,vc(`  ${cat.name}`,'s',ss.cell(false,C.BLACK,'left',bg)));
      months.forEach((m,i) => {
        const v=cat.months[m]??0;
        put(r,i+1,fc(sumAcctMonth(cat.name,m), v, ss.money(false,v>0?C.BLACK:C.LGRAY,bg)));
      });
      put(r,months.length+1,fc(`SUM(${L(1)}${r}:${L(months.length)}${r})`, cat.total, ss.money(true,C.BLACK,bg)));
      r++;
    });

    const revTotalR = r;
    put(r,0,vc('Total Revenue','s',ss.totalLbl(C.GREEN)));
    months.forEach((_,i) => put(r,i+1,fc(`SUM(${L(i+1)}${incStartR}:${L(i+1)}${r-1})`, mom.monthlyRevenue[months[i]]??0, ss.total(C.GREEN))));
    put(r,months.length+1,fc(`SUM(${L(1)}${r}:${L(months.length)}${r})`, mom.totalRevenue, ss.total(C.GREEN)));
    r+=2;

    // Expenses
    put(r,0,vc('▸  EXPENSES','s',ss.secHdr(C.RED)));
    for(let c=1;c<nCols;c++) ws[`${L(c)}${r}`]={v:'',t:'s',s:{fill:{fgColor:{rgb:C.RED},patternType:'solid'}}};
    r++;

    const expStartR = r;
    mom.expenseCategories.forEach((cat,ri) => {
      const bg=ri%2===0?C.WHITE:C.OFF;
      put(r,0,vc(`  ${cat.name}`,'s',ss.cell(false,C.BLACK,'left',bg)));
      months.forEach((m,i) => {
        const v=cat.months[m]??0;
        put(r,i+1,fc(sumAcctMonth(cat.name,m), v, ss.money(false,v>0?C.BLACK:C.LGRAY,bg)));
      });
      put(r,months.length+1,fc(`SUM(${L(1)}${r}:${L(months.length)}${r})`, cat.total, ss.money(true,C.BLACK,bg)));
      r++;
    });

    const expTotalR = r;
    put(r,0,vc('Total Expenses','s',ss.totalLbl(C.RED)));
    months.forEach((_,i) => put(r,i+1,fc(`SUM(${L(i+1)}${expStartR}:${L(i+1)}${r-1})`, mom.monthlyExpenses[months[i]]??0, ss.total(C.RED))));
    put(r,months.length+1,fc(`SUM(${L(1)}${r}:${L(months.length)}${r})`, mom.totalExpenses, ss.total(C.RED)));
    r+=2;

    // Net Profit
    const npColor = mom.totalNetProfit>=0?C.GREEN:C.RED;
    put(r,0,vc('NET PROFIT','s',ss.totalLbl(npColor)));
    months.forEach((_,i) => {
      const v=mom.monthlyNetProfit[months[i]]??0;
      put(r,i+1,fc(`${L(i+1)}${revTotalR}-${L(i+1)}${expTotalR}`, v, ss.total(v>=0?C.GREEN:C.RED)));
    });
    put(r,months.length+1,fc(`${L(months.length+1)}${revTotalR}-${L(months.length+1)}${expTotalR}`, mom.totalNetProfit, ss.total(npColor)));

    ws['!ref']=`A1:${L(nCols-1)}${r}`;
    ws['!cols']=[W(32),...months.map(()=>W(13)),W(15)];
    ws['!merges']=[MG(0,0,0,nCols-1),MG(1,0,1,nCols-1),MG(2,0,2,nCols-1)];
    ws['!freeze']={xSplit:1,ySplit:HDR};
    XLSX.utils.book_append_sheet(wb, ws, 'Month-over-Month');
  }

  // ══ SHEET 5: Cash Flow ════════════════════════════════════════════════════
  {
    const ws: CS = {};
    const put=(r:number,c:number,cell:CS)=>{ ws[`${L(c)}${r}`]=cell; };

    for(let c=0;c<2;c++) for(let r=1;r<=3;r++) { ws[`${L(c)}${r}`]={v:'',t:'s',s:{fill:{fgColor:{rgb:C.NAVY},patternType:'solid'}}}; }
    put(1,0,vc(`CASH FLOW — ${co}`,'s',ss.title(C.NAVY)));
    put(2,0,vc(`Net Profit sourced from P&L sheet (cell B15)`,'s',ss.sub(C.NAVY)));
    put(3,0,vc(`Generated: ${genDate}`,'s',ss.sub(C.NAVY)));

    put(5,0,vc('OPERATING ACTIVITIES','s',ss.secHdr(C.NAVY2)));
    put(5,1,vc('Amount','s',ss.hdr(C.NAVY2)));

    // Net Profit references P&L sheet
    put(6,0,vc('  Net Profit (from P&L)','s',ss.cell(false,C.BLACK,'left',C.BLUE_LT)));
    put(6,1,fc(`'P & L'!B15`, pl.netProfit, ss.money(true,C.NAVY2,C.BLUE_LT)));

    put(7,0,vc('  Adjustments:','s',ss.cell(false,C.DGRAY)));

    let r = 7;
    cf.adjustments.forEach((adj,i) => {
      r++;
      const bg=i%2===0?C.WHITE:C.OFF;
      put(r,0,vc(`    ${adj.account}`,'s',ss.cell(false,C.BLACK,'left',bg)));
      put(r,1,vc(adj.impact,'n',ss.money(false,adj.impact>=0?C.GREEN:C.RED,bg)));
    });

    r++;
    const ocfColor = cf.operatingCashFlow>=0?C.GREEN:C.RED;
    put(r,0,vc('NET OPERATING CASH FLOW','s',ss.totalLbl(ocfColor)));
    put(r,1,fc(`B6+SUM(B8:B${r-1})`, cf.operatingCashFlow, ss.total(ocfColor)));

    r+=2;
    put(r,0,vc(`📌 Net Profit linked from 'P & L' tab · Transactions in '${RL}' tab`,'s',{font:{sz:9,name:'Calibri',color:{rgb:C.DGRAY},italic:true},alignment:{horizontal:'left'}}));

    ws['!ref']=`A1:B${r}`;
    ws['!cols']=[W(42),W(20)];
    ws['!merges']=[MG(0,0,0,1),MG(1,0,1,1),MG(2,0,2,1)];
    ws['!freeze']={xSplit:0,ySplit:4};
    XLSX.utils.book_append_sheet(wb, ws, 'Cash Flow');
  }

  // ══ SHEET 6: Flags ════════════════════════════════════════════════════════
  {
    const ws: CS = {};
    let r = 1;
    const put=(row:number,c:number,cell:CS)=>{ ws[`${L(c)}${row}`]=cell; };

    for(let c=0;c<4;c++) for(let rr=1;rr<=3;rr++) { ws[`${L(c)}${rr}`]={v:'',t:'s',s:{fill:{fgColor:{rgb:C.RED},patternType:'solid'}}}; }
    put(1,0,vc(`FLAGS & AUDIT — ${co}`,'s',ss.title(C.RED)));
    put(2,0,vc(`${analysis.inconsistentVendors.length} inconsistent · ${analysis.duplicates.length} duplicates`,'s',ss.sub(C.RED)));

    r=5;
    ['Vendor','Reason','Accounts'].forEach((h,i) => put(r,i,vc(h,'s',ss.hdr(i===0?C.GOLD:C.NAVY2))));
    analysis.inconsistentVendors.forEach((v,i) => {
      r++; const bg=i%2===0?C.YLW_LT:C.WHITE;
      put(r,0,vc(v.vendor,'s',ss.cell(true,C.BLACK,'left',bg)));
      put(r,1,vc(v.reason,'s',ss.cell(false,C.BLACK,'left',bg)));
      put(r,2,vc(v.accounts.join(', '),'s',ss.cell(false,C.BLACK,'left',bg)));
    });
    if(!analysis.inconsistentVendors.length){r++;put(r,0,vc('None found ✓','s',ss.cell(false,C.GREEN)));}

    r+=2;
    ['Vendor','Amount','Date','Count'].forEach((h,i) => put(r,i,vc(h,'s',ss.hdr(i===0?C.RED:C.NAVY2))));
    analysis.duplicates.forEach((d,i) => {
      r++; const bg=i%2===0?C.RED_LT:C.WHITE;
      put(r,0,vc(d.name,'s',ss.cell(true,C.BLACK,'left',bg)));
      put(r,1,vc(d.amount,'s',ss.cell(false,C.BLACK,'right',bg)));
      put(r,2,vc(d.transactionDate,'s',ss.cell(false,C.BLACK,'center',bg)));
      put(r,3,vc(d.occurrences,'n',ss.cell(true,C.RED,'center',bg)));
    });
    if(!analysis.duplicates.length){r++;put(r,0,vc('None found ✓','s',ss.cell(false,C.GREEN)));}

    ws['!ref']=`A1:D${r}`;
    ws['!cols']=[W(28),W(36),W(44),W(10)];
    ws['!merges']=[MG(0,0,0,3),MG(1,0,1,3)];
    XLSX.utils.book_append_sheet(wb, ws, 'Flags');
  }

  XLSX.writeFile(wb, `${base(fileName)}_Financial_Analysis.xlsx`);
}

// ── PDF ───────────────────────────────────────────────────────────────────────
export function exportPdf(fileName:string,analysis:LedgerAnalysis,pl:ProfitAndLoss,bs:BalanceSheet,cf:CashFlowStatement):void{
  const tbl=(h:string[],rows:(string|number)[][])=>`<table><thead><tr>${h.map(x=>`<th>${x}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map(c=>`<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>${base(fileName)}</title><style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;font-size:12px;color:#111;padding:32px}h1{font-size:20px;margin-bottom:4px}.meta{color:#666;font-size:11px;margin-bottom:24px}.section{margin-bottom:24px}h2{font-size:13px;font-weight:700;text-transform:uppercase;color:#1F3864;border-bottom:2px solid #1F3864;padding-bottom:4px;margin-bottom:8px}table{width:100%;border-collapse:collapse;font-size:11px}th{background:#1F3864;color:#fff;text-align:left;padding:5px 8px}td{padding:4px 8px;border-bottom:1px solid #e5e7eb}tr:nth-child(even) td{background:#f1f5f9}@media print{body{padding:16px}}</style></head><body>
<h1>${base(fileName)} — Financial Analysis</h1><p class="meta">Generated: ${new Date().toLocaleString()} · ${analysis.totalTransactions} transactions</p>
<div class="section"><h2>P&L</h2>${tbl(['Item','Amount','% Rev'],[['Revenue',$$(pl.totalRevenue),'100%'],['COGS',$$(pl.totalCogs),`${(pl.totalRevenue?pl.totalCogs/pl.totalRevenue*100:0).toFixed(1)}%`],['Gross Profit',$$(pl.grossProfit),`${(pl.grossMargin*100).toFixed(1)}%`],['OpEx',$$(pl.totalExpenses),`${(pl.totalRevenue?pl.totalExpenses/pl.totalRevenue*100:0).toFixed(1)}%`],['Net Profit',$$(pl.netProfit),`${(pl.netMargin*100).toFixed(1)}%`]])}</div>
<div class="section"><h2>Balance Sheet ${bs.isBalanced?'✓ Balanced':'✗ Not Balanced'}</h2>${tbl(['Category','Account','Balance'],[...bs.assets.map(({account,value})=>['Asset',account,$$(value)]),...bs.liabilities.map(({account,value})=>['Liability',account,$$(value)]),...bs.equity.map(({account,value})=>['Equity',account,$$(value)])])}</div>
<div class="section"><h2>Cash Flow</h2>${tbl(['Item','Amount'],[['Net Profit',$$(cf.netProfit)],['Operating CF',$$(cf.operatingCashFlow)]])}</div>
</body></html>`;
  const win=window.open('','_blank');if(!win)return;win.document.write(html);win.document.close();win.onload=()=>win.print();
}

function base(f:string){return f.replace(/\.[^.]+$/,'')||'ledger';}
