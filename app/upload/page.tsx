'use client';

import { useCallback, useMemo, useState } from 'react';
import { analyzeLedger, type LedgerAnalysis } from '../../src/lib/analyzeLedger';
import { generateBalanceSheet, type BalanceSheet } from '../../src/lib/generateBalanceSheet';
import { generateCashFlow, type CashFlowStatement } from '../../src/lib/generateCashFlow';
import { generatePL, type ProfitAndLoss } from '../../src/lib/generatePL';
import { generateMoMPL, monthLabel, momChange, type MoMPL } from '../../src/lib/generateMoMPL';
import { generateInsights, type FinancialInsights } from '../../src/lib/generateInsights';
import { exportCsv, exportExcel, exportPdf } from '../../src/lib/exportUtils';
import { FileDropzone, type UploadTheme } from './components/FileDropzone';
import { PreviewTable } from './components/PreviewTable';
import { ValidationPanel } from './components/ValidationPanel';
import { QBOGuide } from './components/QBOGuide';
import { ColumnMapper } from './components/ColumnMapper';
import {
  isCsvFile, isExcelFile, parseCsvFile, parseXlsxFile,
  requiredHeaders, CORE_MANDATORY, type LedgerRow, type RowIssue, type HeaderKey,
} from './upload-utils';

type ThemeClasses = {
  page: string; navbar: string; panel: string; card: string; shell: string;
  listRow: string; heading: string; body: string; muted: string; stat: string;
  label: string; successPill: string; dangerPill: string; successCard: string;
  skyPill: string; skyCard: string; warningRow: string; divider: string;
  tabActive: string; tabInactive: string; settingsControl: string;
  settingsOptionActive: string; settingsOptionInactive: string;
  tableHead: string; tableRow: string; tableAlt: string;
};

const themes: Record<UploadTheme, ThemeClasses> = {
  dark: {
    page: 'bg-slate-950 text-slate-100',
    shell: 'border-slate-700 bg-slate-900/80 shadow-2xl',
    navbar: 'border-slate-700 bg-slate-900/95 backdrop-blur',
    panel: 'border-slate-700 bg-slate-900',
    card: 'bg-slate-950/60',
    listRow: 'bg-slate-950/50 text-slate-300',
    heading: 'text-white', body: 'text-slate-300', muted: 'text-slate-400',
    stat: 'text-white', label: 'text-slate-400',
    successPill: 'bg-emerald-950/70 text-emerald-200',
    dangerPill: 'bg-rose-950/70 text-rose-200',
    successCard: 'border-emerald-500/40 bg-emerald-950/50 text-emerald-100',
    skyPill: 'bg-sky-950/70 text-sky-200',
    skyCard: 'border-sky-500/40 bg-sky-950/50 text-sky-100',
    warningRow: 'bg-amber-950/50 text-amber-200',
    divider: 'border-slate-700',
    tabActive: 'border-cyan-400 text-cyan-300',
    tabInactive: 'border-transparent text-slate-400 hover:text-slate-200 hover:border-slate-500',
    settingsControl: 'bg-slate-800 text-slate-200',
    settingsOptionActive: 'bg-cyan-400 text-slate-950',
    settingsOptionInactive: 'text-slate-300 hover:bg-slate-700',
    tableHead: 'bg-slate-800 text-slate-300',
    tableRow: 'bg-slate-900 text-slate-300',
    tableAlt: 'bg-slate-950/50 text-slate-300',
  },
  light: {
    page: 'bg-slate-100 text-slate-900',
    shell: 'border-slate-200 bg-white shadow-xl',
    navbar: 'border-slate-200 bg-white/95 backdrop-blur',
    panel: 'border-slate-200 bg-slate-50',
    card: 'bg-white',
    listRow: 'bg-slate-50 text-slate-700',
    heading: 'text-slate-950', body: 'text-slate-600', muted: 'text-slate-500',
    stat: 'text-slate-950', label: 'text-slate-500',
    successPill: 'bg-emerald-100 text-emerald-800',
    dangerPill: 'bg-rose-100 text-rose-800',
    successCard: 'border-emerald-200 bg-emerald-50 text-emerald-900',
    skyPill: 'bg-sky-100 text-sky-800',
    skyCard: 'border-sky-200 bg-sky-50 text-sky-900',
    warningRow: 'bg-amber-50 text-amber-900',
    divider: 'border-slate-200',
    tabActive: 'border-slate-900 text-slate-900',
    tabInactive: 'border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300',
    settingsControl: 'bg-slate-100 text-slate-700',
    settingsOptionActive: 'bg-slate-900 text-white',
    settingsOptionInactive: 'text-slate-600 hover:bg-slate-200',
    tableHead: 'bg-slate-100 text-slate-600',
    tableRow: 'bg-white text-slate-700',
    tableAlt: 'bg-slate-50 text-slate-700',
  },
};

type Tab = 'overview' | 'insights' | 'pl' | 'bs' | 'cashflow' | 'mom' | 'preview';

export default function UploadPage() {
  const [theme, setTheme] = useState<UploadTheme>('dark');
  const [view, setView] = useState<'upload' | 'dashboard'>('upload');
  const [activeTab, setActiveTab] = useState<Tab>('overview');
  const [showMoMPct, setShowMoMPct] = useState(false);

  const [previewRows, setPreviewRows] = useState<LedgerRow[]>([]);
  const [analysis, setAnalysis] = useState<LedgerAnalysis | null>(null);
  const [profitAndLoss, setProfitAndLoss] = useState<ProfitAndLoss | null>(null);
  const [balanceSheet, setBalanceSheet] = useState<BalanceSheet | null>(null);
  const [cashFlow, setCashFlow] = useState<CashFlowStatement | null>(null);
  const [momPL, setMomPL] = useState<MoMPL | null>(null);
  const [insights, setInsights] = useState<FinancialInsights | null>(null);
  const [pendingFile, setPendingFile] = useState<File | null>(null);
  const [fileHeaders, setFileHeaders] = useState<string[]>([]);
  const [showMapper, setShowMapper] = useState(false);
  const [headerErrors, setHeaderErrors] = useState<string[]>([]);
  const [rowIssues, setRowIssues] = useState<RowIssue[]>([]);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [fileName, setFileName] = useState('');
  const [isDragging, setIsDragging] = useState(false);

  const ui = themes[theme];
  const fmt = useMemo(() => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }), []);

  const clearState = () => {
    setUploadError(null); setHeaderErrors([]); setRowIssues([]);
    setPreviewRows([]); setAnalysis(null); setProfitAndLoss(null);
    setBalanceSheet(null); setCashFlow(null); setMomPL(null); setInsights(null);
    setPendingFile(null); setFileHeaders([]); setShowMapper(false);
  };

  const processRows = useCallback((rows: LedgerRow[], rowIssues: RowIssue[]) => {
    const plResult = generatePL(rows);
    setAnalysis(analyzeLedger(rows));
    setProfitAndLoss(plResult);
    setBalanceSheet(generateBalanceSheet(rows));
    setCashFlow(generateCashFlow(rows, plResult));
    setMomPL(generateMoMPL(rows));
    setInsights(generateInsights(rows));
    setPreviewRows(rows.slice(0, 100));
    setRowIssues(rowIssues);
    setView('dashboard');
    setActiveTab('overview');
  }, []);

  const detectFileHeaders = async (file: File): Promise<string[]> => {
    return new Promise((resolve) => {
      if (isCsvFile(file)) {
        const reader = new FileReader();
        reader.onload = (e) => {
          const firstLine = (e.target?.result as string).split('\n')[0] ?? '';
          resolve(firstLine.split(',').map(h => h.trim().replace(/^"|"$/g, '')));
        };
        reader.readAsText(file);
      } else {
        import('xlsx').then(XLSX => {
          const reader = new FileReader();
          reader.onload = (e) => {
            const wb = XLSX.read(e.target?.result, { type: 'array' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '', raw: false }) as unknown[][];
            // Find first row with 3+ non-empty cells
            for (const row of rows.slice(0, 20)) {
              const vals = (row as string[]).map(v => String(v ?? '').trim()).filter(Boolean);
              if (vals.length >= 3) { resolve(vals); return; }
            }
            resolve([]);
          };
          reader.readAsArrayBuffer(file);
        });
      }
    });
  };

  const handleMappingConfirmed = useCallback(async (mapping: Record<HeaderKey, string>) => {
    if (!pendingFile) return;
    setShowMapper(false);
    setHeaderErrors([]);
    try {
      const result = isCsvFile(pendingFile)
        ? await parseCsvFile(pendingFile, mapping)
        : await parseXlsxFile(pendingFile, mapping);
      if (result.headerErrors.length > 0) { setHeaderErrors(result.headerErrors); return; }
      processRows(result.rows, result.rowIssues);
    } catch {
      setUploadError('Unable to parse file with the provided mapping.');
    }
  }, [pendingFile, processRows]);

  const handleParse = useCallback(async (file: File) => {
    clearState();
    setFileName(file.name);
    if (!isCsvFile(file) && !isExcelFile(file)) {
      setUploadError('Only CSV and XLSX files are accepted.');
      return;
    }
    try {
      const result = isCsvFile(file) ? await parseCsvFile(file) : await parseXlsxFile(file);
      setHeaderErrors(result.headerErrors);
      setRowIssues(result.rowIssues);

      // If mandatory headers missing — check if file has OTHER headers we can remap
      if (result.headerErrors.length > 0) {
        const detectedHeaders = await detectFileHeaders(file);
        if (detectedHeaders.length > 0) {
          setFileHeaders(detectedHeaders);
          setPendingFile(file);
          setShowMapper(true);
          return;
        }
        setPreviewRows([]);
        return;
      }
      const plResult = generatePL(result.rows);
      setAnalysis(analyzeLedger(result.rows));
      setProfitAndLoss(plResult);
      setBalanceSheet(generateBalanceSheet(result.rows));
      setCashFlow(generateCashFlow(result.rows, plResult));
      setMomPL(generateMoMPL(result.rows));
      setInsights(generateInsights(result.rows));
      setPreviewRows(result.rows.slice(0, 100));
      setView('dashboard');
      setActiveTab('overview');
    } catch (e) {
      setUploadError('Unable to parse file. Please verify the file format.');
    }
  }, []);

  const handleFileChange = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) await handleParse(file);
  }, [handleParse]);

  const handleDrop = useCallback(async (e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault(); setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) await handleParse(file);
  }, [handleParse]);

  const handleDragOver = useCallback((e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault(); setIsDragging(true);
  }, []);

  const handleUploadNew = () => { clearState(); setFileName(''); setView('upload'); };

  const parseCurrencyAmountClient = (v: string) => { const s = v.trim(); if (!s) return 0; const neg = s.startsWith('(') || s.startsWith('-'); return (neg ? -1 : 1) * Math.abs(Number(s.replace(/[,$() -]/g, '')) || 0); };

  const ThemeToggle = () => (
    <div className={`flex rounded-full p-1 text-xs font-semibold ${ui.settingsControl}`}>
      {(['dark', 'light'] as const).map((t) => (
        <button key={t} type="button" onClick={() => setTheme(t)}
          className={`rounded-full px-3 py-1.5 capitalize transition-colors ${theme === t ? ui.settingsOptionActive : ui.settingsOptionInactive}`}>
          {t}
        </button>
      ))}
    </div>
  );

  const metricCard = (label: string, value: string | number) => (
    <div className={`rounded-2xl p-4 shadow-sm ${ui.card}`}>
      <p className={`text-sm ${ui.label}`}>{label}</p>
      <p className={`mt-2 text-2xl font-semibold ${ui.stat}`}>{value}</p>
    </div>
  );

  // ─── Upload View ────────────────────────────────────────────────────────────

  if (view === 'upload') {
    return (
      <div className={`min-h-screen px-6 py-8 ${ui.page}`}>
        <div className={`w-full rounded-3xl border p-8 ${ui.shell}`}>
          <div className="space-y-6">
            <div className="flex items-center justify-between">
              <div>
                <h1 className={`text-2xl font-bold ${ui.heading}`}>Ledger Insight AI</h1>
                <p className={`mt-1 text-sm ${ui.muted}`}>Upload a ledger file to generate financial reports.</p>
                <p className={`mt-1 text-xs ${ui.muted}`}>Built by <span className={`font-semibold ${ui.body}`}>Lovenish Purohit</span></p>
              </div>
              <ThemeToggle />
            </div>

            <div className="grid gap-6 lg:grid-cols-2">
              <div className="space-y-4">
                <div className={`rounded-2xl border p-5 ${ui.panel}`}>
                  <h2 className={`text-sm font-semibold uppercase tracking-widest ${ui.label}`}>Sample Files</h2>
                  <div className="mt-3 flex flex-wrap gap-3">
                    <a href="/samples/sample-ledger.csv" download className="rounded-lg bg-cyan-600 px-4 py-2 text-sm font-semibold text-white shadow transition-colors hover:bg-cyan-700">↓ CSV Sample</a>
                    <a href="/samples/sample-ledger.xlsx" download className="rounded-lg bg-cyan-600 px-4 py-2 text-sm font-semibold text-white shadow transition-colors hover:bg-cyan-700">↓ XLSX Sample</a>
                  </div>
                </div>
                <FileDropzone fileName={fileName} isDragging={isDragging} theme={theme}
                  onFileChange={handleFileChange} onDragOver={handleDragOver}
                  onDragLeave={() => setIsDragging(false)} onDrop={handleDrop} />
                {showMapper && fileHeaders.length > 0 && (
                  <ColumnMapper theme={theme} fileHeaders={fileHeaders}
                    onMappingConfirmed={handleMappingConfirmed}
                    onCancel={() => { setShowMapper(false); setHeaderErrors([]); setPendingFile(null); }} />
                )}
                {!showMapper && (uploadError || headerErrors.length > 0 || rowIssues.length > 0) && (
                  <ValidationPanel headerErrors={headerErrors} uploadError={uploadError} rowIssues={rowIssues} theme={theme} />
                )}
              </div>

              <div className="space-y-4">
                <QBOGuide theme={theme} />
                <div className={`rounded-2xl border p-5 ${ui.panel}`}>
                  <h2 className={`text-sm font-semibold uppercase tracking-widest ${ui.label}`}>Required Headers</h2>
                  <p className={`mt-1 text-xs ${ui.muted}`}>Columns marked ✱ are strictly mandatory.</p>
                  <div className="mt-3 grid gap-2 sm:grid-cols-2">
                    {requiredHeaders.map((h) => {
                      const mandatory = CORE_MANDATORY.includes(h as typeof CORE_MANDATORY[number]);
                      return (
                        <div key={h} className={`rounded-xl px-4 py-2.5 text-sm flex items-center justify-between ${mandatory ? (theme === 'dark' ? 'bg-cyan-950/60 text-cyan-200 border border-cyan-700' : 'bg-cyan-50 text-cyan-800 border border-cyan-200') : `${ui.card} ${ui.body}`}`}>
                          <span>{h}</span>
                          {mandatory && <span className="text-xs font-bold opacity-70">✱</span>}
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ─── Dashboard View ─────────────────────────────────────────────────────────

  const TABS: { id: Tab; label: string }[] = [
    { id: 'overview', label: 'Overview' },
    { id: 'insights', label: '🔍 Insights' },
    { id: 'pl',       label: 'P&L' },
    { id: 'bs',       label: 'Balance Sheet' },
    { id: 'cashflow', label: 'Cash Flow' },
    { id: 'mom',      label: 'Month-over-Month' },
    { id: 'preview',  label: 'Transactions' },
  ];

  return (
    <div className={`min-h-screen ${ui.page}`}>

      {/* Sticky Navbar */}
      <header className={`sticky top-0 z-50 border-b ${ui.navbar}`}>
        <div className="mx-auto flex max-w-7xl flex-wrap items-center justify-between gap-3 px-6 py-3">
          <div className="flex items-center gap-3">
            <h1 className={`text-base font-bold ${ui.heading}`}>Ledger Insight AI</h1>
            <span className={`hidden rounded-full px-3 py-1 text-xs font-medium sm:block ${ui.card} ${ui.muted}`}>{fileName}</span>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <ThemeToggle />
            {analysis && profitAndLoss && balanceSheet && cashFlow && (
              <>
                <button type="button" onClick={() => exportCsv(fileName, analysis, profitAndLoss, balanceSheet, cashFlow)}
                  className="rounded-lg bg-emerald-600 px-3 py-1.5 text-xs font-semibold text-white shadow transition-colors hover:bg-emerald-700">↓ CSV</button>
                <button type="button" onClick={() => exportExcel(fileName, analysis, profitAndLoss, balanceSheet, cashFlow, momPL ?? undefined)}
                  className="rounded-lg bg-blue-600 px-3 py-1.5 text-xs font-semibold text-white shadow transition-colors hover:bg-blue-700">↓ Excel</button>
                <button type="button" onClick={() => exportPdf(fileName, analysis, profitAndLoss, balanceSheet, cashFlow)}
                  className="rounded-lg bg-rose-600 px-3 py-1.5 text-xs font-semibold text-white shadow transition-colors hover:bg-rose-700">↓ PDF</button>
              </>
            )}
            <button type="button" onClick={handleUploadNew}
              className="rounded-lg bg-cyan-600 px-3 py-1.5 text-xs font-semibold text-white shadow transition-colors hover:bg-cyan-700">↑ New File</button>
          </div>
        </div>

        {/* Tab Bar */}
        <div className={`border-t ${ui.divider}`}>
          <div className="mx-auto flex max-w-7xl gap-0 overflow-x-auto px-6">
            {TABS.map((tab) => (
              <button key={tab.id} type="button" onClick={() => setActiveTab(tab.id)}
                className={`whitespace-nowrap border-b-2 px-4 py-3 text-sm font-medium transition-colors ${activeTab === tab.id ? ui.tabActive : ui.tabInactive}`}>
                {tab.label}
              </button>
            ))}
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-7xl space-y-6 px-6 py-8">

        {/* ── OVERVIEW TAB ── */}
        {activeTab === 'overview' && analysis && profitAndLoss && balanceSheet && cashFlow && (
          <div className="space-y-6">
            {/* KPI row */}
            <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
              {metricCard('Total Transactions', analysis.totalTransactions)}
              {metricCard('Total Revenue', fmt.format(profitAndLoss.totalRevenue))}
              {metricCard('Total Expenses', fmt.format(profitAndLoss.totalExpenses))}
              {metricCard('Net Profit', fmt.format(profitAndLoss.netProfit))}
            </div>

            {/* BS reconciliation card (DCF-style) */}
            <div className={`rounded-3xl border p-6 ${ui.panel}`}>
              <div className="flex flex-wrap items-center justify-between gap-3">
                <div>
                  <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>3-Statement Link Check</p>
                  <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Balance Sheet Reconciliation</h2>
                </div>
                <div className={`rounded-full px-3 py-1.5 text-sm font-semibold ${balanceSheet.isBalanced ? ui.successPill : ui.dangerPill}`}>
                  {balanceSheet.isBalanced ? '✅ Balanced' : '❌ Not Balanced'}
                </div>
              </div>

              <div className="mt-6 grid gap-4 sm:grid-cols-3">
                <div className={`rounded-2xl p-4 ${ui.card}`}>
                  <p className={`text-xs ${ui.label}`}>Total Assets</p>
                  <p className={`mt-1 text-2xl font-bold ${ui.stat}`}>{fmt.format(balanceSheet.totals.assetsTotal)}</p>
                </div>
                <div className={`rounded-2xl p-4 ${ui.card}`}>
                  <p className={`text-xs ${ui.label}`}>Liabilities + Equity</p>
                  <p className={`mt-1 text-2xl font-bold ${ui.stat}`}>{fmt.format(balanceSheet.totals.liabilitiesTotal + balanceSheet.totals.equityTotal)}</p>
                  <p className={`mt-1 text-xs ${ui.muted}`}>{fmt.format(balanceSheet.totals.liabilitiesTotal)} + {fmt.format(balanceSheet.totals.equityTotal)}</p>
                </div>
                <div className={`rounded-2xl p-4 ${Math.abs(balanceSheet.variance) <= 1 ? ui.successCard : ui.card}`}>
                  <p className={`text-xs ${ui.label}`}>Variance (A − L − E)</p>
                  <p className={`mt-1 text-2xl font-bold ${Math.abs(balanceSheet.variance) <= 1 ? '' : 'text-rose-400'}`}>{fmt.format(balanceSheet.variance)}</p>
                  <p className={`mt-1 text-xs ${ui.muted}`}>{Math.abs(balanceSheet.variance) <= 1 ? 'Within tolerance' : 'Check account type mapping'}</p>
                </div>
              </div>

              {/* Interlink explanation */}
              <div className={`mt-4 rounded-2xl p-4 text-sm ${ui.card}`}>
                <p className={`font-semibold ${ui.heading}`}>How statements are linked:</p>
                <div className={`mt-2 space-y-1 ${ui.muted}`}>
                  <p>📊 <strong className={ui.body}>P&L Net Profit</strong> ({fmt.format(profitAndLoss.netProfit)}) feeds into Balance Sheet as <strong className={ui.body}>Current Period Earnings</strong></p>
                  <p>💵 <strong className={ui.body}>Cash Flow</strong> starts from Net Profit ({fmt.format(cashFlow.netProfit)}) and adjusts for working capital changes</p>
                  <p>⚖️ <strong className={ui.body}>Balance Sheet</strong> checks: Assets = Liabilities + Equity (incl. CPE)</p>
                </div>
              </div>
            </div>

            {/* Quick analysis */}
            <div className="grid gap-6 lg:grid-cols-2">
              {analysis.inconsistentVendors.length > 0 && (
                <div className={`rounded-3xl border p-6 ${ui.panel}`}>
                  <h3 className={`text-sm font-semibold uppercase tracking-widest ${ui.label}`}>Inconsistent Vendors</h3>
                  <div className="mt-3 space-y-2 max-h-48 overflow-y-auto">
                    {analysis.inconsistentVendors.slice(0, 10).map(({ vendor, accounts, reason }) => (
                      <div key={`${vendor}-${reason}`} className={`rounded-xl px-4 py-2.5 text-sm ${ui.warningRow}`}>
                        <div><span className="font-semibold">{vendor}</span> <span className="text-xs opacity-70">— {reason}</span></div>
                        <div className="mt-0.5 text-xs opacity-80">{accounts.join(', ')}</div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
              {analysis.duplicates.length > 0 && (
                <div className={`rounded-3xl border p-6 ${ui.panel}`}>
                  <h3 className={`text-sm font-semibold uppercase tracking-widest ${ui.label}`}>Duplicate Transactions</h3>
                  <div className="mt-3 space-y-2 max-h-48 overflow-y-auto">
                    {analysis.duplicates.map(({ name, amount, transactionDate, occurrences }) => (
                      <div key={`${name}-${amount}-${transactionDate}`} className={`flex flex-wrap justify-between rounded-xl px-4 py-2.5 text-sm ${ui.warningRow}`}>
                        <span className="font-semibold">{name}</span>
                        <span>{amount} · {transactionDate} · {occurrences}×</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* ── INSIGHTS TAB ── */}
        {activeTab === 'insights' && insights && profitAndLoss && balanceSheet && (
          <div className="space-y-6">
            {/* Burn Rate & Runway */}
            <div className={`rounded-3xl border p-6 ${ui.panel}`}>
              <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Cash Intelligence</p>
              <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Burn Rate &amp; Runway</h2>
              <div className="mt-5 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
                {metricCard('Cash Balance', fmt.format(insights.cashBalance))}
                {metricCard('Avg Monthly Burn', fmt.format(insights.avgMonthlyBurn))}
                {metricCard('Avg Monthly Revenue', fmt.format(insights.avgMonthlyRevenue))}
                <div className={`rounded-2xl p-4 ${insights.runwayMonths !== null && insights.runwayMonths < 3 ? ui.dangerPill : insights.runwayMonths !== null && insights.runwayMonths < 6 ? ui.warningRow : ui.successCard}`}>
                  <p className="text-sm">Cash Runway</p>
                  <p className="mt-2 text-2xl font-semibold">{insights.runwayMonths !== null ? `${insights.runwayMonths.toFixed(1)} mo` : '∞'}</p>
                  <p className="mt-1 text-xs opacity-70">{insights.runwayMonths === null ? 'Revenue covers burn' : insights.runwayMonths < 3 ? '⚠️ Critical' : insights.runwayMonths < 6 ? '⚠️ Watch closely' : '✅ Healthy'}</p>
                </div>
              </div>
            </div>

            {/* Financial Ratios */}
            <div className={`rounded-3xl border p-6 ${ui.panel}`}>
              <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Solvency &amp; Liquidity</p>
              <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Financial Ratios</h2>
              <div className="mt-5 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
                {(() => {
                  const { currentRatio, quickRatio, debtToEquity, debtRatio } = balanceSheet.ratios;
                  const RATIOS = [
                    {
                      label: 'Current Ratio',
                      val: currentRatio,
                      good: (v: number) => v >= 1.5,
                      formula: 'Current Assets ÷ Current Liabilities',
                      numerator: fmt.format(balanceSheet.totals.currentAssetsTotal),
                      denominator: fmt.format(balanceSheet.totals.currentLiabilitiesTotal),
                      benchmark: '≥ 1.5 is healthy',
                    },
                    {
                      label: 'Quick Ratio',
                      val: quickRatio,
                      good: (v: number) => v >= 1.0,
                      formula: '(Current Assets − Inventory) ÷ Current Liabilities',
                      numerator: 'Liquid assets only',
                      denominator: fmt.format(balanceSheet.totals.currentLiabilitiesTotal),
                      benchmark: '≥ 1.0 is healthy',
                    },
                    {
                      label: 'Debt-to-Equity',
                      val: debtToEquity,
                      good: (v: number) => v <= 2.0,
                      formula: 'Total Liabilities ÷ Total Equity',
                      numerator: fmt.format(balanceSheet.totals.liabilitiesTotal),
                      denominator: fmt.format(balanceSheet.totals.equityTotal),
                      benchmark: '≤ 2.0 is healthy',
                    },
                    {
                      label: 'Debt Ratio',
                      val: debtRatio,
                      good: (v: number) => v <= 0.5,
                      formula: 'Total Liabilities ÷ Total Assets',
                      numerator: fmt.format(balanceSheet.totals.liabilitiesTotal),
                      denominator: fmt.format(balanceSheet.totals.assetsTotal),
                      benchmark: '≤ 50% is healthy',
                    },
                  ];
                  return <>
                    {RATIOS.map(({ label, val, good, formula, numerator, denominator, benchmark }) => (
                      <div key={label} className={`rounded-2xl p-4 ${val !== null && good(val) ? ui.successCard : val !== null ? ui.dangerPill : ui.card}`}>
                        <p className="text-sm font-semibold">{label}</p>
                        <p className="mt-2 text-2xl font-bold">{val !== null ? val.toFixed(2) : 'N/A'}</p>
                        <div className={`mt-2 text-[10px] space-y-0.5 opacity-80`}>
                          <p className="font-mono bg-black/10 rounded px-1.5 py-0.5">{formula}</p>
                          <p>{numerator} ÷ {denominator}</p>
                          <p className="font-semibold">{benchmark}</p>
                        </div>
                      </div>
                    ))}
                  </>;
                })()}
              </div>
              <div className={`mt-4 grid gap-4 sm:grid-cols-3 rounded-2xl p-4 ${ui.card}`}>
                {metricCard('Gross Margin', `${(profitAndLoss.grossMargin * 100).toFixed(1)}%`)}
                {metricCard('Net Margin', `${(profitAndLoss.netMargin * 100).toFixed(1)}%`)}
                {metricCard('Tax Estimate', fmt.format(insights.taxEstimate.amount))}
              </div>
            </div>

            {/* Top Vendors + Revenue Sources */}
            <div className="grid gap-6 lg:grid-cols-2">
              <div className={`rounded-3xl border p-6 ${ui.panel}`}>
                <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Pareto Analysis</p>
                <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Top Vendors by Spend</h2>
                <div className="mt-4 space-y-2">
                  {insights.topVendors.map((v, i) => {
                    const totalSpend = insights.topVendors.reduce((t, x) => t + x.total, 0);
                    const pct = totalSpend ? (v.total / totalSpend * 100).toFixed(1) : '0';
                    return (
                      <div key={v.name} className={`rounded-xl px-4 py-3 ${ui.listRow}`}>
                        <div className="flex items-center justify-between">
                          <span className={`font-semibold text-sm ${ui.stat}`}>#{i+1} {v.name}</span>
                          <span className="text-sm font-bold">{fmt.format(v.total)}</span>
                        </div>
                        <div className="mt-1.5 h-1.5 rounded-full bg-slate-700">
                          <div className="h-1.5 rounded-full bg-cyan-500" style={{ width: `${pct}%` }} />
                        </div>
                        <p className={`mt-1 text-xs ${ui.muted}`}>{pct}% of tracked spend · {v.txCount} transactions</p>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className={`rounded-3xl border p-6 ${ui.panel}`}>
                <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Pareto Analysis</p>
                <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Top Revenue Sources</h2>
                <div className="mt-4 space-y-2">
                  {insights.topRevenueSources.length === 0 ? (
                    <p className={`text-sm ${ui.muted}`}>No named revenue sources found.</p>
                  ) : insights.topRevenueSources.map((r, i) => {
                    const totalRev = insights.topRevenueSources.reduce((t, x) => t + x.total, 0);
                    const pct = totalRev ? (r.total / totalRev * 100).toFixed(1) : '0';
                    return (
                      <div key={r.name} className={`rounded-xl px-4 py-3 ${ui.listRow}`}>
                        <div className="flex items-center justify-between">
                          <span className={`font-semibold text-sm ${ui.stat}`}>#{i+1} {r.name}</span>
                          <span className="text-sm font-bold">{fmt.format(r.total)}</span>
                        </div>
                        <div className="mt-1.5 h-1.5 rounded-full bg-slate-700">
                          <div className="h-1.5 rounded-full bg-emerald-500" style={{ width: `${pct}%` }} />
                        </div>
                        <p className={`mt-1 text-xs ${ui.muted}`}>{pct}% of tracked revenue · {r.txCount} transactions</p>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            {/* Anomalies + Audit Flags */}
            <div className="grid gap-6 lg:grid-cols-2">
              {insights.anomalies.length > 0 && (
                <div className={`rounded-3xl border p-6 ${ui.panel}`}>
                  <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Statistical Detection</p>
                  <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Anomalous Transactions</h2>
                  <p className={`mt-1 text-sm ${ui.muted}`}>Transactions &gt;3σ above average amount.</p>
                  <div className="mt-4 space-y-2 max-h-72 overflow-y-auto">
                    {insights.anomalies.map((a, i) => (
                      <div key={i} className={`rounded-xl px-4 py-3 text-sm ${ui.warningRow}`}>
                        <div className="flex justify-between">
                          <span className="font-semibold">{a.row.Name || a.row['Distribution account']}</span>
                          <span className="font-bold">{fmt.format(a.amount)}</span>
                        </div>
                        <p className="text-xs mt-0.5 opacity-80">{a.row['Transaction date']} · {a.reason}</p>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {insights.auditFlags.length > 0 && (
                <div className={`rounded-3xl border p-6 ${ui.panel}`}>
                  <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Audit Checks</p>
                  <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Audit Flags</h2>
                  <div className="mt-4 space-y-3">
                    {insights.auditFlags.map((flag, i) => (
                      <div key={i} className={`rounded-xl px-4 py-3 ${ui.warningRow}`}>
                        <p className="text-sm font-semibold">{flag.type === 'round' ? '🔵' : flag.type === 'weekend' ? '📅' : flag.type === 'gap' ? '🔴' : '⚠️'} {flag.description}</p>
                        {flag.rows.length > 0 && (
                          <div className="mt-2 space-y-1">
                            {flag.rows.slice(0, 3).map((r, j) => (
                              <p key={j} className="text-xs opacity-80">{r['Transaction date']} · {r.Name} · {fmt.format(Math.abs(parseCurrencyAmountClient(r.Amount)))}</p>
                            ))}
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>

            {/* Category Mismatches */}
            {insights.categoryMismatches.length > 0 && (
              <div className={`rounded-3xl border p-6 ${ui.panel}`}>
                <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>AI Category Check</p>
                <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Possible Miscategorised Transactions</h2>
                <p className={`mt-1 text-sm ${ui.muted}`}>Based on memo/description text — verify these manually.</p>
                <div className="mt-4 space-y-2 max-h-72 overflow-y-auto">
                  {insights.categoryMismatches.map((m, i) => (
                    <div key={i} className={`rounded-xl px-4 py-3 text-sm ${ui.warningRow}`}>
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <span className="font-semibold">{m.row['Distribution account']}</span>
                        <div className="flex items-center gap-1.5 text-xs">
                          <span className={`px-2 py-0.5 rounded-full font-semibold ${theme === 'dark' ? 'bg-rose-900/60 text-rose-300' : 'bg-rose-100 text-rose-700'}`}>{m.assignedType}</span>
                          <span className={ui.muted}>→ suggests</span>
                          <span className={`px-2 py-0.5 rounded-full font-semibold ${theme === 'dark' ? 'bg-emerald-900/60 text-emerald-300' : 'bg-emerald-100 text-emerald-700'}`}>{m.suggestedType}</span>
                        </div>
                      </div>
                      <p className="text-xs mt-1 opacity-80">"{m.description.slice(0, 80)}{m.description.length > 80 ? '…' : ''}"</p>
                      <p className={`text-xs mt-0.5 ${ui.muted}`}>{m.reason}</p>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Tax estimate */}
            <div className={`rounded-3xl border p-6 ${ui.panel}`}>
              <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Planning</p>
              <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Tax Estimate</h2>
              <div className="mt-5 grid gap-4 sm:grid-cols-3">
                {metricCard('Rate', `${(insights.taxEstimate.rate * 100).toFixed(0)}%`)}
                {metricCard('Estimated Tax', fmt.format(insights.taxEstimate.amount))}
                <div className={`rounded-2xl p-4 text-sm ${ui.card} ${ui.muted}`}>
                  <p className="font-semibold mb-1">Basis</p>
                  <p>{insights.taxEstimate.basis}</p>
                  <p className="mt-2 text-xs opacity-70">⚠️ Estimate only — consult a tax professional.</p>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ── P&L TAB ── */}
        {activeTab === 'pl' && profitAndLoss && (
          <div className={`rounded-3xl border p-6 ${ui.panel}`}>
            <div className="flex flex-wrap items-start justify-between gap-3">
              <div>
                <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Financials</p>
                <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Profit &amp; Loss</h2>
              </div>
              <div className={`rounded-full px-3 py-1.5 text-xs font-semibold ${ui.successPill}`}>
                Net {fmt.format(profitAndLoss.netProfit)}
              </div>
            </div>
            <div className="mt-5 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
              {metricCard('Revenue', fmt.format(profitAndLoss.totalRevenue))}
              {metricCard('COGS', fmt.format(profitAndLoss.totalCogs))}
              {metricCard('Gross Profit', `${fmt.format(profitAndLoss.grossProfit)} (${(profitAndLoss.grossMargin*100).toFixed(1)}%)`)}
              <div className={`rounded-2xl border p-4 ${ui.successCard}`}>
                <p className="text-sm">Net Profit</p>
                <p className="mt-2 text-2xl font-semibold">{fmt.format(profitAndLoss.netProfit)}</p>
                <p className="mt-1 text-xs opacity-70">Margin: {(profitAndLoss.netMargin*100).toFixed(1)}%</p>
              </div>
            </div>
            {Object.keys(profitAndLoss.monthlyBreakdown).length > 0 && (
              <div className={`mt-4 rounded-2xl p-4 ${ui.card}`}>
                <h3 className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Monthly Breakdown</h3>
                <div className="mt-3 space-y-2">
                  {Object.entries(profitAndLoss.monthlyBreakdown).map(([month, { revenue, expenses }]) => (
                    <div key={month} className={`flex flex-wrap justify-between rounded-xl px-4 py-3 text-sm ${ui.listRow}`}>
                      <span className={`font-semibold ${ui.stat}`}>{month}</span>
                      <span>Rev: {fmt.format(revenue)} · Exp: {fmt.format(expenses)} · Net: {fmt.format(revenue - expenses)}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* ── BALANCE SHEET TAB ── */}
        {activeTab === 'bs' && balanceSheet && (
          <div className="space-y-6">
            <div className={`rounded-3xl border p-6 ${ui.panel}`}>
              <div className="flex flex-wrap items-start justify-between gap-3">
                <div>
                  <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Financials</p>
                  <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Balance Sheet</h2>
                  <p className={`mt-1 text-sm ${ui.muted}`}>Latest balance per account grouped by type.</p>
                </div>
                <div className={`rounded-full px-3 py-1.5 text-sm font-semibold ${balanceSheet.isBalanced ? ui.successPill : ui.dangerPill}`}>
                  {balanceSheet.isBalanced ? 'Balanced ✅' : `Not Balanced ❌ (Δ ${fmt.format(balanceSheet.variance)})`}
                </div>
              </div>
              <div className="mt-5 grid gap-4 sm:grid-cols-3">
                {metricCard('Assets Total', fmt.format(balanceSheet.totals.assetsTotal))}
                {metricCard('Liabilities Total', fmt.format(balanceSheet.totals.liabilitiesTotal))}
                {metricCard('Equity Total', fmt.format(balanceSheet.totals.equityTotal))}
              </div>
              <div className="mt-4 grid gap-4 lg:grid-cols-3">
                {(['assets', 'liabilities', 'equity'] as const).map((section) => (
                  <div key={section} className={`rounded-2xl p-4 ${ui.card}`}>
                    <h3 className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>{section}</h3>
                    <div className="mt-3 max-h-64 space-y-2 overflow-y-auto">
                      {balanceSheet[section].length === 0 ? (
                        <p className={`text-sm ${ui.muted}`}>No accounts.</p>
                      ) : balanceSheet[section].map((e) => (
                        <div key={e.account} className={`flex flex-wrap justify-between rounded-xl px-3 py-2 text-sm ${ui.listRow}`}>
                          <span className={`font-medium ${ui.stat}`}>{e.account}</span>
                          <span>{fmt.format(e.value)}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>

              {/* Reconciliation footer */}
              <div className={`mt-4 rounded-2xl p-4 text-sm border ${Math.abs(balanceSheet.variance) <= 1 ? ui.successCard : ui.dangerPill} `}>
                <div className="flex flex-wrap gap-6">
                  <span>Assets: <strong>{fmt.format(balanceSheet.totals.assetsTotal)}</strong></span>
                  <span>=</span>
                  <span>Liabilities: <strong>{fmt.format(balanceSheet.totals.liabilitiesTotal)}</strong></span>
                  <span>+</span>
                  <span>Equity: <strong>{fmt.format(balanceSheet.totals.equityTotal)}</strong></span>
                  <span className="ml-auto">Variance: <strong>{fmt.format(balanceSheet.variance)}</strong></span>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ── CASH FLOW TAB ── */}
        {activeTab === 'cashflow' && cashFlow && (
          <div className={`rounded-3xl border p-6 ${ui.panel}`}>
            <div className="flex flex-wrap items-start justify-between gap-3">
              <div>
                <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Financials</p>
                <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Cash Flow Statement</h2>
                <p className={`mt-1 text-sm ${ui.muted}`}>Indirect method — starts from Net Profit.</p>
              </div>
              <div className={`rounded-full px-3 py-1.5 text-xs font-semibold ${ui.skyPill}`}>
                Operating CF: {fmt.format(cashFlow.operatingCashFlow)}
              </div>
            </div>
            <div className="mt-5 grid gap-4 sm:grid-cols-2">
              {metricCard('Net Profit', fmt.format(cashFlow.netProfit))}
              <div className={`rounded-2xl border p-4 ${ui.skyCard}`}>
                <p className="text-sm">Operating Cash Flow</p>
                <p className="mt-2 text-2xl font-semibold">{fmt.format(cashFlow.operatingCashFlow)}</p>
              </div>
            </div>
            {cashFlow.adjustments.length > 0 && (
              <div className={`mt-4 rounded-2xl p-4 ${ui.card}`}>
                <h3 className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Working Capital Adjustments</h3>
                <div className="mt-3 space-y-2 max-h-80 overflow-y-auto">
                  {cashFlow.adjustments.map((adj) => (
                    <div key={adj.account} className={`flex flex-wrap justify-between rounded-xl px-4 py-3 text-sm ${ui.listRow}`}>
                      <span className={`font-semibold ${ui.stat}`}>{adj.account}</span>
                      <span>Δ {fmt.format(adj.change)} · Impact: {fmt.format(adj.impact)}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* ── MONTH-OVER-MONTH TAB ── */}
        {activeTab === 'mom' && momPL && (
          <div className={`rounded-3xl border p-6 ${ui.panel}`}>
            <div className="flex flex-wrap items-start justify-between gap-3">
              <div>
                <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Analytics</p>
                <h2 className={`mt-0.5 text-xl font-bold ${ui.heading}`}>Month-over-Month P&amp;L</h2>
                <p className={`mt-1 text-sm ${ui.muted}`}>Revenue and expenses by category across months.</p>
              </div>
              <div className={`flex rounded-full p-1 text-xs font-semibold ${ui.settingsControl}`}>
                <button type="button" onClick={() => setShowMoMPct(false)}
                  className={`rounded-full px-3 py-1.5 transition-colors ${!showMoMPct ? ui.settingsOptionActive : ui.settingsOptionInactive}`}>$ Amount</button>
                <button type="button" onClick={() => setShowMoMPct(true)}
                  className={`rounded-full px-3 py-1.5 transition-colors ${showMoMPct ? ui.settingsOptionActive : ui.settingsOptionInactive}`}>% MoM</button>
              </div>
            </div>

            <div className="mt-6 overflow-auto max-h-[70vh] rounded-2xl">
              <table className="w-full text-sm border-separate border-spacing-0">
                <thead className="sticky top-0 z-20">
                  <tr>
                    <th className={`sticky left-0 z-30 rounded-tl-xl px-4 py-3 text-left text-xs font-semibold uppercase tracking-wider ${ui.tableHead}`}>Account</th>
                    {momPL.months.map((m) => (
                      <th key={m} className={`px-4 py-3 text-right text-xs font-semibold uppercase tracking-wider ${ui.tableHead}`}>{monthLabel(m)}</th>
                    ))}
                    <th className={`rounded-tr-xl px-4 py-3 text-right text-xs font-semibold uppercase tracking-wider ${ui.tableHead}`}>Total</th>
                  </tr>
                </thead>
                <tbody>
                  {/* Income section */}
                  <tr>
                    <td colSpan={momPL.months.length + 2} className={`px-4 py-2 text-xs font-bold uppercase tracking-widest ${ui.muted} ${ui.card}`}>
                      ▸ Income
                    </td>
                  </tr>
                  {momPL.incomeCategories.map((cat, rowIdx) => (
                    <tr key={cat.name}>
                      <td className={`sticky left-0 z-10 px-4 py-2.5 font-medium ${rowIdx % 2 === 0 ? ui.tableRow : ui.tableAlt}`}>{cat.name}</td>
                      {momPL.months.map((m, mIdx) => {
                        const val = cat.months[m] ?? 0;
                        const prev = mIdx > 0 ? (cat.months[momPL.months[mIdx - 1]] ?? 0) : null;
                        const pct = showMoMPct && prev !== null ? momChange(val, prev) : null;
                        return (
                          <td key={m} className={`px-4 py-2.5 text-right ${rowIdx % 2 === 0 ? ui.tableRow : ui.tableAlt}`}>
                            {showMoMPct && mIdx > 0 ? (
                              pct === null ? <span className={ui.muted}>—</span>
                                : <span className={pct >= 0 ? 'text-emerald-400' : 'text-rose-400'}>{pct >= 0 ? '+' : ''}{pct.toFixed(1)}%</span>
                            ) : val > 0 ? fmt.format(val) : <span className={ui.muted}>—</span>}
                          </td>
                        );
                      })}
                      <td className={`px-4 py-2.5 text-right font-semibold ${rowIdx % 2 === 0 ? ui.tableRow : ui.tableAlt}`}>{fmt.format(cat.total)}</td>
                    </tr>
                  ))}
                  {/* Total Revenue row */}
                  <tr>
                    <td className={`sticky left-0 z-10 px-4 py-2.5 font-bold ${ui.tableHead}`}>Total Revenue</td>
                    {momPL.months.map((m, mIdx) => {
                      const val = momPL.monthlyRevenue[m] ?? 0;
                      const prev = mIdx > 0 ? (momPL.monthlyRevenue[momPL.months[mIdx-1]] ?? 0) : null;
                      const pct = showMoMPct && prev !== null ? momChange(val, prev) : null;
                      return (
                        <td key={m} className={`px-4 py-2.5 text-right font-bold ${ui.tableHead}`}>
                          {showMoMPct && mIdx > 0 ? (pct === null ? '—' : `${pct >= 0 ? '+' : ''}${pct.toFixed(1)}%`) : fmt.format(val)}
                        </td>
                      );
                    })}
                    <td className={`px-4 py-2.5 text-right font-bold ${ui.tableHead}`}>{fmt.format(momPL.totalRevenue)}</td>
                  </tr>

                  {/* Spacer */}
                  <tr><td colSpan={momPL.months.length + 2} className="py-2" /></tr>

                  {/* Expense section */}
                  <tr>
                    <td colSpan={momPL.months.length + 2} className={`px-4 py-2 text-xs font-bold uppercase tracking-widest ${ui.muted} ${ui.card}`}>
                      ▸ Expenses
                    </td>
                  </tr>
                  {momPL.expenseCategories.map((cat, rowIdx) => (
                    <tr key={cat.name}>
                      <td className={`sticky left-0 z-10 px-4 py-2.5 font-medium ${rowIdx % 2 === 0 ? ui.tableRow : ui.tableAlt}`}>{cat.name}</td>
                      {momPL.months.map((m, mIdx) => {
                        const val = cat.months[m] ?? 0;
                        const prev = mIdx > 0 ? (cat.months[momPL.months[mIdx - 1]] ?? 0) : null;
                        const pct = showMoMPct && prev !== null ? momChange(val, prev) : null;
                        return (
                          <td key={m} className={`px-4 py-2.5 text-right ${rowIdx % 2 === 0 ? ui.tableRow : ui.tableAlt}`}>
                            {showMoMPct && mIdx > 0 ? (
                              pct === null ? <span className={ui.muted}>—</span>
                                : <span className={pct >= 0 ? 'text-rose-400' : 'text-emerald-400'}>{pct >= 0 ? '+' : ''}{pct.toFixed(1)}%</span>
                            ) : val > 0 ? fmt.format(val) : <span className={ui.muted}>—</span>}
                          </td>
                        );
                      })}
                      <td className={`px-4 py-2.5 text-right font-semibold ${rowIdx % 2 === 0 ? ui.tableRow : ui.tableAlt}`}>{fmt.format(cat.total)}</td>
                    </tr>
                  ))}
                  {/* Total Expenses row */}
                  <tr>
                    <td className={`sticky left-0 z-10 px-4 py-2.5 font-bold ${ui.tableHead}`}>Total Expenses</td>
                    {momPL.months.map((m, mIdx) => {
                      const val = momPL.monthlyExpenses[m] ?? 0;
                      const prev = mIdx > 0 ? (momPL.monthlyExpenses[momPL.months[mIdx-1]] ?? 0) : null;
                      const pct = showMoMPct && prev !== null ? momChange(val, prev) : null;
                      return (
                        <td key={m} className={`px-4 py-2.5 text-right font-bold ${ui.tableHead}`}>
                          {showMoMPct && mIdx > 0 ? (pct === null ? '—' : `${pct >= 0 ? '+' : ''}${pct.toFixed(1)}%`) : fmt.format(val)}
                        </td>
                      );
                    })}
                    <td className={`px-4 py-2.5 text-right font-bold ${ui.tableHead}`}>{fmt.format(momPL.totalExpenses)}</td>
                  </tr>

                  {/* Net Profit row */}
                  <tr>
                    <td className={`sticky left-0 z-10 px-4 py-3 font-bold ${ui.successCard} rounded-bl-xl`}>Net Profit</td>
                    {momPL.months.map((m, mIdx) => {
                      const val = momPL.monthlyNetProfit[m] ?? 0;
                      const prev = mIdx > 0 ? (momPL.monthlyNetProfit[momPL.months[mIdx-1]] ?? 0) : null;
                      const pct = showMoMPct && prev !== null ? momChange(val, prev) : null;
                      return (
                        <td key={m} className={`px-4 py-3 text-right font-bold ${val >= 0 ? ui.successCard : ui.dangerPill}`}>
                          {showMoMPct && mIdx > 0
                            ? (pct === null ? '—' : `${pct >= 0 ? '+' : ''}${pct.toFixed(1)}%`)
                            : fmt.format(val)}
                        </td>
                      );
                    })}
                    <td className={`px-4 py-3 text-right font-bold rounded-br-xl ${momPL.totalNetProfit >= 0 ? ui.successCard : ui.dangerPill}`}>{fmt.format(momPL.totalNetProfit)}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ── TRANSACTIONS TAB ── */}
        {activeTab === 'preview' && previewRows.length > 0 && (
          <div className={`rounded-3xl border p-6 ${ui.panel}`}>
            <p className={`text-xs font-semibold uppercase tracking-widest ${ui.label}`}>Raw Data</p>
            <h2 className={`mt-0.5 mb-4 text-xl font-bold ${ui.heading}`}>Transaction Preview <span className={`text-sm font-normal ${ui.muted}`}>(first 100 rows)</span></h2>
            <PreviewTable columns={requiredHeaders} rows={previewRows} rowIssues={rowIssues} theme={theme} />
          </div>
        )}

      </main>
    </div>
  );
}
