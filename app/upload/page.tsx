'use client';

import { useCallback, useMemo, useState } from 'react';
import { analyzeLedger, type LedgerAnalysis } from '../../src/lib/analyzeLedger';
import { generateBalanceSheet, type BalanceSheet } from '../../src/lib/generateBalanceSheet';
import { generateCashFlow, type CashFlowStatement } from '../../src/lib/generateCashFlow';
import { generatePL, type ProfitAndLoss } from '../../src/lib/generatePL';
import { FileDropzone, type UploadTheme } from './components/FileDropzone';
import { PreviewTable } from './components/PreviewTable';
import { ValidationPanel } from './components/ValidationPanel';
import {
  isCsvFile,
  isExcelFile,
  parseCsvFile,
  parseXlsxFile,
  requiredHeaders,
  type LedgerRow,
  type RowIssue,
} from './upload-utils';

type ThemeClasses = {
  page: string;
  shell: string;
  panel: string;
  card: string;
  nestedCard: string;
  listRow: string;
  heading: string;
  body: string;
  muted: string;
  stat: string;
  label: string;
  accentPill: string;
  successPill: string;
  dangerPill: string;
  successCard: string;
  skyPill: string;
  skyCard: string;
  warningRow: string;
  settingsPanel: string;
  settingsControl: string;
  settingsOptionActive: string;
  settingsOptionInactive: string;
};

const themes: Record<UploadTheme, ThemeClasses> = {
  dark: {
    page: 'bg-slate-950 text-slate-100',
    shell: 'border-slate-700 bg-slate-900/80 shadow-2xl shadow-slate-950/40',
    panel: 'border-slate-700 bg-slate-900',
    card: 'bg-slate-950/60',
    nestedCard: 'bg-slate-900',
    listRow: 'bg-slate-950/50 text-slate-300',
    heading: 'text-white',
    body: 'text-slate-300',
    muted: 'text-slate-400',
    stat: 'text-white',
    label: 'text-slate-400',
    accentPill: 'bg-cyan-950/70 text-cyan-200',
    successPill: 'bg-emerald-950/70 text-emerald-200',
    dangerPill: 'bg-rose-950/70 text-rose-200',
    successCard: 'border-emerald-500/40 bg-emerald-950/50 text-emerald-100',
    skyPill: 'bg-sky-950/70 text-sky-200',
    skyCard: 'border-sky-500/40 bg-sky-950/50 text-sky-100',
    warningRow: 'bg-amber-950/50 text-amber-200',
    settingsPanel: 'border-slate-700 bg-slate-950/50',
    settingsControl: 'bg-slate-950/70 text-slate-200',
    settingsOptionActive: 'bg-cyan-400 text-slate-950',
    settingsOptionInactive: 'text-slate-300 hover:bg-slate-800',
  },
  light: {
    page: 'bg-slate-100 text-slate-900',
    shell: 'border-slate-200 bg-white shadow-xl shadow-slate-200/80',
    panel: 'border-slate-200 bg-slate-50',
    card: 'bg-white',
    nestedCard: 'bg-slate-50',
    listRow: 'bg-slate-50 text-slate-700',
    heading: 'text-slate-950',
    body: 'text-slate-600',
    muted: 'text-slate-500',
    stat: 'text-slate-950',
    label: 'text-slate-500',
    accentPill: 'bg-cyan-100 text-cyan-800',
    successPill: 'bg-emerald-100 text-emerald-800',
    dangerPill: 'bg-rose-100 text-rose-800',
    successCard: 'border-emerald-200 bg-emerald-50 text-emerald-900',
    skyPill: 'bg-sky-100 text-sky-800',
    skyCard: 'border-sky-200 bg-sky-50 text-sky-900',
    warningRow: 'bg-amber-50 text-amber-900',
    settingsPanel: 'border-slate-200 bg-white',
    settingsControl: 'bg-slate-100 text-slate-700',
    settingsOptionActive: 'bg-slate-900 text-white',
    settingsOptionInactive: 'text-slate-600 hover:bg-slate-200',
  },
};

export default function UploadPage() {
  const [theme, setTheme] = useState<UploadTheme>('dark');
  const [previewRows, setPreviewRows] = useState<LedgerRow[]>([]);
  const [analysis, setAnalysis] = useState<LedgerAnalysis | null>(null);
  const [profitAndLoss, setProfitAndLoss] = useState<ProfitAndLoss | null>(null);
  const [balanceSheet, setBalanceSheet] = useState<BalanceSheet | null>(null);
  const [cashFlow, setCashFlow] = useState<CashFlowStatement | null>(null);
  const [headerErrors, setHeaderErrors] = useState<string[]>([]);
  const [rowIssues, setRowIssues] = useState<RowIssue[]>([]);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [fileName, setFileName] = useState('');
  const [isDragging, setIsDragging] = useState(false);

  const ui = themes[theme];

  const clearState = () => {
    setUploadError(null);
    setHeaderErrors([]);
    setRowIssues([]);
    setPreviewRows([]);
    setAnalysis(null);
    setProfitAndLoss(null);
    setBalanceSheet(null);
    setCashFlow(null);
  };

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

      if (result.headerErrors.length > 0) {
        setPreviewRows([]);
        return;
      }

      const profitAndLossResult = generatePL(result.rows);

      setAnalysis(analyzeLedger(result.rows));
      setProfitAndLoss(profitAndLossResult);
      setBalanceSheet(generateBalanceSheet(result.rows));
      setCashFlow(generateCashFlow(result.rows, profitAndLossResult));
      setPreviewRows(result.rows.slice(0, 50));
    } catch {
      setUploadError('Unable to parse file. Please verify the file format.');
    }
  }, []);

  const handleFileChange = useCallback(
    async (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (file) {
        await handleParse(file);
      }
    },
    [handleParse],
  );

  const handleDrop = useCallback(
    async (event: React.DragEvent<HTMLLabelElement>) => {
      event.preventDefault();
      setIsDragging(false);
      const file = event.dataTransfer.files[0];
      if (file) {
        await handleParse(file);
      }
    },
    [handleParse],
  );

  const handleDragOver = useCallback((event: React.DragEvent<HTMLLabelElement>) => {
    event.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback(() => {
    setIsDragging(false);
  }, []);

  const previewColumns = useMemo(() => requiredHeaders, []);
  const currencyFormatter = useMemo(
    () =>
      new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
      }),
    [],
  );

  const renderBalanceSheetEntries = (entries: BalanceSheet['assets']) => {
    if (entries.length === 0) {
      return <p className={`mt-3 text-sm ${ui.muted}`}>No accounts available.</p>;
    }

    return (
      <div className="mt-3 space-y-2">
        {entries.map((entry) => (
          <div key={entry.account} className={`flex flex-wrap items-center justify-between gap-3 rounded-2xl px-4 py-3 text-sm ${ui.listRow}`}>
            <span className={`font-semibold ${ui.stat}`}>{entry.account}</span>
            <span>{currencyFormatter.format(entry.value)}</span>
          </div>
        ))}
      </div>
    );
  };

  const metricCard = (label: string, value: string | number) => (
    <div className={`rounded-2xl p-4 shadow-sm ${ui.card}`}>
      <p className={`text-sm ${ui.label}`}>{label}</p>
      <p className={`mt-2 text-2xl font-semibold ${ui.stat}`}>{value}</p>
    </div>
  );

  return (
    <div className={`min-h-screen px-4 py-10 ${ui.page}`}>
      <div className={`mx-auto w-full max-w-6xl rounded-3xl border p-8 ${ui.shell}`}>
        <div className="space-y-4">
          <div className={`flex flex-wrap items-center justify-between gap-4 rounded-3xl border p-4 ${ui.settingsPanel}`}>
            <div>
              <p className={`text-sm font-semibold uppercase tracking-[0.2em] ${ui.muted}`}>Settings</p>
              <h2 className={`mt-1 text-xl font-semibold ${ui.heading}`}>Theme</h2>
            </div>
            <div className={`flex rounded-full p-1 text-sm font-semibold ${ui.settingsControl}`}>
              {(['dark', 'light'] as const).map((themeOption) => (
                <button
                  key={themeOption}
                  type="button"
                  onClick={() => setTheme(themeOption)}
                  className={`rounded-full px-4 py-2 capitalize transition-colors ${
                    theme === themeOption ? ui.settingsOptionActive : ui.settingsOptionInactive
                  }`}
                  aria-pressed={theme === themeOption}
                >
                  {themeOption}
                </button>
              ))}
            </div>
          </div>

          <FileDropzone
            fileName={fileName}
            isDragging={isDragging}
            theme={theme}
            onFileChange={handleFileChange}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          />

          <div className="grid gap-4 lg:grid-cols-[1fr_320px]">
            <div className="space-y-4">
              <div className={`rounded-3xl border p-5 ${ui.panel}`}>
                <h2 className={`text-lg font-semibold ${ui.heading}`}>Required Headers</h2>
                <div className="mt-4 grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
                  {requiredHeaders.map((header) => (
                    <div key={header} className={`rounded-2xl px-4 py-3 text-sm shadow-sm ${ui.card} ${ui.body}`}>
                      {header}
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <ValidationPanel headerErrors={headerErrors} uploadError={uploadError} rowIssues={rowIssues} theme={theme} />
          </div>

          <PreviewTable columns={previewColumns} rows={previewRows} rowIssues={rowIssues} theme={theme} />

          {analysis ? (
            <div className={`rounded-3xl border p-5 ${ui.panel}`}>
              <div className="flex flex-wrap items-start justify-between gap-4">
                <div>
                  <h2 className={`text-lg font-semibold ${ui.heading}`}>Ledger Analysis</h2>
                  <p className={`mt-1 text-sm ${ui.body}`}>Summary of the parsed transaction set.</p>
                </div>
                <div className={`text-sm ${ui.muted}`}>Total transactions: {analysis.totalTransactions}</div>
              </div>

              <div className="mt-4 grid gap-4 sm:grid-cols-3">
                {metricCard('Total transactions', analysis.totalTransactions)}
                {metricCard('Inconsistent vendors', analysis.inconsistentVendors.length)}
                {metricCard('Duplicate count', analysis.duplicates.length)}
              </div>

              <div className={`mt-4 rounded-2xl p-4 shadow-sm ${ui.card}`}>
                <h3 className={`text-sm font-semibold uppercase tracking-[0.16em] ${ui.label}`}>Inconsistent Vendors</h3>
                {analysis.inconsistentVendors.length > 0 ? (
                  <div className="mt-3 space-y-2">
                    {analysis.inconsistentVendors.map(({ vendor, accounts }) => (
                      <div key={vendor} className={`rounded-2xl px-4 py-3 text-sm ${ui.warningRow}`}>
                        <span className="font-semibold">{vendor}</span>: {accounts.join(', ')}
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className={`mt-3 text-sm ${ui.muted}`}>No vendors were found with multiple distribution accounts.</p>
                )}
              </div>
            </div>
          ) : null}

          {profitAndLoss ? (
            <div className={`rounded-3xl border p-5 ${ui.panel}`}>
              <div className="flex flex-wrap items-start justify-between gap-4">
                <div>
                  <h2 className={`text-lg font-semibold ${ui.heading}`}>Profit &amp; Loss</h2>
                  <p className={`mt-1 text-sm ${ui.body}`}>Revenue and expense totals generated from ledger account types.</p>
                </div>
                <div className={`rounded-full px-4 py-2 text-sm font-semibold ${ui.successPill}`}>
                  Net Profit: {currencyFormatter.format(profitAndLoss.netProfit)}
                </div>
              </div>

              <div className="mt-4 grid gap-4 sm:grid-cols-3">
                {metricCard('Total Revenue', currencyFormatter.format(profitAndLoss.totalRevenue))}
                {metricCard('Total Expenses', currencyFormatter.format(profitAndLoss.totalExpenses))}
                <div className={`rounded-2xl border p-4 shadow-sm ${ui.successCard}`}>
                  <p className="text-sm">Net Profit</p>
                  <p className="mt-2 text-2xl font-semibold">{currencyFormatter.format(profitAndLoss.netProfit)}</p>
                </div>
              </div>

              <div className={`mt-4 rounded-2xl p-4 shadow-sm ${ui.card}`}>
                <h3 className={`text-sm font-semibold uppercase tracking-[0.16em] ${ui.label}`}>Monthly Breakdown</h3>
                {Object.keys(profitAndLoss.monthlyBreakdown).length > 0 ? (
                  <div className="mt-3 space-y-2">
                    {Object.entries(profitAndLoss.monthlyBreakdown).map(([month, summary]) => (
                      <div key={month} className={`flex flex-wrap items-center justify-between gap-3 rounded-2xl px-4 py-3 text-sm ${ui.listRow}`}>
                        <span className={`font-semibold ${ui.stat}`}>{month}</span>
                        <span>
                          Revenue: {currencyFormatter.format(summary.revenue)} | Expenses:{' '}
                          {currencyFormatter.format(summary.expenses)}
                        </span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className={`mt-3 text-sm ${ui.muted}`}>No income or expense transactions with valid dates were available for monthly summaries.</p>
                )}
              </div>
            </div>
          ) : null}

          {balanceSheet ? (
            <div className={`rounded-3xl border p-5 ${ui.panel}`}>
              <div className="flex flex-wrap items-start justify-between gap-4">
                <div>
                  <h2 className={`text-lg font-semibold ${ui.heading}`}>Balance Sheet</h2>
                  <p className={`mt-1 text-sm ${ui.body}`}>Latest balance per account grouped into assets, liabilities, and equity.</p>
                </div>
                <div className={`rounded-full px-4 py-2 text-sm font-semibold ${balanceSheet.isBalanced ? ui.successPill : ui.dangerPill}`}>
                  {balanceSheet.isBalanced ? 'Balanced \u2705' : 'Not Balanced \u274C'}
                </div>
              </div>

              <div className="mt-4 grid gap-4 lg:grid-cols-3">
                <div className={`rounded-2xl p-4 shadow-sm ${ui.card}`}>
                  <h3 className={`text-sm font-semibold uppercase tracking-[0.16em] ${ui.label}`}>Assets</h3>
                  {renderBalanceSheetEntries(balanceSheet.assets)}
                </div>
                <div className={`rounded-2xl p-4 shadow-sm ${ui.card}`}>
                  <h3 className={`text-sm font-semibold uppercase tracking-[0.16em] ${ui.label}`}>Liabilities</h3>
                  {renderBalanceSheetEntries(balanceSheet.liabilities)}
                </div>
                <div className={`rounded-2xl p-4 shadow-sm ${ui.card}`}>
                  <h3 className={`text-sm font-semibold uppercase tracking-[0.16em] ${ui.label}`}>Equity</h3>
                  {renderBalanceSheetEntries(balanceSheet.equity)}
                </div>
              </div>

              <div className="mt-4 grid gap-4 sm:grid-cols-3">
                {metricCard('Assets Total', currencyFormatter.format(balanceSheet.totals.assetsTotal))}
                {metricCard('Liabilities Total', currencyFormatter.format(balanceSheet.totals.liabilitiesTotal))}
                {metricCard('Equity Total', currencyFormatter.format(balanceSheet.totals.equityTotal))}
              </div>
            </div>
          ) : null}

          {cashFlow ? (
            <div className={`rounded-3xl border p-5 ${ui.panel}`}>
              <div className="flex flex-wrap items-start justify-between gap-4">
                <div>
                  <h2 className={`text-lg font-semibold ${ui.heading}`}>Cash Flow Statement</h2>
                  <p className={`mt-1 text-sm ${ui.body}`}>Indirect method using net profit and working capital changes.</p>
                </div>
                <div className={`rounded-full px-4 py-2 text-sm font-semibold ${ui.skyPill}`}>
                  Operating Cash Flow: {currencyFormatter.format(cashFlow.operatingCashFlow)}
                </div>
              </div>

              <div className="mt-4 grid gap-4 sm:grid-cols-2">
                {metricCard('Net Profit', currencyFormatter.format(cashFlow.netProfit))}
                <div className={`rounded-2xl border p-4 shadow-sm ${ui.skyCard}`}>
                  <p className="text-sm">Operating Cash Flow</p>
                  <p className="mt-2 text-2xl font-semibold">{currencyFormatter.format(cashFlow.operatingCashFlow)}</p>
                </div>
              </div>

              <div className={`mt-4 rounded-2xl p-4 shadow-sm ${ui.card}`}>
                <h3 className={`text-sm font-semibold uppercase tracking-[0.16em] ${ui.label}`}>Adjustments</h3>
                {cashFlow.adjustments.length > 0 ? (
                  <div className="mt-3 space-y-2">
                    {cashFlow.adjustments.map((adjustment) => (
                      <div key={adjustment.account} className={`flex flex-wrap items-center justify-between gap-3 rounded-2xl px-4 py-3 text-sm ${ui.listRow}`}>
                        <span className={`font-semibold ${ui.stat}`}>{adjustment.account}</span>
                        <span>
                          Change: {currencyFormatter.format(adjustment.change)} | Impact:{' '}
                          {currencyFormatter.format(adjustment.impact)}
                        </span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className={`mt-3 text-sm ${ui.muted}`}>No asset or liability balance changes were available for working capital adjustments.</p>
                )}
              </div>
            </div>
          ) : null}
        </div>
      </div>
    </div>
  );
}
