'use client';

import { useCallback, useMemo, useState } from 'react';
import { analyzeLedger, type LedgerAnalysis } from '../../src/lib/analyzeLedger';
import { generateBalanceSheet, type BalanceSheet } from '../../src/lib/generateBalanceSheet';
import { generateCashFlow, type CashFlowStatement } from '../../src/lib/generateCashFlow';
import { generatePL, type ProfitAndLoss } from '../../src/lib/generatePL';
import { FileDropzone } from './components/FileDropzone';
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

export default function UploadPage() {
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
      return <p className="mt-3 text-sm text-slate-600">No accounts available.</p>;
    }

    return (
      <div className="mt-3 space-y-2">
        {entries.map((entry) => (
          <div
            key={entry.account}
            className="flex flex-wrap items-center justify-between gap-3 rounded-2xl bg-slate-50 px-4 py-3 text-sm text-slate-700"
          >
            <span className="font-semibold text-slate-900">{entry.account}</span>
            <span>{currencyFormatter.format(entry.value)}</span>
          </div>
        ))}
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-slate-50 px-4 py-10 text-slate-900">
      <div className="mx-auto w-full max-w-6xl rounded-3xl border border-slate-200 bg-white p-8 shadow-sm">
        <div className="space-y-4">
          <FileDropzone
            fileName={fileName}
            isDragging={isDragging}
            onFileChange={handleFileChange}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          />

          <div className="grid gap-4 lg:grid-cols-[1fr_320px]">
            <div className="space-y-4">
              <div className="rounded-3xl border border-slate-200 bg-slate-50 p-5">
                <h2 className="text-lg font-semibold text-slate-900">Required Headers</h2>
                <div className="mt-4 grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
                  {requiredHeaders.map((header) => (
                    <div key={header} className="rounded-2xl bg-white px-4 py-3 text-sm text-slate-700 shadow-sm">
                      {header}
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <ValidationPanel headerErrors={headerErrors} uploadError={uploadError} rowIssues={rowIssues} />
          </div>

          <PreviewTable columns={previewColumns} rows={previewRows} rowIssues={rowIssues} />

          {analysis ? (
            <div className="rounded-3xl border border-slate-200 bg-slate-50 p-5">
              <div className="flex flex-wrap items-start justify-between gap-4">
                <div>
                  <h2 className="text-lg font-semibold text-slate-900">Ledger Analysis</h2>
                  <p className="mt-1 text-sm text-slate-600">Summary of the parsed transaction set.</p>
                </div>
                <div className="text-sm text-slate-500">Total transactions: {analysis.totalTransactions}</div>
              </div>

              <div className="mt-4 grid gap-4 sm:grid-cols-3">
                <div className="rounded-2xl bg-white p-4 shadow-sm">
                  <p className="text-sm text-slate-500">Total transactions</p>
                  <p className="mt-2 text-2xl font-semibold text-slate-900">{analysis.totalTransactions}</p>
                </div>
                <div className="rounded-2xl bg-white p-4 shadow-sm">
                  <p className="text-sm text-slate-500">Inconsistent vendors</p>
                  <p className="mt-2 text-2xl font-semibold text-slate-900">{analysis.inconsistentVendors.length}</p>
                </div>
                <div className="rounded-2xl bg-white p-4 shadow-sm">
                  <p className="text-sm text-slate-500">Duplicate count</p>
                  <p className="mt-2 text-2xl font-semibold text-slate-900">{analysis.duplicates.length}</p>
                </div>
              </div>

              <div className="mt-4 rounded-2xl bg-white p-4 shadow-sm">
                <h3 className="text-sm font-semibold uppercase tracking-[0.16em] text-slate-500">Inconsistent Vendors</h3>
                {analysis.inconsistentVendors.length > 0 ? (
                  <div className="mt-3 space-y-2">
                    {analysis.inconsistentVendors.map(({ vendor, accounts }) => (
                      <div key={vendor} className="rounded-2xl bg-amber-50 px-4 py-3 text-sm text-amber-900">
                        <span className="font-semibold">{vendor}</span>: {accounts.join(', ')}
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="mt-3 text-sm text-slate-600">No vendors were found with multiple distribution accounts.</p>
                )}
              </div>
            </div>
          ) : null}

          {profitAndLoss ? (
            <div className="rounded-3xl border border-slate-200 bg-slate-50 p-5">
              <div className="flex flex-wrap items-start justify-between gap-4">
                <div>
                  <h2 className="text-lg font-semibold text-slate-900">Profit &amp; Loss</h2>
                  <p className="mt-1 text-sm text-slate-600">Revenue and expense totals generated from ledger account types.</p>
                </div>
                <div className="rounded-full bg-emerald-100 px-4 py-2 text-sm font-semibold text-emerald-800">
                  Net Profit: {currencyFormatter.format(profitAndLoss.netProfit)}
                </div>
              </div>

              <div className="mt-4 grid gap-4 sm:grid-cols-3">
                <div className="rounded-2xl bg-white p-4 shadow-sm">
                  <p className="text-sm text-slate-500">Total Revenue</p>
                  <p className="mt-2 text-2xl font-semibold text-slate-900">
                    {currencyFormatter.format(profitAndLoss.totalRevenue)}
                  </p>
                </div>
                <div className="rounded-2xl bg-white p-4 shadow-sm">
                  <p className="text-sm text-slate-500">Total Expenses</p>
                  <p className="mt-2 text-2xl font-semibold text-slate-900">
                    {currencyFormatter.format(profitAndLoss.totalExpenses)}
                  </p>
                </div>
                <div className="rounded-2xl border border-emerald-200 bg-emerald-50 p-4 shadow-sm">
                  <p className="text-sm text-emerald-700">Net Profit</p>
                  <p className="mt-2 text-2xl font-semibold text-emerald-900">
                    {currencyFormatter.format(profitAndLoss.netProfit)}
                  </p>
                </div>
              </div>

              <div className="mt-4 rounded-2xl bg-white p-4 shadow-sm">
                <h3 className="text-sm font-semibold uppercase tracking-[0.16em] text-slate-500">Monthly Breakdown</h3>
                {Object.keys(profitAndLoss.monthlyBreakdown).length > 0 ? (
                  <div className="mt-3 space-y-2">
                    {Object.entries(profitAndLoss.monthlyBreakdown).map(([month, summary]) => (
                      <div
                        key={month}
                        className="flex flex-wrap items-center justify-between gap-3 rounded-2xl bg-slate-50 px-4 py-3 text-sm text-slate-700"
                      >
                        <span className="font-semibold text-slate-900">{month}</span>
                        <span>
                          Revenue: {currencyFormatter.format(summary.revenue)} | Expenses:{' '}
                          {currencyFormatter.format(summary.expenses)}
                        </span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="mt-3 text-sm text-slate-600">
                    No income or expense transactions with valid dates were available for monthly summaries.
                  </p>
                )}
              </div>
            </div>
          ) : null}

          {balanceSheet ? (
            <div className="rounded-3xl border border-slate-200 bg-slate-50 p-5">
              <div className="flex flex-wrap items-start justify-between gap-4">
                <div>
                  <h2 className="text-lg font-semibold text-slate-900">Balance Sheet</h2>
                  <p className="mt-1 text-sm text-slate-600">
                    Latest balance per account grouped into assets, liabilities, and equity.
                  </p>
                </div>
                <div
                  className={`rounded-full px-4 py-2 text-sm font-semibold ${
                    balanceSheet.isBalanced ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
                  }`}
                >
                  {balanceSheet.isBalanced ? 'Balanced \u2705' : 'Not Balanced \u274C'}
                </div>
              </div>

              <div className="mt-4 grid gap-4 lg:grid-cols-3">
                <div className="rounded-2xl bg-white p-4 shadow-sm">
                  <h3 className="text-sm font-semibold uppercase tracking-[0.16em] text-slate-500">Assets</h3>
                  {renderBalanceSheetEntries(balanceSheet.assets)}
                </div>
                <div className="rounded-2xl bg-white p-4 shadow-sm">
                  <h3 className="text-sm font-semibold uppercase tracking-[0.16em] text-slate-500">Liabilities</h3>
                  {renderBalanceSheetEntries(balanceSheet.liabilities)}
                </div>
                <div className="rounded-2xl bg-white p-4 shadow-sm">
                  <h3 className="text-sm font-semibold uppercase tracking-[0.16em] text-slate-500">Equity</h3>
                  {renderBalanceSheetEntries(balanceSheet.equity)}
                </div>
              </div>

              <div className="mt-4 grid gap-4 sm:grid-cols-3">
                <div className="rounded-2xl bg-white p-4 shadow-sm">
                  <p className="text-sm text-slate-500">Assets Total</p>
                  <p className="mt-2 text-2xl font-semibold text-slate-900">
                    {currencyFormatter.format(balanceSheet.totals.assetsTotal)}
                  </p>
                </div>
                <div className="rounded-2xl bg-white p-4 shadow-sm">
                  <p className="text-sm text-slate-500">Liabilities Total</p>
                  <p className="mt-2 text-2xl font-semibold text-slate-900">
                    {currencyFormatter.format(balanceSheet.totals.liabilitiesTotal)}
                  </p>
                </div>
                <div className="rounded-2xl bg-white p-4 shadow-sm">
                  <p className="text-sm text-slate-500">Equity Total</p>
                  <p className="mt-2 text-2xl font-semibold text-slate-900">
                    {currencyFormatter.format(balanceSheet.totals.equityTotal)}
                  </p>
                </div>
              </div>
            </div>
          ) : null}

          {cashFlow ? (
            <div className="rounded-3xl border border-slate-200 bg-slate-50 p-5">
              <div className="flex flex-wrap items-start justify-between gap-4">
                <div>
                  <h2 className="text-lg font-semibold text-slate-900">Cash Flow Statement</h2>
                  <p className="mt-1 text-sm text-slate-600">
                    Indirect method using net profit and working capital changes.
                  </p>
                </div>
                <div className="rounded-full bg-sky-100 px-4 py-2 text-sm font-semibold text-sky-800">
                  Operating Cash Flow: {currencyFormatter.format(cashFlow.operatingCashFlow)}
                </div>
              </div>

              <div className="mt-4 grid gap-4 sm:grid-cols-2">
                <div className="rounded-2xl bg-white p-4 shadow-sm">
                  <p className="text-sm text-slate-500">Net Profit</p>
                  <p className="mt-2 text-2xl font-semibold text-slate-900">
                    {currencyFormatter.format(cashFlow.netProfit)}
                  </p>
                </div>
                <div className="rounded-2xl border border-sky-200 bg-sky-50 p-4 shadow-sm">
                  <p className="text-sm text-sky-700">Operating Cash Flow</p>
                  <p className="mt-2 text-2xl font-semibold text-sky-900">
                    {currencyFormatter.format(cashFlow.operatingCashFlow)}
                  </p>
                </div>
              </div>

              <div className="mt-4 rounded-2xl bg-white p-4 shadow-sm">
                <h3 className="text-sm font-semibold uppercase tracking-[0.16em] text-slate-500">Adjustments</h3>
                {cashFlow.adjustments.length > 0 ? (
                  <div className="mt-3 space-y-2">
                    {cashFlow.adjustments.map((adjustment) => (
                      <div
                        key={adjustment.account}
                        className="flex flex-wrap items-center justify-between gap-3 rounded-2xl bg-slate-50 px-4 py-3 text-sm text-slate-700"
                      >
                        <span className="font-semibold text-slate-900">{adjustment.account}</span>
                        <span>
                          Change: {currencyFormatter.format(adjustment.change)} | Impact:{' '}
                          {currencyFormatter.format(adjustment.impact)}
                        </span>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="mt-3 text-sm text-slate-600">
                    No asset or liability balance changes were available for working capital adjustments.
                  </p>
                )}
              </div>
            </div>
          ) : null}
        </div>
      </div>
    </div>
  );
}
