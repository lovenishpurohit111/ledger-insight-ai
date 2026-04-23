'use client';

import { type HeaderKey, type LedgerRow, type RowIssue } from '../upload-utils';
import { type UploadTheme } from './FileDropzone';

type PreviewTableProps = {
  columns: HeaderKey[];
  rows: LedgerRow[];
  rowIssues: RowIssue[];
  theme: UploadTheme;
};

export function PreviewTable({ columns, rows, rowIssues, theme }: PreviewTableProps) {
  const issueMap = new Map<number, string[]>(rowIssues.map((issue) => [issue.rowIndex, issue.issues]));
  const isDark = theme === 'dark';
  const panelClass = isDark ? 'border-slate-700 bg-slate-900' : 'border-slate-200 bg-white';
  const titleClass = isDark ? 'text-white' : 'text-slate-900';
  const mutedClass = isDark ? 'text-slate-300' : 'text-slate-600';
  const subtleClass = isDark ? 'text-slate-400' : 'text-slate-500';
  const headerClass = isDark
    ? 'border-slate-700 bg-slate-950 text-slate-300'
    : 'border-slate-200 bg-slate-50 text-slate-500';
  const rowClass = isDark ? 'bg-slate-900' : 'bg-white';
  const errorRowClass = isDark ? 'bg-rose-950/40' : 'bg-rose-50';
  const cellClass = isDark ? 'border-slate-800 text-slate-300' : 'border-slate-200 text-slate-700';
  const missingClass = isDark ? 'text-rose-300' : 'text-rose-700';

  return (
    <div className={`rounded-3xl border p-5 shadow-sm ${panelClass}`}>
      <div className="flex items-center justify-between gap-4">
        <div>
          <h2 className={`text-lg font-semibold ${titleClass}`}>Preview</h2>
          <p className={`mt-1 text-sm ${mutedClass}`}>Showing up to 50 rows from the uploaded file.</p>
        </div>
        <div className={`text-sm ${subtleClass}`}>Rows: {rows.length}</div>
      </div>
      <div className="mt-5 overflow-x-auto">
        <table className="min-w-full border-separate border-spacing-0 text-left text-sm">
          <thead>
            <tr>
              {columns.map((column) => (
                <th
                  key={column}
                  className={`sticky top-0 z-10 border-b px-3 py-3 text-xs font-semibold uppercase tracking-[0.16em] ${headerClass}`}
                >
                  {column}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td className={`px-3 py-6 text-sm ${subtleClass}`} colSpan={columns.length}>
                  No preview available. Upload a valid CSV or Excel file to see the first rows.
                </td>
              </tr>
            ) : (
              rows.map((row, index) => {
                const issues = issueMap.get(index);
                const hasError = Array.isArray(issues) && issues.length > 0;

                return (
                  <tr key={`${row.Name}-${index}`} className={hasError ? errorRowClass : rowClass}>
                    {columns.map((column) => (
                      <td
                        key={`${column}-${index}`}
                        className={`whitespace-nowrap border-b px-3 py-3 ${cellClass} ${
                          hasError && !row[column] ? missingClass : ''
                        }`}
                      >
                        {row[column] || '-'}
                      </td>
                    ))}
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
