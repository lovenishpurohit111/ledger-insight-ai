'use client';

import { type HeaderKey, type LedgerRow, type RowIssue } from '../upload-utils';

type PreviewTableProps = {
  columns: HeaderKey[];
  rows: LedgerRow[];
  rowIssues: RowIssue[];
};

export function PreviewTable({ columns, rows, rowIssues }: PreviewTableProps) {
  const issueMap = new Map<number, string[]>(rowIssues.map((issue) => [issue.rowIndex, issue.issues]));

  return (
    <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
      <div className="flex items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-semibold text-slate-900">Preview</h2>
          <p className="mt-1 text-sm text-slate-600">Showing up to 50 rows from the uploaded file.</p>
        </div>
        <div className="text-sm text-slate-500">Rows: {rows.length}</div>
      </div>
      <div className="mt-5 overflow-x-auto">
        <table className="min-w-full border-separate border-spacing-0 text-left text-sm">
          <thead>
            <tr>
              {columns.map((column) => (
                <th
                  key={column}
                  className="sticky top-0 z-10 border-b border-slate-200 bg-slate-50 px-3 py-3 text-xs font-semibold uppercase tracking-[0.16em] text-slate-500"
                >
                  {column}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td className="px-3 py-6 text-sm text-slate-500" colSpan={columns.length}>
                  No preview available. Upload a valid CSV or Excel file to see the first rows.
                </td>
              </tr>
            ) : (
              rows.map((row, index) => {
                const issues = issueMap.get(index);
                const hasError = Array.isArray(issues) && issues.length > 0;

                return (
                  <tr key={`${row.Name}-${index}`} className={hasError ? 'bg-rose-50' : 'bg-white'}>
                    {columns.map((column) => (
                      <td
                        key={`${column}-${index}`}
                        className={`whitespace-nowrap border-b border-slate-200 px-3 py-3 text-slate-700 ${
                          hasError && !row[column] ? 'text-rose-700' : ''
                        }`}
                      >
                        {row[column] || '—'}
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
