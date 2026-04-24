'use client';

import { CORE_MANDATORY, type RowIssue } from '../upload-utils';
import { type UploadTheme } from './FileDropzone';

type Props = {
  headerErrors: string[];
  uploadError: string | null;
  rowIssues: RowIssue[];
  theme: UploadTheme;
};

export function ValidationPanel({ headerErrors, uploadError, rowIssues, theme }: Props) {
  const d = theme === 'dark';
  const panel   = d ? 'border-rose-800 bg-rose-950/40'   : 'border-rose-200 bg-rose-50';
  const title   = d ? 'text-rose-200'  : 'text-rose-800';
  const error   = d ? 'bg-rose-900/60 text-rose-200 border border-rose-700' : 'bg-rose-100 text-rose-700 border border-rose-200';
  const warn    = d ? 'bg-amber-950/50 text-amber-200' : 'bg-amber-50 text-amber-800';
  const muted   = d ? 'text-rose-300'  : 'text-rose-600';
  const badge   = d ? 'bg-rose-800 text-rose-100' : 'bg-rose-200 text-rose-800';

  const isMandatoryError = headerErrors.some((e) =>
    CORE_MANDATORY.some((col) => e.includes(col))
  );

  return (
    <div className={`rounded-3xl border p-5 ${panel}`}>
      <h2 className={`text-base font-bold ${title}`}>
        {headerErrors.length > 0 ? '❌ File Rejected — Missing Required Columns' : '⚠️ Row Validation Issues'}
      </h2>

      {isMandatoryError && (
        <div className={`mt-3 rounded-2xl p-4 text-sm ${error}`}>
          <p className="font-semibold mb-2">This file cannot be processed. The following columns are <strong>mandatory</strong>:</p>
          <div className="flex flex-wrap gap-2 mt-2">
            {CORE_MANDATORY.map((col) => (
              <span key={col} className={`rounded-full px-3 py-1 text-xs font-semibold ${badge}`}>{col}</span>
            ))}
          </div>
          <p className={`mt-3 text-xs ${muted}`}>
            These fields are needed to correctly classify accounts and build the Balance Sheet and P&L.
            QBO exports don't include "Distribution account type" — you'll need to add this column manually
            before uploading, or use a file that already contains it.
          </p>
        </div>
      )}

      {!isMandatoryError && headerErrors.map((e) => (
        <p key={e} className={`mt-2 rounded-2xl px-3 py-2 text-sm ${error}`}>{e}</p>
      ))}

      {uploadError && (
        <p className={`mt-2 rounded-2xl px-3 py-2 text-sm ${error}`}>{uploadError}</p>
      )}

      {rowIssues.length > 0 && headerErrors.length === 0 && (
        <div className="mt-3 space-y-2">
          {rowIssues.slice(0, 20).map((issue) => (
            <div key={`row-${issue.rowIndex}`} className={`rounded-2xl px-3 py-2 text-sm ${warn}`}>
              <p className="font-semibold">Row {issue.rowIndex + 2}</p>
              <ul className="mt-1 list-disc pl-4 space-y-0.5">
                {issue.issues.map((t) => <li key={t}>{t}</li>)}
              </ul>
            </div>
          ))}
          {rowIssues.length > 20 && (
            <p className={`text-xs ${muted}`}>...and {rowIssues.length - 20} more row issues.</p>
          )}
        </div>
      )}
    </div>
  );
}
