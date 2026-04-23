'use client';

import { type RowIssue } from '../upload-utils';
import { type UploadTheme } from './FileDropzone';

type ValidationPanelProps = {
  headerErrors: string[];
  uploadError: string | null;
  rowIssues: RowIssue[];
  theme: UploadTheme;
};

export function ValidationPanel({ headerErrors, uploadError, rowIssues, theme }: ValidationPanelProps) {
  const isDark = theme === 'dark';
  const panelClass = isDark ? 'border-slate-700 bg-slate-900' : 'border-slate-200 bg-slate-50';
  const titleClass = isDark ? 'text-white' : 'text-slate-900';
  const mutedClass = isDark ? 'text-slate-300' : 'text-slate-600';
  const bodyClass = isDark ? 'text-slate-300' : 'text-slate-700';
  const errorClass = isDark ? 'bg-rose-950/60 text-rose-200' : 'bg-rose-50 text-rose-700';
  const warningClass = isDark ? 'bg-amber-950/50 text-amber-200' : 'bg-amber-50 text-amber-800';
  const successClass = isDark ? 'bg-emerald-950/50 text-emerald-200' : 'bg-emerald-50 text-emerald-700';

  return (
    <div className={`rounded-3xl border p-5 ${panelClass}`}>
      <h2 className={`text-lg font-semibold ${titleClass}`}>Validation</h2>
      <p className={`mt-2 text-sm ${mutedClass}`}>The upload will stop if required headers are missing. Errors appear below.</p>
      <div className={`mt-4 space-y-2 text-sm ${bodyClass}`}>
        {uploadError ? (
          <p className={`rounded-2xl px-3 py-2 ${errorClass}`}>{uploadError}</p>
        ) : null}
        {headerErrors.length > 0 ? (
          headerErrors.map((errorText) => (
            <p key={errorText} className={`rounded-2xl px-3 py-2 ${errorClass}`}>
              {errorText}
            </p>
          ))
        ) : rowIssues.length > 0 ? (
          rowIssues.map((issue) => (
            <div key={`row-${issue.rowIndex}`} className={`rounded-2xl px-3 py-2 ${warningClass}`}>
              <p className="font-semibold">Row {issue.rowIndex + 2} issues</p>
              <ul className="mt-2 list-disc space-y-1 pl-4 text-sm">
                {issue.issues.map((issueText) => (
                  <li key={issueText}>{issueText}</li>
                ))}
              </ul>
            </div>
          ))
        ) : (
          <p className={`rounded-2xl px-3 py-2 ${successClass}`}>No header or row validation issues found.</p>
        )}
      </div>
    </div>
  );
}
