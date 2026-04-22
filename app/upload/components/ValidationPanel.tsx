'use client';

import { type RowIssue } from '../upload-utils';

type ValidationPanelProps = {
  headerErrors: string[];
  uploadError: string | null;
  rowIssues: RowIssue[];
};

export function ValidationPanel({ headerErrors, uploadError, rowIssues }: ValidationPanelProps) {
  return (
    <div className="rounded-3xl border border-slate-200 bg-slate-50 p-5">
      <h2 className="text-lg font-semibold text-slate-900">Validation</h2>
      <p className="mt-2 text-sm text-slate-600">The upload will stop if required headers are missing. Errors appear below.</p>
      <div className="mt-4 space-y-2 text-sm text-slate-700">
        {uploadError ? (
          <p className="rounded-2xl bg-rose-50 px-3 py-2 text-rose-700">{uploadError}</p>
        ) : null}
        {headerErrors.length > 0 ? (
          headerErrors.map((errorText) => (
            <p key={errorText} className="rounded-2xl bg-rose-50 px-3 py-2 text-rose-700">
              {errorText}
            </p>
          ))
        ) : rowIssues.length > 0 ? (
          rowIssues.map((issue) => (
            <div key={`row-${issue.rowIndex}`} className="rounded-2xl bg-amber-50 px-3 py-2 text-amber-800">
              <p className="font-semibold">Row {issue.rowIndex + 2} issues</p>
              <ul className="mt-2 list-disc space-y-1 pl-4 text-sm">
                {issue.issues.map((issueText) => (
                  <li key={issueText}>{issueText}</li>
                ))}
              </ul>
            </div>
          ))
        ) : (
          <p className="rounded-2xl bg-emerald-50 px-3 py-2 text-emerald-700">No header or row validation issues found.</p>
        )}
      </div>
    </div>
  );
}
