'use client';

import { type ChangeEvent, type DragEvent } from 'react';

export type UploadTheme = 'dark' | 'light';

type FileDropzoneProps = {
  fileName: string;
  isDragging: boolean;
  theme: UploadTheme;
  onFileChange: (event: ChangeEvent<HTMLInputElement>) => void;
  onDragOver: (event: DragEvent<HTMLLabelElement>) => void;
  onDragLeave: () => void;
  onDrop: (event: DragEvent<HTMLLabelElement>) => void;
};

export function FileDropzone({
  fileName,
  isDragging,
  theme,
  onFileChange,
  onDragOver,
  onDragLeave,
  onDrop,
}: FileDropzoneProps) {
  const isDark = theme === 'dark';
  const panelClass = isDark
    ? 'border-slate-700 bg-slate-900 text-slate-100'
    : 'border-slate-200 bg-white text-slate-900';
  const eyebrowClass = isDark ? 'text-cyan-300' : 'text-cyan-700';
  const titleClass = isDark ? 'text-white' : 'text-slate-950';
  const mutedClass = isDark ? 'text-slate-300' : 'text-slate-600';
  const dropzoneClass = isDragging
    ? isDark
      ? 'border-cyan-300 bg-cyan-950/40'
      : 'border-cyan-600 bg-cyan-50'
    : isDark
      ? 'border-slate-600 bg-slate-950/50'
      : 'border-slate-300 bg-slate-50';
  const iconClass = isDark ? 'bg-cyan-400 text-slate-950' : 'bg-slate-900 text-white';
  const buttonClass = isDark
    ? 'bg-slate-800 text-slate-200 group-hover:bg-slate-700'
    : 'bg-slate-100 text-slate-600 group-hover:bg-slate-200';
  const uploadedClass = isDark
    ? 'border-slate-700 bg-slate-950/50'
    : 'border-slate-200 bg-slate-50';

  return (
    <div className={`flex flex-col gap-3 rounded-3xl border p-6 shadow-sm ${panelClass}`}>
      <div>
        <p className={`text-sm uppercase tracking-[0.24em] ${eyebrowClass}`}>General Ledger Upload</p>
        <h1 className={`mt-2 text-3xl font-semibold tracking-tight ${titleClass}`}>Upload CSV or XLSX</h1>
        <p className={`mt-2 max-w-2xl ${mutedClass}`}>
          Drag and drop a ledger file or click to select. Required headers are enforced strictly.
        </p>
      </div>
      <label
        htmlFor="ledger-upload"
        onDragOver={onDragOver}
        onDragLeave={onDragLeave}
        onDrop={onDrop}
        className={`group relative flex min-h-65 cursor-pointer flex-col items-center justify-center gap-4 rounded-3xl border-2 border-dashed px-6 py-10 text-center transition-all ${dropzoneClass}`}
      >
        <input
          id="ledger-upload"
          type="file"
          accept=".csv,.xlsx,.xls"
          className="hidden"
          onChange={onFileChange}
        />
        <div className="space-y-3">
          <div className={`mx-auto flex h-16 w-16 items-center justify-center rounded-full ${iconClass}`}>
            <span className="text-2xl" aria-hidden="true">
              {'\u21EA'}
            </span>
          </div>
          <div>
            <p className={`text-lg font-semibold ${titleClass}`}>Drop your file here</p>
            <p className={`text-sm ${isDark ? 'text-slate-400' : 'text-slate-500'}`}>CSV or XLSX with required ledger headers</p>
          </div>
        </div>
        <span className={`mt-2 inline-flex rounded-full px-3 py-1 text-sm transition-colors ${buttonClass}`}>
          Browse files
        </span>
      </label>
      {fileName ? (
        <div className={`rounded-3xl border p-5 shadow-sm ${uploadedClass}`}>
          <p className={`text-sm ${isDark ? 'text-slate-400' : 'text-slate-500'}`}>Uploaded file</p>
          <p className={`mt-2 text-base font-medium ${isDark ? 'text-slate-100' : 'text-slate-900'}`}>{fileName}</p>
        </div>
      ) : null}
    </div>
  );
}
