'use client';

import { type ChangeEvent, type DragEvent } from 'react';

type FileDropzoneProps = {
  fileName: string;
  isDragging: boolean;
  onFileChange: (event: ChangeEvent<HTMLInputElement>) => void;
  onDragOver: (event: DragEvent<HTMLLabelElement>) => void;
  onDragLeave: () => void;
  onDrop: (event: DragEvent<HTMLLabelElement>) => void;
};

export function FileDropzone({
  fileName,
  isDragging,
  onFileChange,
  onDragOver,
  onDragLeave,
  onDrop,
}: FileDropzoneProps) {
  return (
    <div className="flex flex-col gap-3 rounded-3xl bg-slate-950/5 p-6 text-slate-900 shadow-sm">
      <div>
        <p className="text-sm uppercase tracking-[0.24em] text-slate-500">General Ledger Upload</p>
        <h1 className="mt-2 text-3xl font-semibold tracking-tight text-slate-950">Upload CSV or XLSX</h1>
        <p className="mt-2 max-w-2xl text-slate-600">
          Drag and drop a ledger file or click to select. Required headers are enforced strictly.
        </p>
      </div>
      <label
        htmlFor="ledger-upload"
        onDragOver={onDragOver}
        onDragLeave={onDragLeave}
        onDrop={onDrop}
        className={`group relative flex min-h-65 cursor-pointer flex-col items-center justify-center gap-4 rounded-3xl border-2 border-dashed px-6 py-10 text-center transition-all ${
          isDragging ? 'border-slate-900 bg-slate-950/5' : 'border-slate-300 bg-slate-50'
        }`}
      >
        <input
          id="ledger-upload"
          type="file"
          accept=".csv,.xlsx,.xls"
          className="hidden"
          onChange={onFileChange}
        />
        <div className="space-y-3">
          <div className="mx-auto flex h-16 w-16 items-center justify-center rounded-full bg-slate-900 text-white">
            <span className="text-2xl" aria-hidden="true">
              {'\u21EA'}
            </span>
          </div>
          <div>
            <p className="text-lg font-semibold text-slate-950">Drop your file here</p>
            <p className="text-sm text-slate-500">CSV or XLSX with required ledger headers</p>
          </div>
        </div>
        <span className="mt-2 inline-flex rounded-full bg-slate-100 px-3 py-1 text-sm text-slate-600 transition-colors group-hover:bg-slate-200">
          Browse files
        </span>
      </label>
      {fileName ? (
        <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
          <p className="text-sm text-slate-500">Uploaded file</p>
          <p className="mt-2 text-base font-medium text-slate-900">{fileName}</p>
        </div>
      ) : null}
    </div>
  );
}
