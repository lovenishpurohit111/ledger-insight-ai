'use client';

import { type ChangeEvent, type DragEvent, useCallback, useMemo, useState } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';

type HeaderKey =
  | 'Distribution account'
  | 'Distribution account type'
  | 'Transaction date'
  | 'Transaction type'
  | 'Num'
  | 'Name'
  | 'Description'
  | 'Split'
  | 'Amount'
  | 'Balance';

type LedgerRow = Record<HeaderKey, string>;

type ParseResult = {
  rows: LedgerRow[];
  errors: string[];
};

const requiredHeaders: HeaderKey[] = [
  'Distribution account',
  'Distribution account type',
  'Transaction date',
  'Transaction type',
  'Num',
  'Name',
  'Description',
  'Split',
  'Amount',
  'Balance',
];

const normalizeHeader = (value: string) => value.trim();

const validateHeaders = (headers: string[]) => {
  const trimmed = headers.map(normalizeHeader);
  const missing = requiredHeaders.filter((required) => !trimmed.includes(required));
  const errors: string[] = [];
  if (missing.length > 0) {
    errors.push(`Missing required headers: ${missing.join(', ')}`);
  }
  if (!trimmed.includes('Distribution account type')) {
    errors.push('Distribution account type header is required and must be present.');
  }
  return { errors, headers: trimmed };
};

const buildLedgerRows = (rawRows: Array<Record<string, unknown>>) => {
  return rawRows.map((rawRow, rowIndex) => {
    const row = requiredHeaders.reduce((acc, header) => {
      const value = rawRow[header] ?? rawRow[header.trim()] ?? '';
      acc[header] = String(value).trim();
      return acc;
    }, {} as LedgerRow);

    return row;
  });
};

const hasRowError = (row: LedgerRow) => {
  return requiredHeaders.some((header) => {
    const value = row[header];
    return value === '' || value === null || value === undefined;
  });
};

const parseCsvFile = async (file: File): Promise<ParseResult> => {
  return new Promise((resolve, reject) => {
    Papa.parse<Record<string, string>>(file, {
      header: true,
      skipEmptyLines: true,
      transformHeader: normalizeHeader,
      transform: (value) => value.trim(),
      complete: (result) => {
        const fields = result.meta.fields ?? [];
        const { errors: headerErrors } = validateHeaders(fields);
        if (headerErrors.length > 0) {
          resolve({ rows: [], errors: headerErrors });
          return;
        }

        const rows = buildLedgerRows(result.data);
        resolve({ rows, errors: [] });
      },
      error: (error) => {
        reject(error);
      },
    });
  });
};

const parseXlsxFile = async (file: File): Promise<ParseResult> => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    return { rows: [], errors: ['Excel file contains no sheets.'] };
  }
  const sheet = workbook.Sheets[firstSheetName];
  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' });
  if (rawRows.length === 0) {
    return { rows: [], errors: ['Excel sheet is empty.'] };
  }

  const headerRow = rawRows[0] as Array<string>;
  const normalizedHeaders = headerRow.map((header) => normalizeHeader(String(header)));
  const { errors: headerErrors } = validateHeaders(normalizedHeaders);
  if (headerErrors.length > 0) {
    return { rows: [], errors: headerErrors };
  }

  const rows = rawRows.slice(1).map((rowArray, rowIndex) => {
    const rawRow = requiredHeaders.reduce((acc, header) => {
      const index = normalizedHeaders.indexOf(header);
      const rawValue = Array.isArray(rowArray) ? rowArray[index] ?? '' : '';
      acc[header] = String(rawValue).trim();
      return acc;
    }, {} as LedgerRow);
    return rawRow;
  });

  return { rows, errors: [] };
};

const isCsvFile = (file: File) => file.type === 'text/csv' || file.name.toLowerCase().endsWith('.csv');
const isExcelFile = (file: File) =>
  file.type.includes('spreadsheet') || file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls');

export default function UploadPage() {
  const [previewRows, setPreviewRows] = useState<LedgerRow[]>([]);
  const [errors, setErrors] = useState<string[]>([]);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string>('');
  const [isDragging, setIsDragging] = useState(false);

  const handleParse = useCallback(async (file: File) => {
    setUploadError(null);
    setErrors([]);
    setPreviewRows([]);
    setFileName(file.name);

    if (!isCsvFile(file) && !isExcelFile(file)) {
      setUploadError('Only CSV and XLSX files are accepted.');
      return;
    }

    try {
      const result = isCsvFile(file) ? await parseCsvFile(file) : await parseXlsxFile(file);
      if (result.errors.length > 0) {
        setErrors(result.errors);
      }
      setPreviewRows(result.rows.slice(0, 50));
    } catch (error) {
      setUploadError('Unable to parse file. Please verify the file format.');
    }
  }, []);

  const handleDrop = useCallback(
    async (event: DragEvent<HTMLDivElement>) => {
      event.preventDefault();
      setIsDragging(false);
      const file = event.dataTransfer.files[0];
      if (file) {
        await handleParse(file);
      }
    },
    [handleParse],
  );

  const handleFileChange = useCallback(
    async (event: ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (file) {
        await handleParse(file);
      }
    },
    [handleParse],
  );

  const previewColumns = useMemo(() => requiredHeaders, []);

  return (
    <div className="min-h-screen bg-slate-50 py-10 px-4 text-slate-900">
      <div className="mx-auto w-full max-w-6xl rounded-3xl border border-slate-200 bg-white p-8 shadow-sm">
        <div className="space-y-4">
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
              onDragOver={(event) => {
                event.preventDefault();
                setIsDragging(true);
              }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={handleDrop}
              className={`group relative flex min-h-65 cursor-pointer flex-col items-center justify-center gap-4 rounded-3xl border-2 border-dashed px-6 py-10 text-center transition-all ${
                isDragging ? 'border-slate-900 bg-slate-950/5' : 'border-slate-300 bg-slate-50'
              }`}
            >
              <input
                id="ledger-upload"
                type="file"
                accept=".csv,.xlsx,.xls"
                className="hidden"
                onChange={handleFileChange}
              />
              <div className="space-y-3">
                <div className="mx-auto flex h-16 w-16 items-center justify-center rounded-full bg-slate-900 text-white">
                  <span className="text-2xl">⇪</span>
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
          </div>

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
              {fileName ? (
                <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
                  <p className="text-sm text-slate-500">Uploaded file</p>
                  <p className="mt-2 text-base font-medium text-slate-900">{fileName}</p>
                </div>
              ) : null}
            </div>

            <div className="space-y-4">
              <div className="rounded-3xl border border-slate-200 bg-slate-50 p-5">
                <h2 className="text-lg font-semibold text-slate-900">Validation</h2>
                <p className="mt-2 text-sm text-slate-600">
                  The upload will stop if required headers are missing. Errors appear below.
                </p>
                <div className="mt-4 space-y-2 text-sm text-slate-700">
                  {errors.length === 0 ? (
                    <p className="rounded-2xl bg-green-50 px-3 py-2 text-green-700">No header errors detected yet.</p>
                  ) : (
                    errors.map((errorText) => (
                      <p key={errorText} className="rounded-2xl bg-rose-50 px-3 py-2 text-rose-700">
                        {errorText}
                      </p>
                    ))
                  )}
                  {uploadError ? (
                    <p className="rounded-2xl bg-rose-50 px-3 py-2 text-rose-700">{uploadError}</p>
                  ) : null}
                </div>
              </div>
            </div>
          </div>

          <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
            <div className="flex items-center justify-between gap-4">
              <div>
                <h2 className="text-lg font-semibold text-slate-900">Preview</h2>
                <p className="mt-1 text-sm text-slate-600">Showing up to 50 rows from the uploaded file.</p>
              </div>
              <div className="text-sm text-slate-500">Rows: {previewRows.length}</div>
            </div>
            <div className="mt-5 overflow-x-auto">
              <table className="min-w-full border-separate border-spacing-0 text-left text-sm">
                <thead>
                  <tr>
                    {previewColumns.map((column) => (
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
                  {previewRows.length === 0 ? (
                    <tr>
                      <td className="px-3 py-6 text-sm text-slate-500" colSpan={previewColumns.length}>
                        No preview available. Upload a valid CSV or Excel file to see the first rows.
                      </td>
                    </tr>
                  ) : (
                    previewRows.map((row, index) => {
                      const rowError = hasRowError(row);
                      return (
                        <tr key={`${row.Name}-${index}`} className={rowError ? 'bg-rose-50' : 'bg-white'}>
                          {previewColumns.map((column) => (
                            <td
                              key={`${column}-${index}`}
                              className={`whitespace-nowrap border-b border-slate-200 px-3 py-3 text-slate-700 ${
                                rowError && row[column] === '' ? 'text-rose-700' : ''
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
        </div>
      </div>
    </div>
  );
}
