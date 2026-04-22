'use client';

import { useCallback, useMemo, useState } from 'react';
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

      setPreviewRows(result.rows.slice(0, 50));
    } catch (error) {
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
    async (event: React.DragEvent<HTMLDivElement>) => {
      event.preventDefault();
      setIsDragging(false);
      const file = event.dataTransfer.files[0];
      if (file) {
        await handleParse(file);
      }
    },
    [handleParse],
  );

  const handleDragOver = useCallback((event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback(() => {
    setIsDragging(false);
  }, []);

  const previewColumns = useMemo(() => requiredHeaders, []);

  return (
    <div className="min-h-screen bg-slate-50 py-10 px-4 text-slate-900">
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
        </div>
      </div>
    </div>
  );
}
