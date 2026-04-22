import Papa from 'papaparse';
import * as XLSX from 'xlsx';

export type HeaderKey =
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

export type LedgerRow = Record<HeaderKey, string>;

export type RowIssue = {
  rowIndex: number;
  issues: string[];
};

export type ParseResult = {
  rows: LedgerRow[];
  headerErrors: string[];
  rowIssues: RowIssue[];
};

export const requiredHeaders: HeaderKey[] = [
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

const requiredRowFields: HeaderKey[] = [
  'Distribution account',
  'Distribution account type',
  'Transaction date',
  'Transaction type',
];

const normalizeHeader = (value: string) => value.trim();

const hasAnyValue = (row: Partial<Record<HeaderKey, string>>) =>
  requiredHeaders.some((header) => Boolean(row[header]?.trim()));

const isNonTransactionRow = (row: Partial<Record<HeaderKey, string>>) =>
  (!row['Transaction date']?.trim() && !row['Transaction type']?.trim()) ||
  (!row.Amount?.trim() && !row.Balance?.trim());

export const validateHeaders = (headers: string[]) => {
  const trimmed = headers.map(normalizeHeader);
  const missing = requiredHeaders.filter((expected) => !trimmed.includes(expected));
  const errors: string[] = [];

  if (missing.length > 0) {
    errors.push(`Missing required headers: ${missing.join(', ')}`);
  }

  if (!trimmed.includes('Distribution account type')) {
    errors.push('Distribution account type header is required and must be present.');
  }

  return { errors, headers: trimmed };
};

export const validateLedgerRow = (row: LedgerRow, rowIndex: number): RowIssue | null => {
  const issues: string[] = [];

  requiredRowFields.forEach((header) => {
    if (!row[header]?.trim()) {
      issues.push(`${header} is required.`);
    }
  });

  if (row['Transaction date']?.trim()) {
    const parsed = Date.parse(row['Transaction date']);
    if (Number.isNaN(parsed)) {
      issues.push('Transaction date must be valid.');
    }
  }

  if (!row.Amount?.trim() && !row.Balance?.trim()) {
    issues.push('Amount or Balance is required.');
  }

  const numericFields: Array<'Amount' | 'Balance'> = ['Amount', 'Balance'];
  numericFields.forEach((field) => {
    const value = row[field]?.trim();
    if (value && !/^[-+]?\d{1,3}(?:,\d{3})*(?:\.\d+)?$/.test(value) && !/^[-+]?\d*(?:\.\d+)?$/.test(value)) {
      issues.push(`${field} must be numeric.`);
    }
  });

  return issues.length > 0 ? { rowIndex, issues } : null;
};

const buildLedgerRows = (rawRows: Array<Record<string, unknown>>) => {
  return rawRows.map((rawRow) => {
    return requiredHeaders.reduce((acc, header) => {
      const rawValue = rawRow[header] ?? rawRow[header.trim()] ?? '';
      acc[header] = String(rawValue).trim();
      return acc;
    }, {} as LedgerRow);
  }).filter((row) => hasAnyValue(row) && !isNonTransactionRow(row));
};

export const isCsvFile = (file: File) => file.type === 'text/csv' || file.name.toLowerCase().endsWith('.csv');

export const isExcelFile = (file: File) =>
  file.type.includes('spreadsheet') || file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls');

export const parseCsvFile = async (file: File): Promise<ParseResult> => {
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
          resolve({ rows: [], headerErrors, rowIssues: [] });
          return;
        }

        const rows = buildLedgerRows(result.data);
        const rowIssues = rows
          .map((row, index) => validateLedgerRow(row, index))
          .filter((issue): issue is RowIssue => issue !== null);

        resolve({ rows, headerErrors: [], rowIssues });
      },
      error: (error) => {
        reject(error);
      },
    });
  });
};

export const parseXlsxFile = async (file: File): Promise<ParseResult> => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    return { rows: [], headerErrors: ['Excel file contains no sheets.'], rowIssues: [] };
  }

  const sheet = workbook.Sheets[firstSheetName];
  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '' }) as unknown[][];
  if (rawRows.length === 0) {
    return { rows: [], headerErrors: ['Excel sheet is empty.'], rowIssues: [] };
  }

  const headerRow = rawRows[0] as Array<string>;
  const normalizedHeaders = headerRow.map((header) => normalizeHeader(String(header)));
  const { errors: headerErrors } = validateHeaders(normalizedHeaders);
  if (headerErrors.length > 0) {
    return { rows: [], headerErrors, rowIssues: [] };
  }

  const rows = rawRows
    .slice(1)
    .map((rowArray) => {
      return requiredHeaders.reduce((acc, header) => {
        const index = normalizedHeaders.indexOf(header);
        const rawValue = Array.isArray(rowArray) ? rowArray[index] ?? '' : '';
        acc[header] = String(rawValue).trim();
        return acc;
      }, {} as LedgerRow);
    })
    .filter((row) => hasAnyValue(row) && !isNonTransactionRow(row));

  const rowIssues = rows
    .map((row, index) => validateLedgerRow(row, index))
    .filter((issue): issue is RowIssue => issue !== null);

  return { rows, headerErrors: [], rowIssues };
};
