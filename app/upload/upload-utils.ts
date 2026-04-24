import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { parseCurrencyAmount } from '../../src/lib/accounting';

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
export type RowIssue = { rowIndex: number; issues: string[] };
export type ParseResult = { rows: LedgerRow[]; headerErrors: string[]; rowIssues: RowIssue[] };

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

// These 3 are strictly mandatory — file is rejected without them
export const CORE_MANDATORY: HeaderKey[] = [
  'Distribution account',
  'Distribution account type',
  'Transaction type',
];

// These are required per-row
const REQUIRED_ROW_FIELDS: HeaderKey[] = [
  'Distribution account',
  'Distribution account type',
  'Transaction type',
  'Transaction date',
];

const normalizeHeader = (v: string) => v.trim();

const hasAnyValue = (row: Partial<Record<HeaderKey, string>>) =>
  requiredHeaders.some((h) => Boolean(row[h]?.trim()));

const isNonTransactionRow = (row: Partial<Record<HeaderKey, string>>) =>
  (!row['Transaction date']?.trim() && !row['Transaction type']?.trim()) ||
  (!row.Amount?.trim() && !row.Balance?.trim());

// ─── Header validation ───────────────────────────────────────────────────────

export const validateHeaders = (headers: string[]): { errors: string[]; headers: string[] } => {
  const trimmed = headers.map(normalizeHeader);
  const missing = CORE_MANDATORY.filter((h) => !trimmed.includes(h));
  const errors: string[] = [];
  if (missing.length > 0) {
    errors.push(
      `Missing mandatory columns: ${missing.join(', ')}. ` +
      `Your file must include "Distribution account", "Distribution account type", and "Transaction type". ` +
      `These are required to correctly build the Balance Sheet and P&L.`
    );
  }
  return { errors, headers: trimmed };
};

// ─── Row validation ──────────────────────────────────────────────────────────

export const validateLedgerRow = (row: LedgerRow, rowIndex: number): RowIssue | null => {
  const issues: string[] = [];

  REQUIRED_ROW_FIELDS.forEach((h) => {
    if (!row[h]?.trim()) issues.push(`"${h}" is required.`);
  });

  if (row['Transaction date']?.trim() && isNaN(Date.parse(row['Transaction date']))) {
    issues.push('Transaction date must be a valid date.');
  }

  if (!row.Amount?.trim() && !row.Balance?.trim()) {
    issues.push('Amount or Balance is required.');
  }

  (['Amount', 'Balance'] as const).forEach((f) => {
    const v = row[f]?.trim();
    if (v && parseCurrencyAmount(v) === null) issues.push(`${f} must be numeric.`);
  });

  return issues.length > 0 ? { rowIndex, issues } : null;
};

// ─── Row builder ─────────────────────────────────────────────────────────────

const buildRows = (rawRows: Array<Record<string, unknown>>): LedgerRow[] =>
  rawRows
    .map((raw) =>
      requiredHeaders.reduce((acc, h) => {
        acc[h] = String(raw[h] ?? '').trim();
        return acc;
      }, {} as LedgerRow)
    )
    .filter((row) => hasAnyValue(row) && !isNonTransactionRow(row));

export const isCsvFile = (f: File) => f.type === 'text/csv' || f.name.toLowerCase().endsWith('.csv');
export const isExcelFile = (f: File) =>
  f.type.includes('spreadsheet') ||
  f.name.toLowerCase().endsWith('.xlsx') ||
  f.name.toLowerCase().endsWith('.xls');

// ─── CSV parser ──────────────────────────────────────────────────────────────

export const parseCsvFile = async (file: File): Promise<ParseResult> => {
  return new Promise((resolve, reject) => {
    Papa.parse<Record<string, string>>(file, {
      header: true,
      skipEmptyLines: true,
      transformHeader: normalizeHeader,
      transform: (v) => v.trim(),
      complete: (result) => {
        const fields = result.meta.fields ?? [];
        const { errors: headerErrors } = validateHeaders(fields);
        if (headerErrors.length > 0) {
          resolve({ rows: [], headerErrors, rowIssues: [] });
          return;
        }
        const rows = buildRows(result.data);
        const rowIssues = rows
          .map((r, i) => validateLedgerRow(r, i))
          .filter((x): x is RowIssue => x !== null);
        resolve({ rows, headerErrors: [], rowIssues });
      },
      error: reject,
    });
  });
};

// ─── Excel parser ─────────────────────────────────────────────────────────────

export const parseXlsxFile = async (file: File): Promise<ParseResult> => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) return { rows: [], headerErrors: ['Excel file contains no sheets.'], rowIssues: [] };

  const sheet = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '', raw: false }) as unknown[][];
  if (rawRows.length === 0) return { rows: [], headerErrors: ['Excel sheet is empty.'], rowIssues: [] };

  // Find the header row (first row containing at least 2 of our required headers)
  let headerRowIndex = -1;
  for (let i = 0; i < Math.min(rawRows.length, 20); i++) {
    const vals = (rawRows[i] as string[]).map((v) => normalizeHeader(String(v)));
    const matches = requiredHeaders.filter((h) => vals.includes(h));
    if (matches.length >= 2) { headerRowIndex = i; break; }
  }

  if (headerRowIndex === -1) {
    return {
      rows: [],
      headerErrors: [
        `Could not find required column headers in the first 20 rows. ` +
        `Your file must have columns: "Distribution account", "Distribution account type", and "Transaction type".`
      ],
      rowIssues: [],
    };
  }

  const headerRow = (rawRows[headerRowIndex] as string[]).map((v) => normalizeHeader(String(v)));
  const { errors: headerErrors } = validateHeaders(headerRow);
  if (headerErrors.length > 0) return { rows: [], headerErrors, rowIssues: [] };

  const dataRows = rawRows.slice(headerRowIndex + 1).map((rowArr) => {
    const obj: Record<string, unknown> = {};
    headerRow.forEach((h, i) => {
      obj[h] = (rowArr as unknown[])[i] ?? '';
    });
    return obj;
  });

  const rows = buildRows(dataRows);
  if (rows.length === 0) return { rows: [], headerErrors: ['No valid transaction rows found.'], rowIssues: [] };

  const rowIssues = rows
    .map((r, i) => validateLedgerRow(r, i))
    .filter((x): x is RowIssue => x !== null);

  return { rows, headerErrors: [], rowIssues };
};
