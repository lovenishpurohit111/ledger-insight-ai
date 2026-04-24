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
  'Transaction date',
  'Transaction type',
];

const normalizeHeader = (value: string) => value.trim();

const hasAnyValue = (row: Partial<Record<HeaderKey, string>>) =>
  requiredHeaders.some((header) => Boolean(row[header]?.trim()));

const isNonTransactionRow = (row: Partial<Record<HeaderKey, string>>) =>
  (!row['Transaction date']?.trim() && !row['Transaction type']?.trim()) ||
  (!row.Amount?.trim() && !row.Balance?.trim());

// ─── QBO Account Type Inference ───────────────────────────────────────────────
// QBO general ledger exports don't include a "Distribution account type" column.
// We infer it from the account name using common QBO chart-of-accounts patterns.

const EXPENSE_KEYWORDS = [
  'expense', 'advertising', 'marketing', 'rent', 'utilities', 'insurance',
  'payroll', 'salary', 'salaries', 'wages', 'supplies', 'meals', 'travel',
  'repairs', 'maintenance', 'license', 'permit', 'legal', 'professional',
  'bank fee', 'service charge', 'interest paid', 'charitable', 'computer',
  'office', 'telephone', 'internet', 'printing', 'shipping', 'postage',
  'software', 'web', 'security deposit', 'memberships', 'subscriptions',
  'r & d', 'research', 'development', 'fees',
];

const INCOME_KEYWORDS = [
  'income', 'revenue', 'sales', 'other income', 'service revenue',
  'product revenue', 'interest income', 'rental income', 'grant',
];

const ASSET_KEYWORDS = [
  'checking', 'savings', 'bank', 'cash', 'accounts receivable', 'a/r',
  'inventory', 'prepaid', 'asset', 'equipment', 'furniture', 'fixture',
  'vehicle', 'building', 'land', 'deposit', 'op acct', 'operating account',
];

const LIABILITY_KEYWORDS = [
  'payable', 'a/p', 'accounts payable', 'credit card', 'loan', 'mortgage',
  'liability', 'liabilities', 'debt', 'note payable', 'deferred',
];

const EQUITY_KEYWORDS = [
  'equity', 'capital', 'contribution', 'retained earnings', 'owner',
  'partner', 'drawing', 'distribution', 'stockholder', 'shareholder',
];

export const inferAccountType = (accountName: string): string => {
  const lower = accountName.toLowerCase();
  if (INCOME_KEYWORDS.some((k) => lower.includes(k))) return 'Income';
  if (EXPENSE_KEYWORDS.some((k) => lower.includes(k))) return 'Expense';
  if (ASSET_KEYWORDS.some((k) => lower.includes(k))) return 'Asset';
  if (LIABILITY_KEYWORDS.some((k) => lower.includes(k))) return 'Liability';
  if (EQUITY_KEYWORDS.some((k) => lower.includes(k))) return 'Equity';
  return 'Expense'; // QBO default fallback — most unrecognized accounts are expenses
};

// ─── Standard format validation ───────────────────────────────────────────────

export const validateHeaders = (headers: string[]) => {
  const trimmed = headers.map(normalizeHeader);
  // Allow missing "Distribution account type" — it can be inferred
  const optionalHeaders = new Set<string>(['Distribution account type']);
  const missing = requiredHeaders.filter(
    (h) => !optionalHeaders.has(h) && !trimmed.includes(h),
  );
  const errors: string[] = [];
  if (missing.length > 0) {
    errors.push(`Missing required headers: ${missing.join(', ')}`);
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
    if (value && parseCurrencyAmount(value) === null) {
      issues.push(`${field} must be numeric.`);
    }
  });

  return issues.length > 0 ? { rowIndex, issues } : null;
};

const buildLedgerRows = (rawRows: Array<Record<string, unknown>>) => {
  return rawRows
    .map((rawRow) => {
      return requiredHeaders.reduce((acc, header) => {
        const rawValue = rawRow[header] ?? rawRow[header.trim()] ?? '';
        acc[header] = String(rawValue).trim();
        return acc;
      }, {} as LedgerRow);
    })
    .map((row) => {
      if (!row['Distribution account type'] && row['Distribution account']) {
        row['Distribution account type'] = inferAccountType(row['Distribution account']);
      }
      return row;
    })
    .filter((row) => hasAnyValue(row) && !isNonTransactionRow(row));
};

export const isCsvFile = (file: File) =>
  file.type === 'text/csv' || file.name.toLowerCase().endsWith('.csv');

export const isExcelFile = (file: File) =>
  file.type.includes('spreadsheet') ||
  file.name.toLowerCase().endsWith('.xlsx') ||
  file.name.toLowerCase().endsWith('.xls');

// ─── CSV Parser ───────────────────────────────────────────────────────────────

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
      error: (error) => reject(error),
    });
  });
};

// ─── Excel Parser (handles standard + QBO format) ────────────────────────────

const QBO_STANDARD_HEADERS = [
  'Transaction date',
  'Transaction type',
  'Num',
  'Name',
  'Description',
  'Split',
  'Amount',
  'Balance',
];

const isQboHeaderRow = (row: unknown[]): boolean => {
  const vals = row.map((v) => String(v ?? '').trim().toLowerCase());
  const matches = QBO_STANDARD_HEADERS.filter((h) =>
    vals.includes(h.toLowerCase()),
  );
  return matches.length >= 4;
};

const isSubtotalRow = (row: unknown[]): boolean => {
  const first = String(row[0] ?? '').trim().toLowerCase();
  return first.startsWith('total for') || first.startsWith('total ');
};

const isMetadataRow = (row: unknown[]): boolean => {
  // Rows with only col A filled and everything else empty = group header or metadata
  const hasColA = Boolean(String(row[0] ?? '').trim());
  const restEmpty = row.slice(1).every((v) => !String(v ?? '').trim());
  return hasColA && restEmpty;
};

export const parseXlsxFile = async (file: File): Promise<ParseResult> => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    return { rows: [], headerErrors: ['Excel file contains no sheets.'], rowIssues: [] };
  }

  const sheet = workbook.Sheets[firstSheetName];
  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: '',
    raw: false, // format dates as strings
  }) as unknown[][];

  if (rawRows.length === 0) {
    return { rows: [], headerErrors: ['Excel sheet is empty.'], rowIssues: [] };
  }

  // ── Auto-detect header row ──────────────────────────────────────────────────
  let headerRowIndex = -1;
  for (let i = 0; i < Math.min(rawRows.length, 20); i++) {
    if (isQboHeaderRow(rawRows[i])) {
      headerRowIndex = i;
      break;
    }
  }

  // ── Standard format (header at row 0) ──────────────────────────────────────
  if (headerRowIndex === -1) {
    const headerRow = rawRows[0] as string[];
    const normalizedHeaders = headerRow.map((h) => normalizeHeader(String(h)));
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
      .map((row) => {
        if (!row['Distribution account type'] && row['Distribution account']) {
          row['Distribution account type'] = inferAccountType(row['Distribution account']);
        }
        return row;
      })
      .filter((row) => hasAnyValue(row) && !isNonTransactionRow(row));

    const rowIssues = rows
      .map((row, index) => validateLedgerRow(row, index))
      .filter((issue): issue is RowIssue => issue !== null);
    return { rows, headerErrors: [], rowIssues };
  }

  // ── QBO format ──────────────────────────────────────────────────────────────
  // Col A contains the account group name; col B onwards = the standard columns
  const headerRow = rawRows[headerRowIndex] as string[];
  // Col A is blank in the header row — map col B+ to standard headers
  // headerRow[0] is blank, headerRow[1..] = 'Distribution account', 'Transaction date' etc.
  const colHeaders = headerRow.map((h) => normalizeHeader(String(h)));

  // Index of each standard header in the column array
  const colIndex = (name: string) => colHeaders.indexOf(name);

  const distAcctCol = colIndex('Distribution account');
  const txDateCol   = colIndex('Transaction date');
  const txTypeCol   = colIndex('Transaction type');
  const numCol      = colIndex('Num');
  const nameCol     = colIndex('Name');
  const descCol     = colIndex('Description');
  const splitCol    = colIndex('Split');
  const amountCol   = colIndex('Amount');
  const balanceCol  = colIndex('Balance');

  let currentAccountGroup = '';
  const rows: LedgerRow[] = [];

  for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
    const row = rawRows[i] as unknown[];

    // Skip subtotal rows
    if (isSubtotalRow(row)) continue;

    const colAVal = String(row[0] ?? '').trim();

    // Group header row (only col A has value) — update current account
    if (isMetadataRow(row)) {
      if (colAVal && !colAVal.toLowerCase().startsWith('accrual') && !colAVal.toLowerCase().startsWith('cash')) {
        currentAccountGroup = colAVal;
      }
      continue;
    }

    // Skip rows with no transaction date or amount
    const txDate = String(row[txDateCol] ?? '').trim();
    const amount = String(row[amountCol] ?? '').trim();
    if (!txDate && !amount) continue;

    const distAcct = distAcctCol >= 0
      ? String(row[distAcctCol] ?? '').trim()
      : currentAccountGroup;

    const accountName = distAcct || currentAccountGroup;
    const accountType = inferAccountType(accountName);

    const ledgerRow: LedgerRow = {
      'Distribution account': accountName,
      'Distribution account type': accountType,
      'Transaction date': txDate,
      'Transaction type': txTypeCol >= 0 ? String(row[txTypeCol] ?? '').trim() : '',
      'Num': numCol >= 0 ? String(row[numCol] ?? '').trim() : '',
      'Name': nameCol >= 0 ? String(row[nameCol] ?? '').trim() : '',
      'Description': descCol >= 0 ? String(row[descCol] ?? '').trim() : '',
      'Split': splitCol >= 0 ? String(row[splitCol] ?? '').trim() : '',
      'Amount': amount,
      'Balance': balanceCol >= 0 ? String(row[balanceCol] ?? '').trim() : '',
    };

    if (!isNonTransactionRow(ledgerRow)) {
      rows.push(ledgerRow);
    }
  }

  if (rows.length === 0) {
    return { rows: [], headerErrors: ['No transaction rows found in file.'], rowIssues: [] };
  }

  const rowIssues = rows
    .map((row, index) => validateLedgerRow(row, index))
    .filter((issue): issue is RowIssue => issue !== null);

  return { rows, headerErrors: [], rowIssues };
};
