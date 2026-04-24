export type BalanceSheetCategory = 'asset' | 'liability' | 'equity';

export const parseCurrencyAmount = (value: string): number | null => {
  const s = value.trim();
  if (!s) return null;
  const isNeg = (s.startsWith('(') && s.endsWith(')')) || s.startsWith('-');
  const num = Number(s.replace(/[,$() -]/g, ''));
  if (!Number.isFinite(num)) return null;
  return isNeg ? -num : num;
};

export const roundCurrency = (v: number) =>
  Math.round((v + Number.EPSILON) * 100) / 100;

// $1 tolerance for real-world rounding drift
export const BS_TOLERANCE = 1.0;

// ─── QBO exact account type → BS category ────────────────────────────────────
// Source: QuickBooks Online Chart of Accounts type list
const QBO_ASSET_TYPES = new Set([
  'bank',
  'accounts receivable (a/r)',
  'accounts receivable',
  'other current assets',
  'other current asset',
  'fixed assets',
  'fixed asset',
  'other assets',
  'other asset',
  'inventory',
  'asset',
  'assets',
]);

const QBO_LIABILITY_TYPES = new Set([
  'accounts payable (a/p)',
  'accounts payable',
  'credit card',
  'other current liabilities',
  'other current liability',
  'long term liabilities',
  'long-term liabilities',
  'long term liability',
  'other liability',
  'other liabilities',
  'liability',
  'liabilities',
]);

const QBO_EQUITY_TYPES = new Set([
  'equity',
  'equities',
  'retained earnings',
  'opening balance equity',
  "owner's equity",
  'partner equity',
  'partner contribution',
  'partner distribution',
  'stockholders equity',
  'shareholder equity',
]);

const QBO_INCOME_TYPES = new Set([
  'income',
  'revenue',
  'sales',
  'other income',
  'other revenue',
  'non-operating income',
]);

const QBO_EXPENSE_TYPES = new Set([
  'expense',
  'expenses',
  'cost of goods sold',
  'cogs',
  'other expense',
  'other expenses',
  'non-operating expense',
]);

export const classifyBalanceSheetType = (value: string): BalanceSheetCategory | null => {
  const t = value.trim().toLowerCase();
  if (QBO_ASSET_TYPES.has(t))     return 'asset';
  if (QBO_LIABILITY_TYPES.has(t)) return 'liability';
  if (QBO_EQUITY_TYPES.has(t))    return 'equity';
  // Keyword fallback for custom/non-standard types
  if (t.includes('asset') || t.includes('receivable') || t.includes('prepaid') || t.includes('bank') || t.includes('cash')) return 'asset';
  if (t.includes('liabilit') || t.includes('payable') || t.includes('credit card') || t.includes('loan') || t.includes('mortgage')) return 'liability';
  if (t.includes('equity') || t.includes('capital') || t.includes('contribution') || t.includes('retained')) return 'equity';
  return null;
};

export const isRevenueType = (value: string): boolean => {
  const t = value.trim().toLowerCase();
  if (QBO_INCOME_TYPES.has(t)) return true;
  return t.includes('income') || t.includes('revenue') || t.includes('sales');
};

export const isExpenseType = (value: string): boolean => {
  const t = value.trim().toLowerCase();
  if (QBO_EXPENSE_TYPES.has(t)) return true;
  return t.includes('expense') || t.includes('cost of goods') || t.includes('cogs');
};
