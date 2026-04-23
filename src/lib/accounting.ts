export type BalanceSheetCategory = 'asset' | 'liability' | 'equity';

const assetTypes = new Set([
  'accounts receivable (a/r)',
  'bank',
  'fixed assets',
  'other assets',
  'other current assets',
]);

const liabilityTypes = new Set([
  'accounts payable (a/p)',
  'credit card',
  'long term liabilities',
  'other current liabilities',
]);

const normalizeValue = (value: string) => value.trim();

export const parseCurrencyAmount = (value: string) => {
  const normalized = normalizeValue(value);
  if (!normalized) {
    return null;
  }

  const isNegative = normalized.startsWith('(') && normalized.endsWith(')');
  const numericPortion = normalized.replace(/[,$()\s]/g, '');
  const parsed = Number(numericPortion);

  if (!Number.isFinite(parsed)) {
    return null;
  }

  return isNegative ? -parsed : parsed;
};

export const roundCurrency = (value: number) => Math.round((value + Number.EPSILON) * 100) / 100;

export const classifyBalanceSheetType = (value: string): BalanceSheetCategory | null => {
  const accountType = normalizeValue(value).toLowerCase();

  if (assetTypes.has(accountType) || accountType.includes('asset')) {
    return 'asset';
  }

  if (liabilityTypes.has(accountType) || accountType.includes('liabilit')) {
    return 'liability';
  }

  if (accountType === 'equity' || accountType.includes('equity')) {
    return 'equity';
  }

  return null;
};

export const isRevenueType = (value: string) => normalizeValue(value).toLowerCase().includes('income');

export const isExpenseType = (value: string) => {
  const accountType = normalizeValue(value).toLowerCase();

  return accountType.includes('expense') || accountType === 'expenses' || accountType === 'cost of goods sold';
};
