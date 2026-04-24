export type BalanceSheetCategory = 'asset' | 'liability' | 'equity';

const normalizeValue = (value: string) => value.trim();

export const parseCurrencyAmount = (value: string) => {
  const normalized = normalizeValue(value);
  if (!normalized) return null;
  const isNegative = normalized.startsWith('(') && normalized.endsWith(')')
    || normalized.startsWith('-');
  const numericPortion = normalized.replace(/[,$()\\-\s]/g, '');
  const parsed = Number(numericPortion);
  if (!Number.isFinite(parsed)) return null;
  return isNegative ? -parsed : parsed;
};

export const roundCurrency = (value: number) =>
  Math.round((value + Number.EPSILON) * 100) / 100;

// Tolerance for balance sheet check — real-world ledgers have $0–$5 rounding drift
export const BS_TOLERANCE = 1.0;

export const classifyBalanceSheetType = (value: string): BalanceSheetCategory | null => {
  const t = normalizeValue(value).toLowerCase();
  if (t.includes('asset') || t.includes('bank') || t.includes('receivable') || t.includes('prepaid') || t.includes('inventory')) return 'asset';
  if (t.includes('liabilit') || t.includes('payable') || t.includes('loan') || t.includes('credit card') || t.includes('mortgage')) return 'liability';
  if (t.includes('equity') || t.includes('capital') || t.includes('contribution') || t.includes('retained') || t.includes('partner')) return 'equity';
  return null;
};

export const isRevenueType = (value: string) => {
  const t = normalizeValue(value).toLowerCase();
  return t.includes('income') || t.includes('revenue') || t.includes('sales');
};

export const isExpenseType = (value: string) => {
  const t = normalizeValue(value).toLowerCase();
  return t.includes('expense') || t === 'expenses' || t.includes('cost of goods') || t.includes('cogs');
};
