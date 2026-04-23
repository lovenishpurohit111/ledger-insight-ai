import type { LedgerRow } from '../../app/upload/upload-utils';

export type MonthlyPLSummary = {
  revenue: number;
  expenses: number;
};

export type ProfitAndLoss = {
  totalRevenue: number;
  totalExpenses: number;
  netProfit: number;
  monthlyBreakdown: Record<string, MonthlyPLSummary>;
};

const normalizeValue = (value: string) => value.trim();

const isRevenueType = (accountType: string) => accountType.includes('income');

const isExpenseType = (accountType: string) =>
  accountType.includes('expense') || accountType === 'expenses' || accountType === 'cost of goods sold';

const parseAmount = (value: string) => {
  const normalized = normalizeValue(value);
  if (!normalized) {
    return 0;
  }

  const isNegative = normalized.startsWith('(') && normalized.endsWith(')');
  const numericPortion = normalized.replace(/[,$()\s]/g, '');
  const parsed = Number(numericPortion);

  if (!Number.isFinite(parsed)) {
    return 0;
  }

  return isNegative ? -parsed : parsed;
};

const padMonth = (value: number) => String(value).padStart(2, '0');

const extractMonthKey = (value: string) => {
  const normalized = normalizeValue(value);
  if (!normalized) {
    return null;
  }

  const isoMatch = normalized.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (isoMatch) {
    const [, year, month] = isoMatch;
    const monthNumber = Number(month);
    if (monthNumber >= 1 && monthNumber <= 12) {
      return `${year}-${padMonth(monthNumber)}`;
    }
  }

  const slashMatch = normalized.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const [, month, , year] = slashMatch;
    const monthNumber = Number(month);
    if (monthNumber >= 1 && monthNumber <= 12) {
      return `${year}-${padMonth(monthNumber)}`;
    }
  }

  const parsedDate = new Date(normalized);
  if (Number.isNaN(parsedDate.getTime())) {
    return null;
  }

  return `${parsedDate.getFullYear()}-${padMonth(parsedDate.getMonth() + 1)}`;
};

const updateMonthlyBreakdown = (
  monthlyBreakdown: Record<string, MonthlyPLSummary>,
  monthKey: string | null,
  amount: number,
  category: 'revenue' | 'expenses',
) => {
  if (!monthKey) {
    return;
  }

  const entry = monthlyBreakdown[monthKey] ?? { revenue: 0, expenses: 0 };
  entry[category] += amount;
  monthlyBreakdown[monthKey] = entry;
};

export function generatePL(rows: LedgerRow[]): ProfitAndLoss {
  const monthlyBreakdown: Record<string, MonthlyPLSummary> = {};
  let totalRevenue = 0;
  let totalExpenses = 0;

  rows.forEach((row) => {
    const accountType = normalizeValue(row['Distribution account type']).toLowerCase();
    const amount = parseAmount(row.Amount);
    const monthKey = extractMonthKey(row['Transaction date']);

    if (isRevenueType(accountType)) {
      totalRevenue += amount;
      updateMonthlyBreakdown(monthlyBreakdown, monthKey, amount, 'revenue');
    }

    if (isExpenseType(accountType)) {
      totalExpenses += amount;
      updateMonthlyBreakdown(monthlyBreakdown, monthKey, amount, 'expenses');
    }
  });

  const sortedMonthlyBreakdown = Object.fromEntries(
    Object.entries(monthlyBreakdown).sort(([leftMonth], [rightMonth]) => leftMonth.localeCompare(rightMonth)),
  );

  return {
    totalRevenue,
    totalExpenses,
    netProfit: totalRevenue - totalExpenses,
    monthlyBreakdown: sortedMonthlyBreakdown,
  };
}
