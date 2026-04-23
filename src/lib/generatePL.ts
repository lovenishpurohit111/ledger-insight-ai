import type { LedgerRow } from '../../app/upload/upload-utils';
import { isExpenseType, isRevenueType, parseCurrencyAmount, roundCurrency } from './accounting';

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

const parseAmount = (value: string) => {
  return parseCurrencyAmount(value) ?? 0;
};

const padMonth = (value: number) => String(value).padStart(2, '0');

const extractMonthKey = (value: string) => {
  const normalized = value.trim();
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
  entry[category] = roundCurrency(entry[category] + amount);
  monthlyBreakdown[monthKey] = entry;
};

export function generatePL(rows: LedgerRow[]): ProfitAndLoss {
  const monthlyBreakdown: Record<string, MonthlyPLSummary> = {};
  let totalRevenue = 0;
  let totalExpenses = 0;

  rows.forEach((row) => {
    const accountType = row['Distribution account type'];
    const amount = parseAmount(row.Amount);
    const monthKey = extractMonthKey(row['Transaction date']);

    if (isRevenueType(accountType)) {
      totalRevenue = roundCurrency(totalRevenue + amount);
      updateMonthlyBreakdown(monthlyBreakdown, monthKey, amount, 'revenue');
    }

    if (isExpenseType(accountType)) {
      totalExpenses = roundCurrency(totalExpenses + amount);
      updateMonthlyBreakdown(monthlyBreakdown, monthKey, amount, 'expenses');
    }
  });

  const sortedMonthlyBreakdown = Object.fromEntries(
    Object.entries(monthlyBreakdown).sort(([leftMonth], [rightMonth]) => leftMonth.localeCompare(rightMonth)),
  );

  return {
    totalRevenue,
    totalExpenses,
    netProfit: roundCurrency(totalRevenue - totalExpenses),
    monthlyBreakdown: sortedMonthlyBreakdown,
  };
}
