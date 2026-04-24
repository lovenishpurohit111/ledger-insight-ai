import type { LedgerRow } from '../../app/upload/upload-utils';
import { isExpenseType, isRevenueType, parseCurrencyAmount, roundCurrency } from './accounting';

export type MoMCategory = {
  name: string;
  type: 'income' | 'expense';
  months: Record<string, number>; // month key → amount
  total: number;
};

export type MoMPL = {
  months: string[];              // sorted month keys e.g. ["2024-01", "2024-02"]
  incomeCategories: MoMCategory[];
  expenseCategories: MoMCategory[];
  monthlyRevenue: Record<string, number>;
  monthlyExpenses: Record<string, number>;
  monthlyNetProfit: Record<string, number>;
  totalRevenue: number;
  totalExpenses: number;
  totalNetProfit: number;
};

const padMonth = (v: number) => String(v).padStart(2, '0');

const extractMonthKey = (value: string): string | null => {
  const s = value.trim();
  if (!s) return null;

  const iso = s.match(/^(\d{4})[-/](\d{1,2})[-/]\d{1,2}$/);
  if (iso) {
    const m = Number(iso[2]);
    if (m >= 1 && m <= 12) return `${iso[1]}-${padMonth(m)}`;
  }

  const slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slash) {
    const m = Number(slash[1]);
    if (m >= 1 && m <= 12) return `${slash[3]}-${padMonth(m)}`;
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) return `${d.getFullYear()}-${padMonth(d.getMonth() + 1)}`;
  return null;
};

export const monthLabel = (key: string) => {
  const [year, month] = key.split('-');
  const d = new Date(Number(year), Number(month) - 1, 1);
  return d.toLocaleString('en-US', { month: 'short', year: 'numeric' });
};

export function generateMoMPL(rows: LedgerRow[]): MoMPL {
  // category name → month key → total amount
  const incomeCats  = new Map<string, Map<string, number>>();
  const expenseCats = new Map<string, Map<string, number>>();
  const monthSet    = new Set<string>();

  rows.forEach((row) => {
    const accountType = row['Distribution account type'];
    const account     = row['Distribution account'].trim();
    const amount      = parseCurrencyAmount(row.Amount) ?? 0;
    const monthKey    = extractMonthKey(row['Transaction date']);
    if (!monthKey || !account) return;

    monthSet.add(monthKey);

    if (isRevenueType(accountType)) {
      const cat = incomeCats.get(account) ?? new Map<string, number>();
      cat.set(monthKey, roundCurrency((cat.get(monthKey) ?? 0) + Math.abs(amount)));
      incomeCats.set(account, cat);
    }

    if (isExpenseType(accountType)) {
      const cat = expenseCats.get(account) ?? new Map<string, number>();
      cat.set(monthKey, roundCurrency((cat.get(monthKey) ?? 0) + Math.abs(amount)));
      expenseCats.set(account, cat);
    }
  });

  const months = Array.from(monthSet).sort();

  const buildCategory = (name: string, monthMap: Map<string, number>, type: 'income' | 'expense'): MoMCategory => {
    const monthAmounts: Record<string, number> = {};
    let total = 0;
    for (const m of months) {
      const v = monthMap.get(m) ?? 0;
      monthAmounts[m] = v;
      total = roundCurrency(total + v);
    }
    return { name, type, months: monthAmounts, total };
  };

  const incomeCategories  = Array.from(incomeCats.entries())
    .map(([name, monthMap]) => buildCategory(name, monthMap, 'income'))
    .sort((a, b) => b.total - a.total);

  const expenseCategories = Array.from(expenseCats.entries())
    .map(([name, monthMap]) => buildCategory(name, monthMap, 'expense'))
    .sort((a, b) => b.total - a.total);

  // Monthly totals
  const monthlyRevenue:   Record<string, number> = {};
  const monthlyExpenses:  Record<string, number> = {};
  const monthlyNetProfit: Record<string, number> = {};

  for (const m of months) {
    const rev = roundCurrency(incomeCategories.reduce((t, c) => t + (c.months[m] ?? 0), 0));
    const exp = roundCurrency(expenseCategories.reduce((t, c) => t + (c.months[m] ?? 0), 0));
    monthlyRevenue[m]   = rev;
    monthlyExpenses[m]  = exp;
    monthlyNetProfit[m] = roundCurrency(rev - exp);
  }

  const totalRevenue   = roundCurrency(incomeCategories.reduce((t, c) => t + c.total, 0));
  const totalExpenses  = roundCurrency(expenseCategories.reduce((t, c) => t + c.total, 0));

  return {
    months, incomeCategories, expenseCategories,
    monthlyRevenue, monthlyExpenses, monthlyNetProfit,
    totalRevenue, totalExpenses,
    totalNetProfit: roundCurrency(totalRevenue - totalExpenses),
  };
}

export const momChange = (current: number, previous: number): number | null => {
  if (previous === 0) return null;
  return roundCurrency(((current - previous) / Math.abs(previous)) * 100);
};
