import type { LedgerRow } from '../../app/upload/upload-utils';
import { isCogsType, isExpenseType, isRevenueType, parseCurrencyAmount, roundCurrency } from './accounting';

export type MonthlyPLSummary = { revenue: number; cogs: number; expenses: number };

export type ProfitAndLoss = {
  totalRevenue: number;
  totalCogs: number;
  grossProfit: number;
  grossMargin: number;       // 0–1
  totalExpenses: number;     // opex only (excl COGS)
  totalAllExpenses: number;  // cogs + opex
  netProfit: number;
  netMargin: number;         // 0–1
  monthlyBreakdown: Record<string, MonthlyPLSummary>;
};

const pad = (v: number) => String(v).padStart(2, '0');

export const extractMonthKey = (value: string): string | null => {
  const s = value.trim();
  if (!s) return null;
  const iso = s.match(/^(\d{4})[-/](\d{1,2})[-/]\d{1,2}$/);
  if (iso) { const m = Number(iso[2]); if (m >= 1 && m <= 12) return `${iso[1]}-${pad(m)}`; }
  const slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slash) { const m = Number(slash[1]); if (m >= 1 && m <= 12) return `${slash[3]}-${pad(m)}`; }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
  return null;
};

export function generatePL(rows: LedgerRow[]): ProfitAndLoss {
  const monthly: Record<string, MonthlyPLSummary> = {};
  let totalRevenue = 0, totalCogs = 0, totalExpenses = 0;

  rows.forEach((row) => {
    const t = row['Distribution account type'];
    const amount = parseCurrencyAmount(row.Amount) ?? 0;
    const mk = extractMonthKey(row['Transaction date']);
    if (mk && !monthly[mk]) monthly[mk] = { revenue: 0, cogs: 0, expenses: 0 };

    if (isRevenueType(t)) {
      totalRevenue = roundCurrency(totalRevenue + amount);
      if (mk) monthly[mk].revenue = roundCurrency(monthly[mk].revenue + amount);
    } else if (isCogsType(t)) {
      totalCogs = roundCurrency(totalCogs + amount);
      if (mk) monthly[mk].cogs = roundCurrency(monthly[mk].cogs + amount);
    } else if (isExpenseType(t)) {
      totalExpenses = roundCurrency(totalExpenses + amount);
      if (mk) monthly[mk].expenses = roundCurrency(monthly[mk].expenses + amount);
    }
  });

  const grossProfit = roundCurrency(totalRevenue - totalCogs);
  const netProfit = roundCurrency(grossProfit - totalExpenses);

  return {
    totalRevenue,
    totalCogs,
    grossProfit,
    grossMargin: totalRevenue ? roundCurrency(grossProfit / totalRevenue) : 0,
    totalExpenses,
    totalAllExpenses: roundCurrency(totalCogs + totalExpenses),
    netProfit,
    netMargin: totalRevenue ? roundCurrency(netProfit / totalRevenue) : 0,
    monthlyBreakdown: Object.fromEntries(Object.entries(monthly).sort()),
  };
}
