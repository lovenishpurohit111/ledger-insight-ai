import type { LedgerRow } from '../../app/upload/upload-utils';
import { isAnyExpense, isCogsType, isCurrentAsset, isCurrentLiability, isExpenseType, isRevenueType, parseCurrencyAmount, roundCurrency } from './accounting';
import { extractMonthKey } from './generatePL';
import { verifyCategoriesFree, type CategoryVerification } from './verifyCategoryFree';

export type VendorSpend = { name: string; total: number; txCount: number; accounts: string[] };
export type RevenueSource = { name: string; total: number; txCount: number };
export type AnomalyTransaction = { row: LedgerRow; reason: string; amount: number; zScore: number };
export type AuditFlag = { type: 'gap' | 'round' | 'weekend' | 'large' | 'unsequenced'; description: string; rows: LedgerRow[] };

export type FinancialInsights = {
  // Burn rate & runway
  avgMonthlyBurn: number;
  avgMonthlyRevenue: number;
  cashBalance: number;             // latest bank balance
  runwayMonths: number | null;

  // Pareto
  topVendors: VendorSpend[];
  topRevenueSources: RevenueSource[];

  // Anomalies
  anomalies: AnomalyTransaction[];

  // Audit flags
  auditFlags: AuditFlag[];

  // Tax estimate (simple: net profit × 21% corp / 30% sole)
  taxEstimate: { rate: number; amount: number; basis: string };

  // Monthly trends
  monthlyBurn: Record<string, number>;
  monthlyCashPosition: Record<string, number>;
  categoryMismatches: CategoryVerification[];
};

export function generateInsights(rows: LedgerRow[]): FinancialInsights {
  const monthlyBurn: Record<string, number> = {};
  const monthlyRevenue: Record<string, number> = {};
  const vendorMap  = new Map<string, VendorSpend>();
  const revenueMap = new Map<string, RevenueSource>();
  const amounts: number[] = [];
  let cashBalance = 0;

  // Collect last bank balance for cash position
  const bankBalances = new Map<string, number>();

  rows.forEach(row => {
    const t    = row['Distribution account type'].trim();
    const acct = row['Distribution account'].trim();
    const name = row.Name.trim();
    const amt  = parseCurrencyAmount(row.Amount) ?? 0;
    const bal  = parseCurrencyAmount(row.Balance);
    const mk   = extractMonthKey(row['Transaction date']);

    amounts.push(Math.abs(amt));

    // Bank cash tracking
    if (isCurrentAsset(t) && acct.toLowerCase().includes('bank') || acct.toLowerCase().includes('checking') || acct.toLowerCase().includes('savings') || acct.toLowerCase().includes('op acct')) {
      if (bal !== null) bankBalances.set(acct, bal);
    }

    // Monthly burn
    if (isAnyExpense(t) && mk) {
      monthlyBurn[mk] = roundCurrency((monthlyBurn[mk] ?? 0) + Math.abs(amt));
    }
    if (isRevenueType(t) && mk) {
      monthlyRevenue[mk] = roundCurrency((monthlyRevenue[mk] ?? 0) + amt);
    }

    // Vendor spend (expenses only)
    if (isAnyExpense(t) && name) {
      const v = vendorMap.get(name) ?? { name, total: 0, txCount: 0, accounts: [] };
      v.total = roundCurrency(v.total + Math.abs(amt));
      v.txCount++;
      if (!v.accounts.includes(acct)) v.accounts.push(acct);
      vendorMap.set(name, v);
    }

    // Revenue sources
    if (isRevenueType(t) && name) {
      const r = revenueMap.get(name) ?? { name, total: 0, txCount: 0 };
      r.total = roundCurrency(r.total + amt);
      r.txCount++;
      revenueMap.set(name, r);
    }
  });

  cashBalance = Array.from(bankBalances.values()).reduce((t, v) => t + v, 0);

  // Burn rate & runway
  const burnValues = Object.values(monthlyBurn);
  const avgMonthlyBurn = burnValues.length ? roundCurrency(burnValues.reduce((a, b) => a + b, 0) / burnValues.length) : 0;
  const revValues = Object.values(monthlyRevenue);
  const avgMonthlyRevenue = revValues.length ? roundCurrency(revValues.reduce((a, b) => a + b, 0) / revValues.length) : 0;
  const netBurn = roundCurrency(avgMonthlyBurn - avgMonthlyRevenue);
  const runwayMonths = netBurn > 0 && cashBalance > 0 ? roundCurrency(cashBalance / netBurn) : null;

  // Top vendors (top 10)
  const topVendors = Array.from(vendorMap.values())
    .filter(v => v.name)
    .sort((a, b) => b.total - a.total)
    .slice(0, 10);

  // Top revenue sources (top 10)
  const topRevenueSources = Array.from(revenueMap.values())
    .filter(r => r.name && r.total > 0)
    .sort((a, b) => b.total - a.total)
    .slice(0, 10);

  // Anomaly detection — z-score on amounts
  const absAmounts = rows.map(r => Math.abs(parseCurrencyAmount(r.Amount) ?? 0)).filter(a => a > 0);
  const mean = absAmounts.reduce((a, b) => a + b, 0) / (absAmounts.length || 1);
  const stdDev = Math.sqrt(absAmounts.map(a => (a - mean) ** 2).reduce((a, b) => a + b, 0) / (absAmounts.length || 1));
  const ZSCORE_THRESHOLD = 3;

  const anomalies: AnomalyTransaction[] = rows
    .map(row => {
      const amount = Math.abs(parseCurrencyAmount(row.Amount) ?? 0);
      const zScore = stdDev > 0 ? roundCurrency((amount - mean) / stdDev) : 0;
      if (zScore < ZSCORE_THRESHOLD) return null;
      return { row, amount, zScore, reason: `Amount ${zScore.toFixed(1)}σ above average ($${mean.toFixed(0)} avg)` };
    })
    .filter((x): x is AnomalyTransaction => x !== null)
    .sort((a, b) => b.zScore - a.zScore)
    .slice(0, 15);

  // Audit flags
  const auditFlags: AuditFlag[] = [];

  // 1. Large round-number transactions (potential estimates/placeholders)
  const roundRows = rows.filter(r => {
    const amt = Math.abs(parseCurrencyAmount(r.Amount) ?? 0);
    return amt >= 1000 && amt % 500 === 0;
  });
  if (roundRows.length > 0) auditFlags.push({ type: 'round', description: `${roundRows.length} large round-number transactions (multiples of $500) — may be estimates`, rows: roundRows.slice(0, 10) });

  // 2. Weekend transactions (unusual for most businesses)
  const weekendRows = rows.filter(r => {
    const d = new Date(r['Transaction date']);
    const day = d.getDay();
    return day === 0 || day === 6;
  });
  if (weekendRows.length > 3) auditFlags.push({ type: 'weekend', description: `${weekendRows.length} transactions on weekends — verify these are legitimate`, rows: weekendRows.slice(0, 5) });

  // 3. Missing transaction number sequences
  const nums = rows.map(r => parseInt(r.Num, 10)).filter(n => !isNaN(n) && n > 0).sort((a, b) => a - b);
  const gaps: number[] = [];
  for (let i = 1; i < nums.length; i++) {
    if (nums[i] - nums[i - 1] > 1) gaps.push(nums[i - 1]);
  }
  if (gaps.length > 0) auditFlags.push({ type: 'gap', description: `${gaps.length} gaps in transaction number sequence — possible missing entries after #${gaps.slice(0, 3).join(', #')}`, rows: [] });

  // Monthly cash position
  const monthlyCashPosition: Record<string, number> = {};
  const sortedMonths = Object.keys(monthlyBurn).sort();
  let runningCash = cashBalance;
  for (const m of sortedMonths.reverse()) {
    const burn = monthlyBurn[m] ?? 0;
    const rev  = monthlyRevenue[m] ?? 0;
    runningCash += burn - rev;
    monthlyCashPosition[m] = roundCurrency(runningCash);
  }

  // Tax estimate (simple)
  const netProfit = roundCurrency(avgMonthlyRevenue * revValues.length - avgMonthlyBurn * burnValues.length);
  const taxRate = 0.21;
  const taxEstimate = { rate: taxRate, amount: roundCurrency(Math.max(0, netProfit * taxRate)), basis: `${(taxRate * 100).toFixed(0)}% flat rate on estimated net profit of $${netProfit.toFixed(2)}` };

  return { avgMonthlyBurn, avgMonthlyRevenue, cashBalance, runwayMonths, topVendors, topRevenueSources, anomalies, auditFlags, taxEstimate, monthlyBurn, monthlyCashPosition, categoryMismatches: verifyCategoriesFree(rows) };
}
