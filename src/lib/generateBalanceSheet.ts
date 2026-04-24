import type { LedgerRow } from '../../app/upload/upload-utils';
import { BS_TOLERANCE, classifyBalanceSheetType, isExpenseType, isRevenueType, parseCurrencyAmount, roundCurrency } from './accounting';

export type BalanceSheetEntry = { account: string; value: number };

export type BalanceSheet = {
  assets: BalanceSheetEntry[];
  liabilities: BalanceSheetEntry[];
  equity: BalanceSheetEntry[];
  totals: { assetsTotal: number; liabilitiesTotal: number; equityTotal: number };
  isBalanced: boolean;
  variance: number;        // Assets - (Liabilities + Equity)
  currentPeriodEarnings: number;
};

const buildEntries = (m: Map<string, number>) =>
  Array.from(m.entries())
    .map(([account, value]) => ({ account, value }))
    .sort((a, b) => a.account.localeCompare(b.account));

const sumEntries = (entries: BalanceSheetEntry[]) =>
  roundCurrency(entries.reduce((t, e) => t + e.value, 0));

export function generateBalanceSheet(rows: LedgerRow[]): BalanceSheet {
  const assetsMap     = new Map<string, number>();
  const liabMap       = new Map<string, number>();
  const equityMap     = new Map<string, number>();
  let cpe = 0;

  rows.forEach((row) => {
    const accountType = row['Distribution account type'];
    const account     = row['Distribution account'].trim();
    const amount      = parseCurrencyAmount(row.Amount) ?? 0;
    const balance     = parseCurrencyAmount(row.Balance);
    const bsType      = classifyBalanceSheetType(accountType);

    // Current Period Earnings: Revenue amounts add, Expense amounts subtract
    // QBO amounts are positive for both income and expense entries (debit-normal convention)
    if (isRevenueType(accountType)) cpe = roundCurrency(cpe + Math.abs(amount));
    if (isExpenseType(accountType)) cpe = roundCurrency(cpe - Math.abs(amount));

    if (!account || balance === null) return;

    if (bsType === 'asset')     assetsMap.set(account, roundCurrency(balance));
    if (bsType === 'liability') liabMap.set(account, roundCurrency(balance));
    if (bsType === 'equity')    equityMap.set(account, roundCurrency(balance));
  });

  // Only inject CPE if it's not already captured in equity account balances
  if (cpe !== 0) equityMap.set('Current Period Earnings', cpe);

  const assets      = buildEntries(assetsMap);
  const liabilities = buildEntries(liabMap);
  const equity      = buildEntries(equityMap);

  const assetsTotal      = sumEntries(assets);
  const liabilitiesTotal = sumEntries(liabilities);
  const equityTotal      = sumEntries(equity);
  const variance         = roundCurrency(assetsTotal - (liabilitiesTotal + equityTotal));

  return {
    assets, liabilities, equity,
    totals: { assetsTotal, liabilitiesTotal, equityTotal },
    isBalanced: Math.abs(variance) <= BS_TOLERANCE,
    variance,
    currentPeriodEarnings: cpe,
  };
}
