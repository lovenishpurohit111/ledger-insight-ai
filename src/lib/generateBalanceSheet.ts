import type { LedgerRow } from '../../app/upload/upload-utils';
import { BS_TOLERANCE, classifyBalanceSheetType, isExpenseType, isRevenueType, parseCurrencyAmount, roundCurrency } from './accounting';

export type BalanceSheetEntry = { account: string; value: number };

export type BalanceSheet = {
  assets: BalanceSheetEntry[];
  liabilities: BalanceSheetEntry[];
  equity: BalanceSheetEntry[];
  totals: { assetsTotal: number; liabilitiesTotal: number; equityTotal: number };
  isBalanced: boolean;
  variance: number;
  currentPeriodEarnings: number;
};

const buildEntries = (m: Map<string, number>) =>
  Array.from(m.entries()).map(([account, value]) => ({ account, value }))
    .sort((a, b) => a.account.localeCompare(b.account));

const sumEntries = (entries: BalanceSheetEntry[]) =>
  roundCurrency(entries.reduce((t, e) => t + e.value, 0));

export function generateBalanceSheet(rows: LedgerRow[]): BalanceSheet {
  const assetsMap  = new Map<string, number>();
  const liabMap    = new Map<string, number>();
  const equityMap  = new Map<string, number>();
  // Track the last seen balance for each account (including $0 from empty)
  const lastBalance = new Map<string, { value: number; bsType: BalanceSheetCategory }>();
  let cpe = 0;

  rows.forEach((row) => {
    const accountType = row['Distribution account type'].trim();
    const account     = row['Distribution account'].trim();
    const amount      = parseCurrencyAmount(row.Amount) ?? 0;
    const bsType      = classifyBalanceSheetType(accountType);

    // CPE: use signed amounts directly
    if (isRevenueType(accountType)) cpe = roundCurrency(cpe + amount);
    if (isExpenseType(accountType)) cpe = roundCurrency(cpe - amount);

    if (!account || !bsType) return;

    // Use last balance per account — empty string means $0 (e.g. paid-off liability)
    const balRaw = row.Balance.trim();
    const bal = balRaw === '' ? 0 : parseCurrencyAmount(balRaw);
    if (bal === null) return; // unparseable — skip

    lastBalance.set(account, { value: roundCurrency(bal), bsType });
  });

  // Build BS maps from last-seen balance per account
  lastBalance.forEach(({ value, bsType }, account) => {
    if (bsType === 'asset')     assetsMap.set(account, value);
    if (bsType === 'liability') liabMap.set(account, value);
    if (bsType === 'equity')    equityMap.set(account, value);
  });

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
