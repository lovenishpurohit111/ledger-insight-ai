import type { LedgerRow } from '../../app/upload/upload-utils';
import { classifyBalanceSheetType, isExpenseType, isRevenueType, parseCurrencyAmount, roundCurrency } from './accounting';

export type BalanceSheetEntry = {
  account: string;
  value: number;
};

export type BalanceSheet = {
  assets: BalanceSheetEntry[];
  liabilities: BalanceSheetEntry[];
  equity: BalanceSheetEntry[];
  totals: {
    assetsTotal: number;
    liabilitiesTotal: number;
    equityTotal: number;
  };
  isBalanced: boolean;
};

const normalizeValue = (value: string) => value.trim();

const buildEntries = (accounts: Map<string, number>) =>
  Array.from(accounts.entries())
    .map(([account, value]) => ({ account, value }))
    .sort((left, right) => left.account.localeCompare(right.account));

const sumEntries = (entries: BalanceSheetEntry[]) => roundCurrency(entries.reduce((total, entry) => total + entry.value, 0));

export function generateBalanceSheet(rows: LedgerRow[]): BalanceSheet {
  const assetsMap = new Map<string, number>();
  const liabilitiesMap = new Map<string, number>();
  const equityMap = new Map<string, number>();
  let currentPeriodEarnings = 0;

  rows.forEach((row) => {
    const accountType = row['Distribution account type'];
    const account = normalizeValue(row['Distribution account']);
    const amount = parseCurrencyAmount(row.Amount) ?? 0;
    const balance = parseCurrencyAmount(row.Balance);
    const balanceSheetType = classifyBalanceSheetType(accountType);

    if (isRevenueType(accountType)) {
      currentPeriodEarnings = roundCurrency(currentPeriodEarnings + amount);
    }

    if (isExpenseType(accountType)) {
      currentPeriodEarnings = roundCurrency(currentPeriodEarnings - amount);
    }

    if (!account || balance === null) {
      return;
    }

    if (balanceSheetType === 'asset') {
      assetsMap.set(account, roundCurrency(balance));
    }

    if (balanceSheetType === 'liability') {
      liabilitiesMap.set(account, roundCurrency(balance));
    }

    if (balanceSheetType === 'equity') {
      equityMap.set(account, roundCurrency(balance));
    }
  });

  if (currentPeriodEarnings !== 0) {
    equityMap.set('Current Period Earnings', currentPeriodEarnings);
  }

  const assets = buildEntries(assetsMap);
  const liabilities = buildEntries(liabilitiesMap);
  const equity = buildEntries(equityMap);

  const assetsTotal = sumEntries(assets);
  const liabilitiesTotal = sumEntries(liabilities);
  const equityTotal = sumEntries(equity);

  return {
    assets,
    liabilities,
    equity,
    totals: {
      assetsTotal,
      liabilitiesTotal,
      equityTotal,
    },
    isBalanced: roundCurrency(assetsTotal) === roundCurrency(liabilitiesTotal + equityTotal),
  };
}
