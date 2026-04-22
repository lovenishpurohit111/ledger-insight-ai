import type { LedgerRow } from '../../app/upload/upload-utils';

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

const parseNumericValue = (value: string) => {
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

const buildEntries = (accounts: Map<string, number>) =>
  Array.from(accounts.entries())
    .map(([account, value]) => ({ account, value }))
    .sort((left, right) => left.account.localeCompare(right.account));

const sumEntries = (entries: BalanceSheetEntry[]) => entries.reduce((total, entry) => total + entry.value, 0);

export function generateBalanceSheet(rows: LedgerRow[]): BalanceSheet {
  const assetsMap = new Map<string, number>();
  const liabilitiesMap = new Map<string, number>();
  const equityMap = new Map<string, number>();

  rows.forEach((row) => {
    const accountType = normalizeValue(row['Distribution account type']).toLowerCase();
    const account = normalizeValue(row['Distribution account']);
    const balance = parseNumericValue(row.Balance);

    if (!account || balance === null) {
      return;
    }

    if (accountType === 'asset') {
      assetsMap.set(account, balance);
    }

    if (accountType === 'liability') {
      liabilitiesMap.set(account, balance);
    }

    if (accountType === 'equity') {
      equityMap.set(account, balance);
    }
  });

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
    isBalanced: assetsTotal === liabilitiesTotal + equityTotal,
  };
}
