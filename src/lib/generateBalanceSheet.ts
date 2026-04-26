import type { LedgerRow } from '../../app/upload/upload-utils';
import {
  BS_TOLERANCE, classifyBalanceSheetType,
  isCurrentAsset, isCurrentLiability,
  isCogsType, isExpenseType, isRevenueType,
  parseCurrencyAmount, roundCurrency,
  type BalanceSheetCategory,
} from './accounting';

export type BalanceSheetEntry = { account: string; value: number; isCurrent: boolean };

export type BalanceSheet = {
  assets: BalanceSheetEntry[];
  liabilities: BalanceSheetEntry[];
  equity: BalanceSheetEntry[];
  totals: {
    assetsTotal: number;
    currentAssetsTotal: number;
    nonCurrentAssetsTotal: number;
    liabilitiesTotal: number;
    currentLiabilitiesTotal: number;
    nonCurrentLiabilitiesTotal: number;
    equityTotal: number;
  };
  ratios: {
    currentRatio: number | null;      // current assets / current liabilities
    quickRatio: number | null;        // (current assets - inventory) / current liabilities
    debtToEquity: number | null;      // total liabilities / total equity
    debtRatio: number | null;         // total liabilities / total assets
  };
  isBalanced: boolean;
  variance: number;
  currentPeriodEarnings: number;
};

export function generateBalanceSheet(rows: LedgerRow[]): BalanceSheet {
  const lastBalance = new Map<string, { value: number; bsType: BalanceSheetCategory; accountType: string }>();
  let cpe = 0;
  // Track if ledger already has a net income equity account (to avoid double-counting)
  let hasNetIncomeEquity = false;

  rows.forEach((row) => {
    const accountType = row['Distribution account type'].trim();
    const account     = row['Distribution account'].trim();
    const amount      = parseCurrencyAmount(row.Amount) ?? 0;
    const bsType      = classifyBalanceSheetType(accountType);

    // Check if equity accounts already contain net income
    if (bsType === 'equity') {
      const al = account.toLowerCase();
      if (al.includes('net income') || al.includes('net profit') || al.includes('current year')) {
        hasNetIncomeEquity = true;
      }
    }

    if (isRevenueType(accountType)) cpe = roundCurrency(cpe + amount);
    else if (isCogsType(accountType) || isExpenseType(accountType)) cpe = roundCurrency(cpe - amount);

    if (!account || !bsType) return;
    const balRaw = row.Balance.trim();
    const bal = balRaw === '' ? 0 : parseCurrencyAmount(balRaw);
    if (bal === null) return;

    lastBalance.set(account, { value: roundCurrency(bal), bsType, accountType });
  });

  const assetsMap     = new Map<string, BalanceSheetEntry>();
  const liabMap       = new Map<string, BalanceSheetEntry>();
  const equityMap     = new Map<string, BalanceSheetEntry>();

  lastBalance.forEach(({ value, bsType, accountType }, account) => {
    const entry: BalanceSheetEntry = { account, value, isCurrent: bsType === 'asset' ? isCurrentAsset(accountType) : isCurrentLiability(accountType) };
    if (bsType === 'asset')     assetsMap.set(account, entry);
    if (bsType === 'liability') liabMap.set(account, entry);
    if (bsType === 'equity')    equityMap.set(account, { ...entry, isCurrent: false });
  });

  // Only inject CPE if not already captured in equity
  if (cpe !== 0 && !hasNetIncomeEquity) {
    equityMap.set('Current Period Earnings', { account: 'Current Period Earnings', value: cpe, isCurrent: false });
  }

  const assets      = Array.from(assetsMap.values()).sort((a, b) => a.account.localeCompare(b.account));
  const liabilities = Array.from(liabMap.values()).sort((a, b) => a.account.localeCompare(b.account));
  const equity      = Array.from(equityMap.values()).sort((a, b) => a.account.localeCompare(b.account));

  const sum = (arr: BalanceSheetEntry[]) => roundCurrency(arr.reduce((t, e) => t + e.value, 0));
  const sumIf = (arr: BalanceSheetEntry[], flag: boolean) => roundCurrency(arr.filter(e => e.isCurrent === flag).reduce((t, e) => t + e.value, 0));

  const assetsTotal              = sum(assets);
  const currentAssetsTotal       = sumIf(assets, true);
  const nonCurrentAssetsTotal    = sumIf(assets, false);
  const liabilitiesTotal         = sum(liabilities);
  const currentLiabilitiesTotal  = sumIf(liabilities, true);
  const nonCurrentLiabilitiesTotal = sumIf(liabilities, false);
  const equityTotal              = sum(equity);

  const variance = roundCurrency(assetsTotal - (liabilitiesTotal + equityTotal));

  // Financial ratios
  const inventoryValue = assets.filter(e => e.account.toLowerCase().includes('inventory')).reduce((t, e) => t + e.value, 0);
  const currentRatio   = currentLiabilitiesTotal !== 0 ? roundCurrency(currentAssetsTotal / currentLiabilitiesTotal) : null;
  const quickRatio     = currentLiabilitiesTotal !== 0 ? roundCurrency((currentAssetsTotal - inventoryValue) / currentLiabilitiesTotal) : null;
  const debtToEquity   = equityTotal !== 0 ? roundCurrency(liabilitiesTotal / equityTotal) : null;
  const debtRatio      = assetsTotal !== 0 ? roundCurrency(liabilitiesTotal / assetsTotal) : null;

  return {
    assets, liabilities, equity,
    totals: { assetsTotal, currentAssetsTotal, nonCurrentAssetsTotal, liabilitiesTotal, currentLiabilitiesTotal, nonCurrentLiabilitiesTotal, equityTotal },
    ratios: { currentRatio, quickRatio, debtToEquity, debtRatio },
    isBalanced: Math.abs(variance) <= BS_TOLERANCE,
    variance,
    currentPeriodEarnings: cpe,
  };
}
