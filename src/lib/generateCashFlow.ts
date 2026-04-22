import type { LedgerRow } from '../../app/upload/upload-utils';
import type { ProfitAndLoss } from './generatePL';

export type CashFlowAdjustment = {
  account: string;
  change: number;
  impact: number;
};

export type CashFlowStatement = {
  netProfit: number;
  adjustments: CashFlowAdjustment[];
  operatingCashFlow: number;
};

type AccountSnapshot = {
  type: 'asset' | 'liability';
  firstBalance: number;
  lastBalance: number;
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

const calculateImpact = (type: AccountSnapshot['type'], change: number) => {
  if (type === 'asset') {
    return -change;
  }

  return change;
};

export function generateCashFlow(rows: LedgerRow[], profitAndLoss: ProfitAndLoss): CashFlowStatement {
  const accountSnapshots = new Map<string, AccountSnapshot>();

  rows.forEach((row) => {
    const accountType = normalizeValue(row['Distribution account type']).toLowerCase();
    const account = normalizeValue(row['Distribution account']);
    const balance = parseNumericValue(row.Balance);

    if (!account || balance === null) {
      return;
    }

    if (accountType !== 'asset' && accountType !== 'liability') {
      return;
    }

    const existingSnapshot = accountSnapshots.get(account);
    if (existingSnapshot) {
      existingSnapshot.lastBalance = balance;
      return;
    }

    accountSnapshots.set(account, {
      type: accountType,
      firstBalance: balance,
      lastBalance: balance,
    });
  });

  const adjustments = Array.from(accountSnapshots.entries())
    .map(([account, snapshot]) => {
      const change = snapshot.lastBalance - snapshot.firstBalance;
      const impact = calculateImpact(snapshot.type, change);

      return {
        account,
        change,
        impact,
      };
    })
    .filter((adjustment) => adjustment.change !== 0)
    .sort((left, right) => left.account.localeCompare(right.account));

  const operatingCashFlow =
    profitAndLoss.netProfit + adjustments.reduce((total, adjustment) => total + adjustment.impact, 0);

  return {
    netProfit: profitAndLoss.netProfit,
    adjustments,
    operatingCashFlow,
  };
}
