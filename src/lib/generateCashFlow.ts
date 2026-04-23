import type { LedgerRow } from '../../app/upload/upload-utils';
import { classifyBalanceSheetType, parseCurrencyAmount, roundCurrency } from './accounting';
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

const calculateImpact = (type: AccountSnapshot['type'], change: number) => {
  if (type === 'asset') {
    return -change;
  }

  return change;
};

export function generateCashFlow(rows: LedgerRow[], profitAndLoss: ProfitAndLoss): CashFlowStatement {
  const accountSnapshots = new Map<string, AccountSnapshot>();

  rows.forEach((row) => {
    const accountType = classifyBalanceSheetType(row['Distribution account type']);
    const account = normalizeValue(row['Distribution account']);
    const balance = parseCurrencyAmount(row.Balance);

    if (!account || balance === null) {
      return;
    }

    if (accountType !== 'asset' && accountType !== 'liability') {
      return;
    }

    const existingSnapshot = accountSnapshots.get(account);
    if (existingSnapshot) {
      existingSnapshot.lastBalance = roundCurrency(balance);
      return;
    }

    accountSnapshots.set(account, {
      type: accountType,
      firstBalance: roundCurrency(balance),
      lastBalance: roundCurrency(balance),
    });
  });

  const adjustments = Array.from(accountSnapshots.entries())
    .map(([account, snapshot]) => {
      const change = roundCurrency(snapshot.lastBalance - snapshot.firstBalance);
      const impact = roundCurrency(calculateImpact(snapshot.type, change));

      return {
        account,
        change,
        impact,
      };
    })
    .filter((adjustment) => adjustment.change !== 0)
    .sort((left, right) => left.account.localeCompare(right.account));

  const operatingCashFlow =
    roundCurrency(profitAndLoss.netProfit + adjustments.reduce((total, adjustment) => total + adjustment.impact, 0));

  return {
    netProfit: profitAndLoss.netProfit,
    adjustments,
    operatingCashFlow,
  };
}
