import type { LedgerRow } from '../../app/upload/upload-utils';
import { isAnyExpense, isRevenueType } from './accounting';

export type CategoryVerification = {
  row: LedgerRow;
  description: string;
  assignedType: string;
  suggestedType: string | null;
  confidence: 'high' | 'medium' | 'low';
  mismatch: boolean;
  reason: string;
};

// Free keyword-based classification — no API needed
// Built from common accounting/bookkeeping knowledge

const INCOME_SIGNALS = [
  'payment received','invoice paid','customer payment','sales receipt','revenue',
  'income','refund received','grant','deposit from client','service fee received',
  'subscription revenue','royalty','commission earned',
];

const EXPENSE_SIGNALS = [
  'paid','bill','purchase','bought','fee','charge','subscription paid',
  'insurance','rent paid','utilities','payroll','salary paid','vendor payment',
  'supplies','maintenance','repair','advertising','marketing spend',
  'software subscription','membership','dues','postage','travel',
];

const ASSET_SIGNALS = [
  'equipment purchase','bought asset','capital expenditure','vehicle purchase',
  'computer purchase','furniture','building','land purchase','leasehold',
  'security deposit paid','loan disbursement received','transfer to savings',
];

const LIABILITY_SIGNALS = [
  'loan payment','loan received','credit card charge','credit card payment',
  'mortgage payment','accounts payable','borrowed','line of credit',
  'note payable','deferred revenue',
];

const COGS_SIGNALS = [
  'cost of goods','inventory purchase','raw material','product cost',
  'manufacturing','cost of sales','direct labor','direct material',
];

function scoreText(text: string, signals: string[]): number {
  const lower = text.toLowerCase();
  return signals.reduce((score, signal) => score + (lower.includes(signal) ? 1 : 0), 0);
}

function inferTypeFromText(text: string): string | null {
  if (!text.trim()) return null;

  const scores = {
    Income:   scoreText(text, INCOME_SIGNALS),
    Expense:  scoreText(text, EXPENSE_SIGNALS),
    Asset:    scoreText(text, ASSET_SIGNALS),
    Liability: scoreText(text, LIABILITY_SIGNALS),
    'Cost of Goods Sold': scoreText(text, COGS_SIGNALS),
  };

  const max = Math.max(...Object.values(scores));
  if (max === 0) return null;

  const best = Object.entries(scores).find(([, v]) => v === max);
  return best ? best[0] : null;
}

function normalizeType(t: string): string {
  const l = t.toLowerCase();
  if (l.includes('income') || l.includes('revenue') || l.includes('sales')) return 'Income';
  if (l.includes('cost of goods') || l.includes('cogs')) return 'Cost of Goods Sold';
  if (l.includes('expense')) return 'Expense';
  if (l.includes('asset') || l.includes('bank') || l.includes('receivable')) return 'Asset';
  if (l.includes('liabilit') || l.includes('payable')) return 'Liability';
  if (l.includes('equity')) return 'Equity';
  return t;
}

export function verifyCategoriesFree(rows: LedgerRow[]): CategoryVerification[] {
  const results: CategoryVerification[] = [];

  rows.forEach(row => {
    const description = [row.Description, row.Split, row.Name].filter(Boolean).join(' ').trim();
    if (!description) return;

    const assignedType = normalizeType(row['Distribution account type']);
    const suggestedType = inferTypeFromText(description);
    if (!suggestedType) return;

    // Only flag if suggestion conflicts with assignment
    const normalizedSuggested = normalizeType(suggestedType);

    // Define conflict: e.g. assigned=Expense but text strongly suggests Income
    const conflict =
      (assignedType === 'Income' && ['Expense','Cost of Goods Sold'].includes(normalizedSuggested)) ||
      (assignedType === 'Expense' && normalizedSuggested === 'Income') ||
      (assignedType === 'Asset' && normalizedSuggested === 'Liability') ||
      (assignedType === 'Liability' && normalizedSuggested === 'Asset');

    if (conflict) {
      results.push({
        row,
        description,
        assignedType,
        suggestedType: normalizedSuggested,
        confidence: 'medium',
        mismatch: true,
        reason: `Description suggests "${normalizedSuggested}" but assigned as "${assignedType}"`,
      });
    }
  });

  // Deduplicate by account + assignedType
  const seen = new Set<string>();
  return results.filter(r => {
    const key = `${r.row['Distribution account']}::${r.assignedType}::${r.suggestedType}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  }).slice(0, 20);
}
