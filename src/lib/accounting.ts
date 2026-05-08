export type BalanceSheetCategory = 'asset' | 'liability' | 'equity';

export function parseCurrencyAmount(value: string): number | null {
  const s = value.trim();
  if (!s) return null;
  const isNeg = (s.startsWith('(') && s.endsWith(')')) || s.startsWith('-');
  const num = Number(s.replace(/[,$() -]/g, ''));
  if (!Number.isFinite(num) || s.replace(/[,$().\d\s-]/g, '').length > 0) return null;
  return isNeg ? -num : num;
}

export function roundCurrency(v: number): number {
  return Math.round((v + Number.EPSILON) * 100) / 100;
}

export var BS_TOLERANCE = 1.0;

// ─── QBO account type sets (var to avoid TDZ in production) ──────────────────
var ASSET_TYPES = new Set(['bank','accounts receivable (a/r)','accounts receivable','other current assets','other current asset','fixed assets','fixed asset','other assets','other asset','inventory','asset','assets']);
var CURRENT_ASSET_TYPES = new Set(['bank','accounts receivable (a/r)','accounts receivable','other current assets','other current asset','inventory']);
var NONCURRENT_ASSET_TYPES = new Set(['fixed assets','fixed asset','other assets','other asset']);
var LIABILITY_TYPES = new Set(['accounts payable (a/p)','accounts payable','credit card','other current liabilities','other current liability','long term liabilities','long-term liabilities','long term liability','other liability','other liabilities','liability','liabilities']);
var CURRENT_LIABILITY_TYPES = new Set(['accounts payable (a/p)','accounts payable','credit card','other current liabilities','other current liability']);
var NONCURRENT_LIABILITY_TYPES = new Set(['long term liabilities','long-term liabilities','long term liability']);
var EQUITY_TYPES = new Set(['equity','equities','retained earnings','opening balance equity',"owner's equity",'partner equity','partner contribution','partner distribution','stockholders equity','shareholder equity']);
var INCOME_TYPES = new Set(['income','revenue','sales','other income','other revenue','non-operating income']);
var COGS_TYPES = new Set(['cost of goods sold','cogs','cost of sales']);
var EXPENSE_TYPES = new Set(['expense','expenses','other expense','other expenses','non-operating expense']);

function _match(set: Set<string>, kws: string[], v: string): boolean {
  const t = v.trim().toLowerCase();
  return set.has(t) || kws.some(k => t.includes(k));
}

export function classifyBalanceSheetType(v: string): BalanceSheetCategory | null {
  const t = v.trim().toLowerCase();
  if (ASSET_TYPES.has(t) || ['receivable','prepaid'].some(k => t.includes(k))) return 'asset';
  if (LIABILITY_TYPES.has(t) || ['liabilit','payable','credit card','loan','mortgage'].some(k => t.includes(k))) return 'liability';
  if (EQUITY_TYPES.has(t) || ['equity','capital','contribution','retained'].some(k => t.includes(k))) return 'equity';
  return null;
}

export function isCurrentAsset(v: string): boolean {
  const t = v.trim().toLowerCase();
  return CURRENT_ASSET_TYPES.has(t) || ['bank','receivable','prepaid','cash','inventory'].some(k => t.includes(k));
}

export function isNonCurrentAsset(v: string): boolean {
  const t = v.trim().toLowerCase();
  return NONCURRENT_ASSET_TYPES.has(t) || ['fixed','equipment','furniture','vehicle','building','land'].some(k => t.includes(k));
}

export function isCurrentLiability(v: string): boolean {
  const t = v.trim().toLowerCase();
  return CURRENT_LIABILITY_TYPES.has(t) || ['payable','credit card'].some(k => t.includes(k));
}

export function isNonCurrentLiability(v: string): boolean {
  const t = v.trim().toLowerCase();
  return NONCURRENT_LIABILITY_TYPES.has(t) || ['long term','long-term','mortgage','note payable'].some(k => t.includes(k));
}

export function isRevenueType(v: string): boolean {
  return _match(INCOME_TYPES, ['income','revenue','sales'], v);
}

export function isCogsType(v: string): boolean {
  return _match(COGS_TYPES, ['cost of goods','cogs','cost of sales'], v);
}

export function isExpenseType(v: string): boolean {
  return _match(EXPENSE_TYPES, ['expense'], v) && !isCogsType(v);
}

export function isAnyExpense(v: string): boolean {
  return isExpenseType(v) || isCogsType(v);
}
