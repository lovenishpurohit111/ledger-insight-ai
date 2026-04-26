export type BalanceSheetCategory = 'asset' | 'liability' | 'equity';

export const parseCurrencyAmount = (value: string): number | null => {
  const s = value.trim();
  if (!s) return null;
  const isNeg = (s.startsWith('(') && s.endsWith(')')) || s.startsWith('-');
  const num = Number(s.replace(/[,$() -]/g, ''));
  if (!Number.isFinite(num) || s.replace(/[,$().\d\s-]/g, '').length > 0) return null;
  return isNeg ? -num : num;
};

export const roundCurrency = (v: number) => Math.round((v + Number.EPSILON) * 100) / 100;
export const BS_TOLERANCE = 1.0;

// ─── QBO account type sets ────────────────────────────────────────────────────
const ASSET_TYPES = new Set(['bank','accounts receivable (a/r)','accounts receivable','other current assets','other current asset','fixed assets','fixed asset','other assets','other asset','inventory','asset','assets']);
const CURRENT_ASSET_TYPES = new Set(['bank','accounts receivable (a/r)','accounts receivable','other current assets','other current asset','inventory']);
const NONCURRENT_ASSET_TYPES = new Set(['fixed assets','fixed asset','other assets','other asset']);

const LIABILITY_TYPES = new Set(['accounts payable (a/p)','accounts payable','credit card','other current liabilities','other current liability','long term liabilities','long-term liabilities','long term liability','other liability','other liabilities','liability','liabilities']);
const CURRENT_LIABILITY_TYPES = new Set(['accounts payable (a/p)','accounts payable','credit card','other current liabilities','other current liability']);
const NONCURRENT_LIABILITY_TYPES = new Set(['long term liabilities','long-term liabilities','long term liability']);

const EQUITY_TYPES = new Set(['equity','equities','retained earnings','opening balance equity',"owner's equity",'partner equity','partner contribution','partner distribution','stockholders equity','shareholder equity']);
const INCOME_TYPES = new Set(['income','revenue','sales','other income','other revenue','non-operating income']);
const COGS_TYPES = new Set(['cost of goods sold','cogs','cost of sales']);
const EXPENSE_TYPES = new Set(['expense','expenses','other expense','other expenses','non-operating expense']);

const match = (set: Set<string>, kws: string[], v: string) => {
  const t = v.trim().toLowerCase();
  return set.has(t) || kws.some(k => t.includes(k));
};

export const classifyBalanceSheetType = (v: string): BalanceSheetCategory | null => {
  const t = v.trim().toLowerCase();
  if (ASSET_TYPES.has(t) || ['receivable','prepaid'].some(k => t.includes(k))) return 'asset';
  if (LIABILITY_TYPES.has(t) || ['liabilit','payable','credit card','loan','mortgage'].some(k => t.includes(k))) return 'liability';
  if (EQUITY_TYPES.has(t) || ['equity','capital','contribution','retained'].some(k => t.includes(k))) return 'equity';
  return null;
};

export const isCurrentAsset     = (v: string) => { const t = v.trim().toLowerCase(); return CURRENT_ASSET_TYPES.has(t) || ['bank','receivable','prepaid','cash','inventory'].some(k => t.includes(k)); };
export const isNonCurrentAsset  = (v: string) => { const t = v.trim().toLowerCase(); return NONCURRENT_ASSET_TYPES.has(t) || ['fixed','equipment','furniture','vehicle','building','land'].some(k => t.includes(k)); };
export const isCurrentLiability = (v: string) => { const t = v.trim().toLowerCase(); return CURRENT_LIABILITY_TYPES.has(t) || ['payable','credit card'].some(k => t.includes(k)); };
export const isNonCurrentLiability = (v: string) => { const t = v.trim().toLowerCase(); return NONCURRENT_LIABILITY_TYPES.has(t) || ['long term','long-term','mortgage','note payable'].some(k => t.includes(k)); };

export const isRevenueType = (v: string) => match(INCOME_TYPES, ['income','revenue','sales'], v);
export const isCogsType    = (v: string) => match(COGS_TYPES,   ['cost of goods','cogs','cost of sales'], v);
export const isExpenseType = (v: string) => match(EXPENSE_TYPES, ['expense'], v) && !isCogsType(v);
export const isAnyExpense  = (v: string) => isExpenseType(v) || isCogsType(v);
