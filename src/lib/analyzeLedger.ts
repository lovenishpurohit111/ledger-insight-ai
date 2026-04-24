import { isExpenseType, isRevenueType, classifyBalanceSheetType } from './accounting';
import type { LedgerRow } from '../../app/upload/upload-utils';

export type InconsistentVendor = {
  vendor: string;
  accounts: string[];
  reason: string;
};

export type DuplicateTransaction = {
  name: string;
  amount: string;
  transactionDate: string;
  occurrences: number;
};

export type LedgerAnalysis = {
  inconsistentVendors: InconsistentVendor[];
  duplicates: DuplicateTransaction[];
  totalTransactions: number;
};

// Broad account "family" — debit/credit pairs naturally span TWO families (e.g. Expense + Asset).
// Inconsistency = same vendor posting to MULTIPLE accounts within the SAME family.
const getAccountFamily = (accountType: string): string => {
  if (isRevenueType(accountType)) return 'income';
  if (isExpenseType(accountType)) return 'expense';
  const bsType = classifyBalanceSheetType(accountType);
  return bsType ?? 'other';
};

const normalizeValue = (v: string) => v.trim();

export function analyzeLedger(rows: LedgerRow[]): LedgerAnalysis {
  // vendor → family → Set of account names
  const vendorFamilyAccounts = new Map<string, Map<string, Set<string>>>();
  const duplicateCounts = new Map<string, DuplicateTransaction>();

  rows.forEach((row) => {
    const vendor = normalizeValue(row.Name);
    const account = normalizeValue(row['Distribution account']);
    const accountType = normalizeValue(row['Distribution account type']);
    const amount = normalizeValue(row.Amount);
    const transactionDate = normalizeValue(row['Transaction date']);

    if (vendor && account && accountType) {
      const family = getAccountFamily(accountType);
      const familyMap = vendorFamilyAccounts.get(vendor) ?? new Map<string, Set<string>>();
      const acctSet = familyMap.get(family) ?? new Set<string>();
      acctSet.add(account);
      familyMap.set(family, acctSet);
      vendorFamilyAccounts.set(vendor, familyMap);
    }

    if (!vendor || !amount || !transactionDate) return;
    const key = [vendor, amount, transactionDate].join('::');
    const existing = duplicateCounts.get(key);
    if (existing) { existing.occurrences += 1; return; }
    duplicateCounts.set(key, { name: vendor, amount, transactionDate, occurrences: 1 });
  });

  const inconsistentVendors: InconsistentVendor[] = [];

  vendorFamilyAccounts.forEach((familyMap, vendor) => {
    familyMap.forEach((acctSet, family) => {
      // Only flag if vendor uses 2+ different accounts within the same family
      if (acctSet.size > 1) {
        inconsistentVendors.push({
          vendor,
          accounts: Array.from(acctSet).sort(),
          reason: `Uses ${acctSet.size} different ${family} accounts`,
        });
      }
    });
  });

  inconsistentVendors.sort((a, b) => a.vendor.localeCompare(b.vendor));

  const duplicates = Array.from(duplicateCounts.values()).filter((e) => e.occurrences > 1);

  return {
    inconsistentVendors,
    duplicates,
    totalTransactions: rows.length,
  };
}
