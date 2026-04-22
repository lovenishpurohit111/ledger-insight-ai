import type { LedgerRow } from '../../app/upload/upload-utils';

export type InconsistentVendor = {
  vendor: string;
  accounts: string[];
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

const normalizeValue = (value: string) => value.trim();

export function analyzeLedger(rows: LedgerRow[]): LedgerAnalysis {
  const vendorAccounts = new Map<string, Set<string>>();
  const duplicateCounts = new Map<string, DuplicateTransaction>();

  rows.forEach((row) => {
    const vendor = normalizeValue(row.Name);
    const account = normalizeValue(row['Distribution account']);
    const amount = normalizeValue(row.Amount);
    const transactionDate = normalizeValue(row['Transaction date']);

    if (vendor) {
      const accounts = vendorAccounts.get(vendor) ?? new Set<string>();
      if (account) {
        accounts.add(account);
      }
      vendorAccounts.set(vendor, accounts);
    }

    const duplicateKey = [vendor, amount, transactionDate].join('::');
    const existingDuplicate = duplicateCounts.get(duplicateKey);

    if (existingDuplicate) {
      existingDuplicate.occurrences += 1;
      return;
    }

    duplicateCounts.set(duplicateKey, {
      name: vendor,
      amount,
      transactionDate,
      occurrences: 1,
    });
  });

  const inconsistentVendors = Array.from(vendorAccounts.entries())
    .filter(([, accounts]) => accounts.size > 1)
    .map(([vendor, accounts]) => ({
      vendor,
      accounts: Array.from(accounts).sort(),
    }))
    .sort((left, right) => left.vendor.localeCompare(right.vendor));

  const duplicates = Array.from(duplicateCounts.values()).filter((entry) => entry.occurrences > 1);

  return {
    inconsistentVendors,
    duplicates,
    totalTransactions: rows.length,
  };
}
