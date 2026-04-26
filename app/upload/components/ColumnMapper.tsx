'use client';

import { useState, useCallback, type DragEvent, type ChangeEvent } from 'react';
import { requiredHeaders, CORE_MANDATORY, type HeaderKey } from '../upload-utils';
import type { UploadTheme } from './FileDropzone';

type Props = {
  theme: UploadTheme;
  fileHeaders: string[];          // headers from the uploaded file
  onMappingConfirmed: (mapping: Record<HeaderKey, string>) => void;
  onCancel: () => void;
};

const HEADER_DESCRIPTIONS: Record<HeaderKey, string> = {
  'Distribution account':      'The account name (e.g. Checking, Sales Revenue)',
  'Distribution account type': 'Account category: Asset, Liability, Equity, Income, Expense',
  'Transaction date':          'Date of transaction (e.g. 01/15/2024)',
  'Transaction type':          'Type: Invoice, Payment, Bill, Deposit, etc.',
  'Num':                       'Transaction or check number',
  'Name':                      'Customer or vendor name',
  'Description':               'Memo or description of transaction',
  'Split':                     'Offsetting account for double-entry',
  'Amount':                    'Transaction amount (positive or negative)',
  'Balance':                   'Running account balance after transaction',
};

const isMandatory = (h: HeaderKey) => CORE_MANDATORY.includes(h as typeof CORE_MANDATORY[number]);

// Auto-suggest best match from file headers
const autoSuggest = (target: HeaderKey, fileHeaders: string[]): string => {
  const aliases: Partial<Record<HeaderKey, string[]>> = {
    'Distribution account':      ['account', 'acct', 'gl account', 'ledger account', 'account name'],
    'Distribution account type': ['account type', 'type', 'account category', 'category', 'acct type'],
    'Transaction date':          ['date', 'txn date', 'trans date', 'posting date', 'entry date'],
    'Transaction type':          ['transaction type', 'type', 'txn type', 'trans type', 'entry type'],
    'Num':                       ['num', 'number', 'ref', 'reference', 'check no', 'doc no', 'id'],
    'Name':                      ['name', 'vendor', 'customer', 'payee', 'party', 'entity'],
    'Description':               ['description', 'memo', 'details', 'notes', 'narrative', 'desc'],
    'Split':                     ['split', 'offset', 'contra', 'opposite account'],
    'Amount':                    ['amount', 'debit', 'credit', 'value', 'amt'],
    'Balance':                   ['balance', 'running balance', 'bal', 'ending balance'],
  };

  const candidates = aliases[target] ?? [];
  for (const alias of candidates) {
    const match = fileHeaders.find(h => h.toLowerCase().includes(alias.toLowerCase()) || alias.toLowerCase().includes(h.toLowerCase()));
    if (match) return match;
  }
  return '';
};

export function ColumnMapper({ theme, fileHeaders, onMappingConfirmed, onCancel }: Props) {
  const d = theme === 'dark';

  const [mapping, setMapping] = useState<Record<HeaderKey, string>>(() => {
    const initial = {} as Record<HeaderKey, string>;
    requiredHeaders.forEach(h => { initial[h] = autoSuggest(h, fileHeaders); });
    return initial;
  });

  const [dragOver, setDragOver] = useState<HeaderKey | null>(null);
  const [dragSource, setDragSource] = useState<string | null>(null);

  const unmappedHeaders = fileHeaders.filter(h => !Object.values(mapping).includes(h));
  const mandatoryMapped = CORE_MANDATORY.every(h => mapping[h as HeaderKey]);

  const handleDrop = useCallback((e: DragEvent<HTMLDivElement>, target: HeaderKey) => {
    e.preventDefault();
    const source = e.dataTransfer.getData('text/plain');
    if (source) setMapping((m: Record<HeaderKey, string>) => ({ ...m, [target]: source }));
    setDragOver(null);
    setDragSource(null);
  }, []);

  const handleDragStart = useCallback((e: DragEvent<HTMLDivElement>, header: string) => {
    e.dataTransfer.setData('text/plain', header);
    setDragSource(header);
  }, []);

  const clearMapping = (key: HeaderKey) => setMapping((m: Record<HeaderKey, string>) => ({ ...m, [key]: '' }));

  // panel/card styles
  const panel   = d ? 'bg-slate-900 border-slate-700' : 'bg-white border-slate-200';
  const cardBg  = d ? 'bg-slate-800' : 'bg-slate-50';
  const txt     = d ? 'text-slate-100' : 'text-slate-900';
  const muted   = d ? 'text-slate-400' : 'text-slate-500';
  const chipBg  = d ? 'bg-slate-700 text-slate-200 border-slate-600' : 'bg-white text-slate-700 border-slate-300';
  const dropZoneBase = `rounded-xl border-2 border-dashed px-3 py-2 text-xs transition-all duration-200 min-h-[38px] flex items-center gap-2`;

  return (
    <div className={`rounded-3xl border p-6 ${panel}`}>
      {/* Header */}
      <div className="flex items-start justify-between gap-4 mb-6">
        <div>
          <div className="flex items-center gap-2">
            <span className="text-2xl">🗂️</span>
            <h2 className={`text-lg font-bold ${txt}`}>Map Your Columns</h2>
          </div>
          <p className={`mt-1 text-sm ${muted}`}>
            Your file headers don't match exactly. Drag columns from the right onto the fields below, or use the dropdowns.
          </p>
        </div>
        <button type="button" onClick={onCancel} className={`text-xs px-3 py-1.5 rounded-lg transition-colors ${d ? 'bg-slate-800 text-slate-400 hover:bg-slate-700' : 'bg-slate-100 text-slate-500 hover:bg-slate-200'}`}>Cancel</button>
      </div>

      <div className="flex flex-col lg:flex-row gap-6">
        {/* Left: mapping fields */}
        <div className="flex-1 space-y-2">
          <p className={`text-xs font-semibold uppercase tracking-widest mb-3 ${muted}`}>Required Fields</p>
          {requiredHeaders.map(key => {
            const mandatory = isMandatory(key);
            const mapped = mapping[key];
            const isOver = dragOver === key;

            return (
              <div key={key} className={`rounded-xl p-3 ${cardBg} border ${mandatory ? d ? 'border-cyan-800/50' : 'border-cyan-200' : d ? 'border-slate-700' : 'border-slate-200'}`}>
                <div className="flex items-start justify-between gap-2 mb-1.5">
                  <div className="flex items-center gap-1.5">
                    {mandatory && <span className="text-[9px] font-bold bg-cyan-500/20 text-cyan-400 px-1.5 py-0.5 rounded-full">✱ Required</span>}
                    <span className={`text-xs font-semibold ${txt}`}>{key}</span>
                  </div>
                  {mapped && (
                    <button type="button" onClick={() => clearMapping(key)} className="text-[9px] text-rose-400 hover:text-rose-300 transition-colors">✕ clear</button>
                  )}
                </div>
                <p className={`text-[10px] mb-2 ${muted}`}>{HEADER_DESCRIPTIONS[key]}</p>

                {/* Drop zone */}
                <div
                  onDragOver={e => { e.preventDefault(); setDragOver(key); }}
                  onDragLeave={() => setDragOver(null)}
                  onDrop={e => handleDrop(e, key)}
                  className={`${dropZoneBase} ${
                    isOver ? 'border-cyan-400 bg-cyan-500/10 scale-[1.01]' :
                    mapped ? d ? 'border-emerald-600 bg-emerald-900/30' : 'border-emerald-400 bg-emerald-50' :
                    mandatory ? d ? 'border-amber-600/50 bg-amber-900/10' : 'border-amber-300 bg-amber-50' :
                    d ? 'border-slate-600 bg-slate-900/30' : 'border-slate-300 bg-white'
                  }`}
                >
                  {mapped ? (
                    <span className={`text-xs font-semibold ${d ? 'text-emerald-300' : 'text-emerald-700'}`}>✓ {mapped}</span>
                  ) : (
                    <>
                      <span className={`text-[10px] ${isOver ? 'text-cyan-300' : mandatory ? d ? 'text-amber-500' : 'text-amber-600' : muted}`}>
                        {isOver ? '📌 Drop here' : '← Drag a column here'}
                      </span>
                    </>
                  )}
                </div>

                {/* Dropdown fallback */}
                <select
                  value={mapping[key]}
                  onChange={e => setMapping((m: Record<HeaderKey, string>) => ({ ...m, [key]: e.target.value }))}
                  className={`mt-2 w-full text-[10px] rounded-lg px-2 py-1.5 border outline-none transition-colors ${d ? 'bg-slate-900 border-slate-600 text-slate-300' : 'bg-white border-slate-300 text-slate-700'}`}
                >
                  <option value="">— select column —</option>
                  {fileHeaders.map(h => (
                    <option key={h} value={h}>{h}</option>
                  ))}
                </select>
              </div>
            );
          })}
        </div>

        {/* Right: draggable file headers */}
        <div className="lg:w-56 flex-shrink-0">
          <p className={`text-xs font-semibold uppercase tracking-widest mb-3 ${muted}`}>Your File's Columns</p>
          <div className={`rounded-xl p-3 border ${cardBg} ${d ? 'border-slate-700' : 'border-slate-200'} sticky top-4`}>
            <p className={`text-[10px] mb-3 ${muted}`}>Drag these onto the fields ←</p>
            <div className="space-y-1.5">
              {fileHeaders.map(h => {
                const isUsed = Object.values(mapping).includes(h);
                return (
                  <div
                    key={h}
                    draggable={!isUsed}
                    onDragStart={e => handleDragStart(e, h)}
                    onDragEnd={() => setDragSource(null)}
                    className={`px-3 py-2 rounded-lg border text-xs font-medium cursor-grab active:cursor-grabbing transition-all select-none ${
                      isUsed
                        ? d ? 'bg-slate-800/50 border-slate-700 text-slate-600 line-through' : 'bg-slate-100 border-slate-200 text-slate-400 line-through'
                        : dragSource === h
                        ? 'border-cyan-400 bg-cyan-500/20 text-cyan-300 scale-95 rotate-1 shadow-lg'
                        : `${chipBg} border hover:border-cyan-400 hover:shadow-md`
                    }`}
                  >
                    <span className="mr-1">{isUsed ? '✓' : '⠿'}</span>
                    {h}
                  </div>
                );
              })}
            </div>

            {/* Unmapped count */}
            {unmappedHeaders.length > 0 && (
              <p className={`mt-3 text-[10px] ${muted}`}>{unmappedHeaders.length} column{unmappedHeaders.length !== 1 ? 's' : ''} not yet mapped</p>
            )}
          </div>
        </div>
      </div>

      {/* Confirm */}
      <div className={`mt-6 pt-4 border-t flex items-center justify-between gap-4 ${d ? 'border-slate-700' : 'border-slate-200'}`}>
        <div>
          {!mandatoryMapped && (
            <p className="text-xs text-amber-400">⚠️ Map all ✱ Required fields to continue</p>
          )}
          {mandatoryMapped && (
            <p className="text-xs text-emerald-400">✅ Required fields mapped — ready to process</p>
          )}
        </div>
        <button
          type="button"
          disabled={!mandatoryMapped}
          onClick={() => onMappingConfirmed(mapping)}
          className={`px-5 py-2.5 rounded-xl text-sm font-semibold transition-all ${mandatoryMapped ? 'bg-cyan-600 text-white hover:bg-cyan-500 shadow-lg shadow-cyan-600/30' : 'bg-slate-700 text-slate-500 cursor-not-allowed'}`}
        >
          Process with this mapping →
        </button>
      </div>
    </div>
  );
}
