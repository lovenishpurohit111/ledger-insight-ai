'use client';

import { useState } from 'react';
import type { UploadTheme } from './FileDropzone';

type Props = { theme: UploadTheme };

const STEPS = [
  {
    id: 1,
    icon: '📊',
    title: 'Open Reports',
    subtitle: 'Navigate to Reports in QBO',
    description: 'In QuickBooks Online, click "Reports" in the left sidebar.',
    visual: (
      <div className="relative w-full h-48 rounded-xl overflow-hidden bg-[#1a3a52] border border-[#2d6a9f]/40 font-mono text-xs">
        {/* QBO sidebar mockup */}
        <div className="absolute left-0 top-0 bottom-0 w-28 bg-[#0c2a3e] flex flex-col gap-1 p-2">
          {['Dashboard','Banking','Sales','Expenses','Payroll','Reports','Taxes','Accounting'].map((item, i) => (
            <div key={item} className={`rounded px-2 py-1.5 text-[10px] font-medium cursor-default transition-all ${item === 'Reports' ? 'bg-[#2d6a9f] text-white shadow-lg shadow-[#2d6a9f]/40' : 'text-slate-400 hover:text-slate-200'}`}>
              {item === 'Reports' && <span className="mr-1">›</span>}{item}
            </div>
          ))}
        </div>
        <div className="absolute left-28 top-0 right-0 bottom-0 flex items-center justify-center">
          <div className="text-center">
            <div className="text-2xl mb-2">📋</div>
            <div className="text-slate-300 text-xs">Reports Center</div>
          </div>
        </div>
        {/* Animated highlight ring */}
        <div className="absolute left-1 top-[86px] w-26 h-7 rounded border-2 border-cyan-400 animate-pulse" style={{width:'108px'}} />
      </div>
    ),
  },
  {
    id: 2,
    icon: '🔍',
    title: 'Find General Ledger',
    subtitle: 'Search in the Reports list',
    description: 'In the search box at the top of Reports, type "General Ledger" and select it.',
    visual: (
      <div className="relative w-full h-48 rounded-xl overflow-hidden bg-[#1a3a52] border border-[#2d6a9f]/40">
        <div className="p-4">
          <div className="flex items-center gap-2 bg-[#0c2a3e] rounded-lg px-3 py-2 border border-[#2d6a9f]/60">
            <span className="text-slate-400 text-xs">🔍</span>
            <span className="text-cyan-300 text-xs font-mono">General Ledger</span>
            <span className="animate-pulse text-cyan-400">|</span>
          </div>
          <div className="mt-2 bg-[#0c2a3e] rounded-lg border border-[#2d6a9f]/40 overflow-hidden">
            {['General Ledger', 'General Journal', 'General Ledger Detail'].map((r, i) => (
              <div key={r} className={`px-3 py-2 text-xs flex items-center gap-2 ${i === 0 ? 'bg-[#2d6a9f]/40 text-white' : 'text-slate-400'}`}>
                {i === 0 && <span className="text-cyan-400">✓</span>}
                <span>{r}</span>
              </div>
            ))}
          </div>
        </div>
        <div className="absolute bottom-3 right-3 bg-cyan-500 text-white text-[10px] px-2 py-1 rounded-full animate-bounce">Select this ↑</div>
      </div>
    ),
  },
  {
    id: 3,
    icon: '📅',
    title: 'Set Date Range',
    subtitle: 'Choose your reporting period',
    description: 'Set the date range (e.g. This Year, Last Year, or Custom). Then click "Run Report".',
    visual: (
      <div className="relative w-full h-48 rounded-xl overflow-hidden bg-[#1a3a52] border border-[#2d6a9f]/40 p-4">
        <div className="text-[11px] text-slate-300 mb-2 font-semibold">General Ledger</div>
        <div className="flex gap-2 flex-wrap">
          {['Report period', 'From', 'To'].map((label, i) => (
            <div key={label} className="flex flex-col gap-1">
              <span className="text-[9px] text-slate-400">{label}</span>
              <div className={`rounded px-2 py-1.5 text-[10px] border ${i === 0 ? 'bg-[#2d6a9f]/40 border-cyan-500/60 text-cyan-200' : 'bg-[#0c2a3e] border-slate-600 text-slate-300'}`}>
                {i === 0 ? 'This Year' : i === 1 ? '01/01/2024' : '12/31/2024'}
              </div>
            </div>
          ))}
        </div>
        <div className="mt-3 flex justify-end">
          <div className="bg-[#2d6a9f] text-white text-[10px] px-4 py-2 rounded-lg font-semibold shadow-lg shadow-[#2d6a9f]/30 cursor-default">
            Run Report ▶
          </div>
        </div>
        <div className="mt-2 text-[9px] text-slate-400">Preview: 1,247 transactions loaded...</div>
      </div>
    ),
  },
  {
    id: 4,
    icon: '⬇️',
    title: 'Export to Excel',
    subtitle: 'Download as .xlsx file',
    description: 'Click the Export icon (↓) in the top-right of the report → select "Export to Excel (.xlsx)".',
    visual: (
      <div className="relative w-full h-48 rounded-xl overflow-hidden bg-[#1a3a52] border border-[#2d6a9f]/40 p-4">
        <div className="flex justify-between items-center mb-3">
          <div className="text-[11px] text-slate-300 font-semibold">General Ledger Report</div>
          <div className="flex gap-1.5">
            {['✉️','🖨️','⚙️'].map(icon => <div key={icon} className="bg-[#0c2a3e] text-xs w-7 h-7 rounded flex items-center justify-center border border-slate-600">{icon}</div>)}
            <div className="relative">
              <div className="bg-cyan-600 text-white text-xs w-7 h-7 rounded flex items-center justify-center border border-cyan-400 animate-pulse">↓</div>
              <div className="absolute top-8 right-0 bg-white text-slate-800 rounded-lg shadow-xl border border-slate-200 w-40 z-10">
                <div className="px-3 py-2 text-[10px] font-semibold text-slate-500 border-b">Export</div>
                <div className="px-3 py-2 text-[10px] bg-cyan-50 text-cyan-700 font-semibold flex items-center gap-1.5">
                  <span>📗</span> Export to Excel (.xlsx)
                </div>
                <div className="px-3 py-2 text-[10px] text-slate-500 flex items-center gap-1.5">
                  <span>📄</span> Export to PDF
                </div>
              </div>
            </div>
          </div>
        </div>
        <div className="text-[9px] text-slate-400 space-y-1">
          {['Checking Account','Savings Account','Accounts Receivable'].map(a => (
            <div key={a} className="flex justify-between border-b border-slate-700/50 pb-1">
              <span>{a}</span><span className="text-slate-300">$12,450.00</span>
            </div>
          ))}
        </div>
      </div>
    ),
  },
  {
    id: 5,
    icon: '🏷️',
    title: 'Add Account Type Column',
    subtitle: 'Required — not in QBO export',
    description: 'Open the downloaded Excel file. Add a new column "Distribution account type" and fill in Asset, Liability, Equity, Income, Expense, or Cost of Goods Sold for each account.',
    isWarning: true,
    visual: (
      <div className="relative w-full h-48 rounded-xl overflow-hidden bg-[#1a1a1a] border border-amber-500/40 font-mono text-[9px]">
        {/* Excel mockup */}
        <div className="bg-[#217346] px-3 py-1 text-white text-[9px] font-semibold">📗 General_Ledger.xlsx</div>
        <div className="overflow-hidden">
          {/* Header row */}
          <div className="flex bg-[#e8e8e8] text-[#333] font-bold border-b border-slate-300">
            {['Distribution account', '+ Distribution account type', 'Transaction date', 'Amount'].map((h, i) => (
              <div key={h} className={`px-1.5 py-1 border-r border-slate-300 ${i === 1 ? 'bg-amber-100 text-amber-800 min-w-[120px]' : 'min-w-[90px]'}`} style={{minWidth: i===1?120:80}}>
                {i === 1 ? <span className="flex items-center gap-0.5"><span className="text-amber-500 font-black">+</span> {h.replace('+ ','')}</span> : h}
              </div>
            ))}
          </div>
          {[
            ['Checking', 'Asset', '2024-01-05', '$1,200.00'],
            ['Sales Revenue', 'Income', '2024-01-05', '$1,200.00'],
            ['Rent Expense', 'Expense', '2024-01-10', '$2,500.00'],
            ['SBA Loan', 'Liability', '2024-01-10', '$2,500.00'],
          ].map((row, ri) => (
            <div key={ri} className={`flex border-b ${ri % 2 === 0 ? 'bg-white' : 'bg-[#f8f8f8]'} text-slate-700`}>
              {row.map((cell, ci) => (
                <div key={ci} className={`px-1.5 py-1 border-r border-slate-200 truncate ${ci === 1 ? 'bg-amber-50 text-amber-700 font-semibold' : ''}`} style={{minWidth: ci===1?120:80}}>
                  {cell}
                </div>
              ))}
            </div>
          ))}
        </div>
        <div className="absolute bottom-2 left-2 right-2 bg-amber-500 text-white text-[9px] px-2 py-1 rounded-md font-semibold text-center">
          ⚠️ Add this column manually — QBO doesn't include it
        </div>
      </div>
    ),
  },
  {
    id: 6,
    icon: '✅',
    title: 'Upload Here',
    subtitle: 'Drag & drop your file above',
    description: 'Your file is ready! Upload the Excel file here. Valid account types: Asset, Liability, Equity, Income, Expense, Cost of Goods Sold.',
    visual: (
      <div className="relative w-full h-48 rounded-xl overflow-hidden bg-[#1a3a52] border border-emerald-500/40 flex flex-col items-center justify-center gap-3">
        <div className="text-4xl animate-bounce">📂</div>
        <div className="text-center">
          <div className="text-emerald-300 font-semibold text-sm">Ready to upload!</div>
          <div className="text-slate-400 text-xs mt-1">Drag & drop your .xlsx file</div>
        </div>
        <div className="flex flex-wrap gap-1 px-4 justify-center">
          {['Asset','Liability','Equity','Income','Expense','COGS'].map(t => (
            <span key={t} className="bg-emerald-900/60 text-emerald-300 text-[9px] px-2 py-0.5 rounded-full border border-emerald-500/30">{t}</span>
          ))}
        </div>
      </div>
    ),
  },
];

export function QBOGuide({ theme }: Props) {
  const [activeStep, setActiveStep] = useState(0);
  const [expanded, setExpanded] = useState(false);
  const d = theme === 'dark';

  const step = STEPS[activeStep];

  return (
    <div className={`rounded-2xl border overflow-hidden ${d ? 'border-slate-700 bg-slate-900' : 'border-slate-200 bg-white'}`}>
      {/* Header */}
      <button
        type="button"
        onClick={() => setExpanded(!expanded)}
        className={`w-full flex items-center justify-between px-5 py-4 transition-colors ${d ? 'hover:bg-slate-800' : 'hover:bg-slate-50'}`}
      >
        <div className="flex items-center gap-3">
          <span className="text-xl">📗</span>
          <div className="text-left">
            <p className={`text-sm font-bold ${d ? 'text-white' : 'text-slate-900'}`}>How to export from QuickBooks Online</p>
            <p className={`text-xs ${d ? 'text-slate-400' : 'text-slate-500'}`}>Interactive step-by-step guide · 6 steps</p>
          </div>
        </div>
        <span className={`text-lg transition-transform duration-300 ${expanded ? 'rotate-180' : ''} ${d ? 'text-slate-400' : 'text-slate-500'}`}>⌄</span>
      </button>

      {expanded && (
        <div className={`border-t ${d ? 'border-slate-700' : 'border-slate-200'}`}>
          {/* Step pills */}
          <div className="flex gap-1.5 px-5 pt-4 overflow-x-auto pb-1">
            {STEPS.map((s, i) => (
              <button
                key={s.id}
                type="button"
                onClick={() => setActiveStep(i)}
                className={`flex-shrink-0 flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-semibold transition-all ${
                  i === activeStep
                    ? s.isWarning ? 'bg-amber-500 text-white shadow-lg shadow-amber-500/30' : 'bg-cyan-500 text-white shadow-lg shadow-cyan-500/30'
                    : i < activeStep
                    ? d ? 'bg-emerald-900/50 text-emerald-300 border border-emerald-700' : 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                    : d ? 'bg-slate-800 text-slate-400 border border-slate-700' : 'bg-slate-100 text-slate-500 border border-slate-200'
                }`}
              >
                <span>{s.icon}</span>
                <span className="hidden sm:inline">{s.title}</span>
                <span className={`w-4 h-4 rounded-full flex items-center justify-center text-[9px] font-bold ${i === activeStep ? 'bg-white/20' : ''}`}>{s.id}</span>
              </button>
            ))}
          </div>

          {/* Active step */}
          <div className="p-5">
            <div className="grid gap-4 sm:grid-cols-2">
              {/* Visual */}
              <div>
                {step.visual}
              </div>
              {/* Description */}
              <div className="flex flex-col justify-between">
                <div>
                  <div className="flex items-center gap-2 mb-2">
                    <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${step.isWarning ? 'bg-amber-500/20 text-amber-400' : 'bg-cyan-500/20 text-cyan-400'}`}>Step {step.id} of {STEPS.length}</span>
                    {step.isWarning && <span className="text-xs bg-amber-500/20 text-amber-400 px-2 py-0.5 rounded-full">⚠️ Required</span>}
                  </div>
                  <h3 className={`text-lg font-bold ${d ? 'text-white' : 'text-slate-900'}`}>{step.title}</h3>
                  <p className={`text-xs font-medium mt-0.5 ${step.isWarning ? 'text-amber-400' : d ? 'text-cyan-400' : 'text-cyan-600'}`}>{step.subtitle}</p>
                  <p className={`mt-3 text-sm leading-relaxed ${d ? 'text-slate-300' : 'text-slate-600'}`}>{step.description}</p>

                  {step.isWarning && (
                    <div className={`mt-3 rounded-xl p-3 text-xs space-y-1 ${d ? 'bg-amber-950/40 border border-amber-800/50 text-amber-200' : 'bg-amber-50 border border-amber-200 text-amber-800'}`}>
                      <p className="font-bold">Valid account types:</p>
                      <div className="flex flex-wrap gap-1 mt-1">
                        {['Asset','Liability','Equity','Income','Expense','Cost of Goods Sold'].map(t => (
                          <span key={t} className={`px-2 py-0.5 rounded-full text-[10px] font-semibold ${d ? 'bg-amber-900/60 text-amber-300' : 'bg-amber-100 text-amber-700'}`}>{t}</span>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                {/* Navigation */}
                <div className="flex items-center gap-2 mt-4">
                  <button type="button" onClick={() => setActiveStep(Math.max(0, activeStep - 1))}
                    disabled={activeStep === 0}
                    className={`px-3 py-2 rounded-lg text-xs font-semibold transition-colors disabled:opacity-30 ${d ? 'bg-slate-800 text-slate-300 hover:bg-slate-700' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}>
                    ← Prev
                  </button>
                  <div className="flex gap-1 flex-1 justify-center">
                    {STEPS.map((_, i) => (
                      <button key={i} type="button" onClick={() => setActiveStep(i)}
                        className={`w-1.5 h-1.5 rounded-full transition-all ${i === activeStep ? 'bg-cyan-400 w-4' : d ? 'bg-slate-600' : 'bg-slate-300'}`} />
                    ))}
                  </div>
                  <button type="button" onClick={() => setActiveStep(Math.min(STEPS.length - 1, activeStep + 1))}
                    disabled={activeStep === STEPS.length - 1}
                    className="px-3 py-2 rounded-lg text-xs font-semibold bg-cyan-600 text-white hover:bg-cyan-500 transition-colors disabled:opacity-30">
                    Next →
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
