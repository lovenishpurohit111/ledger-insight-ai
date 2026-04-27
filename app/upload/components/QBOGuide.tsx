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
    description: 'In QuickBooks Online, click "Reports" in the left sidebar. Scroll down to the "For my accountant" section and click "General Ledger".',
    visual: (
      <div className="relative w-full h-48 rounded-xl overflow-hidden bg-[#1a3a52] border border-[#2d6a9f]/40 font-mono text-xs">
        <div className="absolute left-0 top-0 bottom-0 w-28 bg-[#0c2a3e] flex flex-col gap-1 p-2">
          {['Dashboard','Banking','Sales','Expenses','Payroll','Reports','Taxes','Accounting'].map((item) => (
            <div key={item} className={`rounded px-2 py-1.5 text-[10px] font-medium cursor-default ${item === 'Reports' ? 'bg-[#2d6a9f] text-white shadow-lg shadow-[#2d6a9f]/40' : 'text-slate-400'}`}>
              {item === 'Reports' && <span className="mr-1">›</span>}{item}
            </div>
          ))}
        </div>
        <div className="absolute left-32 top-3 right-3">
          <div className="text-[9px] text-slate-400 font-semibold mb-1 uppercase tracking-wider">For My Accountant</div>
          {['General Ledger','Trial Balance','Journal','Profit and Loss'].map((r, i) => (
            <div key={r} className={`px-2 py-1.5 text-[10px] rounded mb-0.5 ${i === 0 ? 'bg-cyan-600/40 text-cyan-200 border border-cyan-500/50' : 'text-slate-400'}`}>{r}</div>
          ))}
        </div>
        <div className="absolute left-1 top-[86px] rounded border-2 border-cyan-400 animate-pulse" style={{width:'108px', height:'28px'}} />
      </div>
    ),
  },
  {
    id: 2,
    icon: '⚙️',
    title: 'Click Customize',
    subtitle: 'Open the customization panel',
    description: 'On the General Ledger report page, click the "Customize" button in the top-right corner. This opens the report customization panel.',
    visual: (
      <div className="relative w-full h-48 rounded-xl overflow-hidden bg-[#1a3a52] border border-[#2d6a9f]/40 p-3">
        <div className="flex justify-between items-center mb-3">
          <div className="text-[11px] text-white font-semibold">General Ledger</div>
          <div className="flex gap-2">
            <div className="bg-[#0c2a3e] text-slate-300 text-[10px] px-2 py-1 rounded border border-slate-600">Save customization</div>
            <div className="relative">
              <div className="bg-cyan-600 text-white text-[10px] px-3 py-1 rounded border border-cyan-400 animate-pulse font-semibold">Customize</div>
              <div className="absolute -bottom-1 left-1/2 -translate-x-1/2 w-0 h-0 border-l-4 border-r-4 border-t-4 border-l-transparent border-r-transparent border-t-cyan-400 animate-bounce" />
            </div>
          </div>
        </div>
        <div className="text-[9px] text-slate-400 space-y-1 border-t border-slate-700 pt-2">
          {['Checking Account — Balance: $12,450','Sales Revenue — Balance: $34,200','Rent Expense — Balance: $8,500'].map(r => (
            <div key={r} className="border-b border-slate-700/50 pb-1">{r}</div>
          ))}
        </div>
        <div className="absolute bottom-3 right-3 bg-cyan-900/60 text-cyan-300 text-[9px] px-2 py-1 rounded-full border border-cyan-700">← Click here</div>
      </div>
    ),
  },
  {
    id: 3,
    icon: '📅',
    title: 'Set Date Range',
    subtitle: 'Choose your reporting period',
    description: 'In the Customize panel, set your Report Period (e.g. This Year, Last Year, or Custom date range). This defines which transactions appear.',
    visual: (
      <div className="relative w-full h-48 rounded-xl overflow-hidden bg-[#1a3a52] border border-[#2d6a9f]/40 p-3">
        <div className="text-[10px] text-white font-semibold mb-3">Customize report</div>
        <div className="bg-[#0c2a3e] rounded-lg p-3 border border-slate-600">
          <div className="text-[9px] text-slate-400 font-semibold uppercase mb-2">General</div>
          <div className="flex gap-2 flex-wrap">
            {['Report period','From','To'].map((label, i) => (
              <div key={label} className="flex flex-col gap-1">
                <span className="text-[9px] text-slate-400">{label}</span>
                <div className={`rounded px-2 py-1.5 text-[10px] border ${i === 0 ? 'bg-[#2d6a9f]/40 border-cyan-500/60 text-cyan-200' : 'bg-[#0c2a3e] border-slate-600 text-slate-300'}`}>
                  {i === 0 ? 'This Year ▾' : i === 1 ? '01/01/2024' : '12/31/2024'}
                </div>
              </div>
            ))}
          </div>
          <div className="mt-2 flex gap-2 flex-wrap">
            {['This Year','Last Year','Last Quarter','Custom'].map(t => (
              <span key={t} className={`text-[9px] px-2 py-0.5 rounded-full cursor-default ${t==='This Year' ? 'bg-cyan-600/40 text-cyan-300 border border-cyan-500/50' : 'text-slate-500'}`}>{t}</span>
            ))}
          </div>
        </div>
      </div>
    ),
  },
  {
    id: 4,
    icon: '➕',
    title: 'Add Columns',
    subtitle: 'Add "Distribution Account Type"',
    description: 'Still in Customize, click "Rows/Columns" → then "Change columns". Scroll through the list and check "Distribution account type". This is the key column our app needs.',
    isWarning: true,
    visual: (
      <div className="relative w-full h-48 rounded-xl overflow-hidden bg-[#1a3a52] border border-amber-500/40 p-3">
        <div className="text-[10px] text-white font-semibold mb-2">Customize → Rows/Columns</div>
        <div className="bg-[#0c2a3e] rounded-lg border border-slate-600 overflow-hidden">
          <div className="px-3 py-1.5 text-[9px] font-semibold text-slate-400 bg-slate-800 border-b border-slate-700">Change columns</div>
          {[
            { label: 'Transaction type', checked: true },
            { label: 'Num', checked: true },
            { label: 'Distribution account type', checked: true, highlight: true },
            { label: 'Name', checked: false },
            { label: 'Memo/Description', checked: false },
          ].map(col => (
            <div key={col.label} className={`flex items-center gap-2 px-3 py-1.5 text-[10px] border-b border-slate-700/50 ${col.highlight ? 'bg-amber-900/40' : ''}`}>
              <div className={`w-3.5 h-3.5 rounded border flex items-center justify-center flex-shrink-0 ${col.checked ? col.highlight ? 'bg-amber-500 border-amber-400' : 'bg-cyan-600 border-cyan-500' : 'border-slate-500'}`}>
                {col.checked && <span className="text-white text-[8px]">✓</span>}
              </div>
              <span className={col.highlight ? 'text-amber-300 font-semibold' : 'text-slate-300'}>{col.label}</span>
              {col.highlight && <span className="ml-auto text-[8px] bg-amber-500/30 text-amber-400 px-1.5 py-0.5 rounded-full">Required ✱</span>}
            </div>
          ))}
        </div>
      </div>
    ),
  },
  {
    id: 5,
    icon: '▶️',
    title: 'Run Report',
    subtitle: 'Apply and generate the report',
    description: 'Click "Run Report" to apply your customizations. Verify the report shows the "Distribution account type" column. Then click "Save customization" to reuse this setup.',
    visual: (
      <div className="relative w-full h-48 rounded-xl overflow-hidden bg-[#1a3a52] border border-[#2d6a9f]/40 p-3">
        <div className="flex justify-between items-center mb-3">
          <div className="text-[11px] text-white font-semibold">General Ledger</div>
          <div className="bg-[#2d6a9f] text-white text-[10px] px-3 py-1.5 rounded font-semibold shadow-lg shadow-[#2d6a9f]/40 cursor-default">▶ Run Report</div>
        </div>
        <div className="text-[9px] text-slate-400 font-semibold uppercase mb-1">Preview — columns visible:</div>
        <div className="overflow-auto">
          <div className="flex bg-slate-800 text-[9px] text-slate-300 font-semibold">
            {['Distribution account','Distribution account type','Date','Type','Amount'].map((h,i) => (
              <div key={h} className={`px-2 py-1 border-r border-slate-600 flex-shrink-0 ${i===1?'bg-amber-900/30 text-amber-300':''}`} style={{minWidth:i===1?110:80}}>{h}</div>
            ))}
          </div>
          {[['Checking','Asset','2024-01','Deposit','$1,200'],['Sales Rev','Income','2024-01','Invoice','$3,400']].map((row,ri) => (
            <div key={ri} className={`flex text-[9px] ${ri%2===0?'bg-[#0c2a3e]':'bg-[#0f3245]'}`}>
              {row.map((c,ci) => (
                <div key={ci} className={`px-2 py-1 border-r border-slate-700/50 flex-shrink-0 ${ci===1?'text-amber-300':''}`} style={{minWidth:ci===1?110:80}}>{c}</div>
              ))}
            </div>
          ))}
        </div>
      </div>
    ),
  },
  {
    id: 6,
    icon: '⬇️',
    title: 'Export to Excel',
    subtitle: 'Download as .xlsx and upload here',
    description: 'Click the Export icon (↓) in the top-right → select "Export to Excel (.xlsx)". Save the file, then drag & drop it into the upload area on this page.',
    visual: (
      <div className="relative w-full h-48 rounded-xl overflow-hidden bg-[#1a3a52] border border-emerald-500/40 p-3">
        <div className="flex justify-between items-center mb-2">
          <div className="text-[11px] text-white font-semibold">General Ledger</div>
          <div className="relative">
            <div className="bg-emerald-600 text-white text-[10px] w-7 h-7 rounded flex items-center justify-center border border-emerald-400 animate-pulse font-bold">↓</div>
            <div className="absolute top-8 right-0 bg-white text-slate-800 rounded-lg shadow-xl border border-slate-200 w-44 z-10">
              <div className="px-3 py-2 text-[9px] font-semibold text-slate-500 border-b">Export</div>
              <div className="px-3 py-2 text-[10px] bg-emerald-50 text-emerald-700 font-semibold flex items-center gap-1.5">
                <span>📗</span> Export to Excel (.xlsx)
              </div>
              <div className="px-3 py-2 text-[10px] text-slate-400 flex items-center gap-1.5">
                <span>📄</span> Export to PDF
              </div>
            </div>
          </div>
        </div>
        <div className="mt-12 flex flex-col items-center justify-center gap-2">
          <div className="text-3xl">📂</div>
          <div className="text-emerald-300 font-semibold text-xs">Drop the downloaded file here!</div>
          <div className="flex flex-wrap gap-1 justify-center">
            {['Asset','Liability','Equity','Income','Expense','COGS'].map(t => (
              <span key={t} className="bg-emerald-900/60 text-emerald-300 text-[9px] px-2 py-0.5 rounded-full border border-emerald-500/30">{t}</span>
            ))}
          </div>
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
      <button type="button" onClick={() => setExpanded(!expanded)}
        className={`w-full flex items-center justify-between px-5 py-4 transition-colors ${d ? 'hover:bg-slate-800' : 'hover:bg-slate-50'}`}>
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
          <div className="flex gap-1.5 px-5 pt-4 overflow-x-auto pb-1">
            {STEPS.map((s, i) => (
              <button key={s.id} type="button" onClick={() => setActiveStep(i)}
                className={`flex-shrink-0 flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-semibold transition-all ${
                  i === activeStep ? s.isWarning ? 'bg-amber-500 text-white shadow-lg shadow-amber-500/30' : 'bg-cyan-500 text-white shadow-lg shadow-cyan-500/30'
                  : i < activeStep ? d ? 'bg-emerald-900/50 text-emerald-300 border border-emerald-700' : 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                  : d ? 'bg-slate-800 text-slate-400 border border-slate-700' : 'bg-slate-100 text-slate-500 border border-slate-200'
                }`}>
                <span>{s.icon}</span>
                <span className="hidden sm:inline">{s.title}</span>
                <span className="w-4 h-4 rounded-full flex items-center justify-center text-[9px] font-bold">{s.id}</span>
              </button>
            ))}
          </div>

          <div className="p-5">
            <div className="grid gap-4 sm:grid-cols-2">
              <div>{step.visual}</div>
              <div className="flex flex-col justify-between">
                <div>
                  <div className="flex items-center gap-2 mb-2">
                    <span className={`text-xs font-bold px-2 py-0.5 rounded-full ${step.isWarning ? 'bg-amber-500/20 text-amber-400' : 'bg-cyan-500/20 text-cyan-400'}`}>Step {step.id} of {STEPS.length}</span>
                    {step.isWarning && <span className="text-xs bg-amber-500/20 text-amber-400 px-2 py-0.5 rounded-full">⚠️ Key step</span>}
                  </div>
                  <h3 className={`text-lg font-bold ${d ? 'text-white' : 'text-slate-900'}`}>{step.title}</h3>
                  <p className={`text-xs font-medium mt-0.5 ${step.isWarning ? 'text-amber-400' : d ? 'text-cyan-400' : 'text-cyan-600'}`}>{step.subtitle}</p>
                  <p className={`mt-3 text-sm leading-relaxed ${d ? 'text-slate-300' : 'text-slate-600'}`}>{step.description}</p>
                  {step.isWarning && (
                    <div className={`mt-3 rounded-xl p-3 text-xs ${d ? 'bg-amber-950/40 border border-amber-800/50 text-amber-200' : 'bg-amber-50 border border-amber-200 text-amber-800'}`}>
                      <p className="font-bold mb-1">⚠️ QBO does NOT export account type by default.</p>
                      <p>You must add it via <strong>Customize → Rows/Columns → Change columns → Distribution account type</strong> before exporting.</p>
                    </div>
                  )}
                </div>
                <div className="flex items-center gap-2 mt-4">
                  <button type="button" onClick={() => setActiveStep(Math.max(0, activeStep - 1))} disabled={activeStep === 0}
                    className={`px-3 py-2 rounded-lg text-xs font-semibold transition-colors disabled:opacity-30 ${d ? 'bg-slate-800 text-slate-300 hover:bg-slate-700' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}>
                    ← Prev
                  </button>
                  <div className="flex gap-1 flex-1 justify-center">
                    {STEPS.map((_, i) => (
                      <button key={i} type="button" onClick={() => setActiveStep(i)}
                        className={`h-1.5 rounded-full transition-all ${i === activeStep ? 'bg-cyan-400 w-4' : d ? 'bg-slate-600 w-1.5' : 'bg-slate-300 w-1.5'}`} />
                    ))}
                  </div>
                  <button type="button" onClick={() => setActiveStep(Math.min(STEPS.length - 1, activeStep + 1))} disabled={activeStep === STEPS.length - 1}
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
