import React from 'react';
import { ContainerCounts, FreightConfig, JobCalculationSummary } from '../types/procurement';
import { formatMoney, formatPct } from '../utils/calculations';

interface FreightPanelProps {
  freight: FreightConfig;
  calc: JobCalculationSummary;
  curr: string;
  disabled: boolean;
  primary: string;
  onChange: (patch: Partial<FreightConfig>) => void;
}

const pctValue = (rate: number) => Math.round((rate * 100 + Number.EPSILON) * 100) / 100;

function numberInput(value: number | string, onChange: (raw: string) => void, disabled: boolean, label: string, step = '0.01') {
  return (
    <input
      type="number"
      min="0"
      step={step}
      value={value}
      disabled={disabled}
      aria-label={label}
      onChange={(e) => onChange(e.target.value)}
      className="w-full min-h-[36px] rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-right text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50 disabled:text-slate-400"
    />
  );
}

export const FreightPanel: React.FC<FreightPanelProps> = ({ freight, calc, curr, disabled, primary, onChange }) => {
  const mode = freight.mode === 'sea' ? 'sea' : 'air';
  const containers: ContainerCounts = freight.containers ?? { c20: 0, c40: 0, c40hc: 0 };
  const rates: ContainerCounts = freight.containerRates ?? { c20: 0, c40: 0, c40hc: 0 };
  const markupDisplay = pctValue(calc.freightMarkupPct);
  const docDisplay = pctValue(freight.docChargePct);
  const loadingDisplay = pctValue(freight.loadingChargePct ?? 0.05);

  return (
    <div className="pt-4 border-t border-slate-100 space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <div className="text-xs font-semibold text-slate-700">Freight method</div>
          <div className="text-[11px] text-slate-500">Air uses CBM and kg rates above. Sea uses containers.</div>
        </div>
        <div className="inline-flex rounded-lg bg-slate-100 p-1">
          {(['air', 'sea'] as const).map((m) => (
            <button
              key={m}
              type="button"
              disabled={disabled}
              onClick={() => onChange({ mode: m })}
              className={`px-4 min-h-[36px] rounded-md text-xs font-semibold cursor-pointer disabled:opacity-50 ${
                mode === m ? 'text-white shadow-sm' : 'text-slate-600 hover:text-slate-900'
              }`}
              style={mode === m ? { backgroundColor: primary } : undefined}
            >
              {m === 'air' ? 'Air freight' : 'Sea freight (containers)'}
            </button>
          ))}
        </div>
      </div>

      {mode === 'sea' && (
        <div className="overflow-x-auto">
          <table className="w-full text-xs border-collapse">
            <thead>
              <tr className="text-slate-500 border-b border-slate-200">
                <th className="py-2 px-2 text-left font-semibold">Container</th>
                <th className="py-2 px-2 text-right font-semibold w-28">Containers</th>
                <th className="py-2 px-2 text-right font-semibold w-36">Rate per container ({curr})</th>
                <th className="py-2 px-2 text-right font-semibold w-36">Freight ({curr})</th>
              </tr>
            </thead>
            <tbody>
              {calc.seaLines.map((line) => (
                <tr key={line.key} className="border-b border-slate-100">
                  <td className="py-2 px-2 font-sans text-slate-800">{line.label}</td>
                  <td className="py-1.5 px-2">
                    {numberInput(
                      line.qty,
                      (raw) => onChange({ containers: { ...containers, [line.key]: Math.max(0, Math.floor(Number(raw) || 0)) } }),
                      disabled,
                      `${line.label} count`,
                      '1'
                    )}
                  </td>
                  <td className="py-1.5 px-2">
                    {numberInput(
                      rates[line.key] ?? 0,
                      (raw) => onChange({ containerRates: { ...rates, [line.key]: Math.max(0, Number(raw) || 0) } }),
                      disabled,
                      `${line.label} rate`
                    )}
                  </td>
                  <td className="py-2 px-2 text-right font-mono-tabular text-slate-900">{formatMoney(line.amount)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="space-y-1">
          <label className="block text-xs font-medium text-slate-600">Documentation charge %</label>
          {numberInput(docDisplay, (raw) => onChange({ docChargePct: Math.min(100, Math.max(0, Number(raw) || 0)) / 100 }), disabled, 'Documentation charge percent')}
        </div>
        <div className="space-y-1">
          <label className="block text-xs font-medium text-slate-600">
            Loading charge % {mode === 'air' ? '(sea only)' : ''}
          </label>
          {numberInput(
            loadingDisplay,
            (raw) => onChange({ loadingChargePct: Math.min(100, Math.max(0, Number(raw) || 0)) / 100 }),
            disabled || mode === 'air',
            'Loading charge percent'
          )}
        </div>
        <div className="space-y-1">
          <label className="block text-xs font-medium text-slate-600">Freight markup % (own input)</label>
          {numberInput(
            markupDisplay,
            (raw) => onChange({ freightMarkupPct: Math.min(500, Math.max(0, Number(raw) || 0)) / 100 }),
            disabled,
            'Freight markup percent'
          )}
          <p className="text-[11px] text-slate-500">Separate from the item markup above.</p>
        </div>
      </div>

      {/* Itemised breakdown for the internal sheet (customer documents show one freight line) */}
      <div className="rounded-lg border border-slate-200 bg-slate-50/60 p-3 text-xs font-mono-tabular">
        <div className="font-sans font-semibold text-slate-700 mb-2">Freight breakdown (internal only)</div>
        <div className="space-y-1">
          {mode === 'sea' ? (
            calc.seaLines.map((l) => (
              <div key={l.key} className="flex justify-between">
                <span className="font-sans text-slate-600">
                  {l.label} × {l.qty} @ {formatMoney(l.rate)}
                </span>
                <span>{formatMoney(l.amount, curr)}</span>
              </div>
            ))
          ) : (
            <div className="flex justify-between">
              <span className="font-sans text-slate-600">Base freight ({calc.freightBasis === 'KG' ? 'by kg' : 'by CBM'})</span>
              <span>{formatMoney(calc.baseFreight, curr)}</span>
            </div>
          )}
          <div className="flex justify-between">
            <span className="font-sans text-slate-600">Documentation ({formatPct(freight.docChargePct)})</span>
            <span>{formatMoney(calc.docChargeAmount, curr)}</span>
          </div>
          {mode === 'sea' && (
            <div className="flex justify-between">
              <span className="font-sans text-slate-600">Loading ({formatPct(freight.loadingChargePct ?? 0.05)})</span>
              <span>{formatMoney(calc.loadingChargeAmount, curr)}</span>
            </div>
          )}
          <div className="flex justify-between">
            <span className="font-sans text-slate-600">Freight markup ({formatPct(calc.freightMarkupPct)})</span>
            <span>{formatMoney(calc.freightMarkupAmount, curr)}</span>
          </div>
          <div className="flex justify-between border-t border-slate-300 pt-1 font-semibold text-slate-900">
            <span className="font-sans">Customer freight (one line on customer documents)</span>
            <span>{formatMoney(calc.customerFreight, curr)}</span>
          </div>
        </div>
      </div>
    </div>
  );
};
