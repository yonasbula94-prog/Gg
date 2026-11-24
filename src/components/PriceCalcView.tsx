import React, { useState } from 'react';
import { useRowSelection } from '../hooks/useRowSelection';
import { PasteIntoRows } from './PasteIntoRows';
import { parseFlexibleNumber } from '../utils/supplierQuoteParser';
import { ArrowRight, Building2, Lock } from 'lucide-react';
import { MARKUP_RATES } from '../constants/themes';
import {
  ConfirmDialogRequest,
  FreightConfig,
  JobCalculationSummary,
  JobControls,
  JobItemRow,
  MarkupRate,
  PlatformMode,
  ProcurementJob,
  ThemePalette,
  WorkspaceTab,
} from '../types/procurement';
import { formatMoney, formatNumber, formatPct } from '../utils/calculations';
import { HelpTip } from './HelpTip';
import { IPadJobEditor } from './IPadJobEditor';
import { FreightPanel } from './FreightPanel';

interface PriceCalcViewProps {
  job: ProcurementJob;
  amending: boolean;
  calc: JobCalculationSummary;
  theme: ThemePalette;
  platformMode: PlatformMode;
  onUpdateItem: (rowNumber: number, patch: Partial<JobItemRow>) => void;
  onUpdateFreight: (patch: Partial<FreightConfig>) => void;
  onUpdateControls: (patch: Partial<JobControls>) => void;
  onNavigateTab?: (tab: WorkspaceTab) => void;
  onRequestConfirm?: (req: ConfirmDialogRequest) => void;
}

function parseOptionalNumber(val: string): number | null {
  if (val.trim() === '') return null;
  const n = Number(val);
  if (!Number.isFinite(n) || n < 0) return null;
  return n;
}

export const PriceCalcView: React.FC<PriceCalcViewProps> = ({
  job,
  amending,
  calc,
  theme,
  platformMode,
  onUpdateItem,
  onUpdateFreight,
  onUpdateControls,
  onNavigateTab,
  onRequestConfirm,
}) => {
  const curr = job.controls.currency;
  const activeMarkup = job.freight.selectedMarkupPct;
  const isPresetMarkup = MARKUP_RATES.some((r) => Math.abs(r - activeMarkup) < 0.0001);
  const activeMarkupPctDisplay = Math.round((activeMarkup * 100 + Number.EPSILON) * 100) / 100;
  const includeVat = job.controls.includeVat !== false;
  const vatRateDisplay = Math.round((job.controls.vatRate * 100 + Number.EPSILON) * 100) / 100;
  const effectiveVat = calc.effectiveVatRate;
  // Once any document is issued the costing is frozen: edits would change the issued totals.
  const isIssued = !amending && Object.values(job.issuedDocs).some((d) => d && !d.isVoided);
  const guardItem = (rowNumber: number, patch: Partial<JobItemRow>) => {
    if (!isIssued) onUpdateItem(rowNumber, patch);
  };
  const guardFreight = (patch: Partial<FreightConfig>) => {
    if (!isIssued) onUpdateFreight(patch);
  };
  const guardControls = (patch: Partial<JobControls>) => {
    if (!isIssued) onUpdateControls(patch);
  };
  const rowSel = useRowSelection<number>();
  const [bulkCost, setBulkCost] = useState('');
  // Every row can be selected, including empty ones (so values can be applied before the description is entered)
  const selectableRows = calc.rows.map((r) => r.rowNumber);
  const applyBulk = (patch: Partial<JobItemRow>) => {
    rowSel.selected.forEach((rn) => guardItem(rn, patch));
  };
  const supName = job.supplierQuote?.supplierName || job.controls.supplierName;
  const supRef = job.supplierQuote?.supplierQuoteRef || job.controls.supplierRef;
  const supTerms = job.supplierQuote?.paymentTerms || '';

  return (
    <div className="space-y-6">
      {/* Confidential internal notice */}
      <div className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-red-200 bg-red-50/70 px-4 sm:px-5 py-3">
        <div className="flex items-center gap-2 text-sm text-red-900">
          <Lock className="w-4 h-4 text-red-700 shrink-0" />
          <span className="font-semibold">Internal Cost Sheet (PRICE CALC)</span>
          <HelpTip text="Holds confidential supplier costs, markup scenarios, and internal VAT provision. Automatically excluded from all customer-facing exports (QT, PI, CI, PL, DN)." />
        </div>
        <div className="text-xs text-red-800 font-mono-tabular">
          {effectiveVat > 0
            ? `Unit Price = Cost × (1 + Markup) × (1 + ${formatPct(effectiveVat)} VAT)`
            : 'Unit Price = Cost × (1 + Markup) · VAT Off'}
        </div>
      </div>

      {/* Linked Internal Supplier Quote Strip */}
      <div className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-slate-200 bg-white px-4 sm:px-5 py-3">
        <div className="flex flex-wrap items-center gap-2 text-xs text-slate-600 min-w-0">
          <Building2 className="w-4 h-4 text-slate-500 shrink-0" />
          <span className="font-semibold text-slate-900">
            Supplier Quote: {supName || 'Not Entered Yet'}
          </span>
          {supRef && (
            <>
              <span className="text-slate-300">·</span>
              <span className="font-mono-tabular">Ref {supRef}</span>
            </>
          )}
          {supTerms && (
            <>
              <span className="text-slate-300">·</span>
              <span className="truncate max-w-[260px]">{supTerms}</span>
            </>
          )}
          <span className="text-slate-300">·</span>
          <span className="font-mono-tabular font-semibold text-slate-800">
            Total Supplier Cost: {formatMoney(calc.totalSupplierCost, curr)}
          </span>
        </div>
        {onNavigateTab && (
          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={() => onNavigateTab('INTERNAL_QUOTE')}
              className="inline-flex items-center gap-1.5 rounded-md border border-slate-200 bg-slate-50 hover:bg-slate-100 px-3 py-1.5 text-xs font-medium text-slate-800 cursor-pointer transition-colors whitespace-nowrap"
            >
              <span>Edit Internal Quote</span>
            </button>
            <button
              type="button"
              onClick={() => onNavigateTab('QT')}
              className="inline-flex items-center gap-1 rounded-md px-3 py-1.5 text-xs font-semibold text-white hover:opacity-90 cursor-pointer transition-opacity whitespace-nowrap"
              style={{ backgroundColor: theme.primary }}
            >
              <span>Go to Quotation</span>
              <ArrowRight className="w-3.5 h-3.5" />
            </button>
          </div>
        )}
      </div>

      {isIssued && (
        <div className="flex items-center gap-2 rounded-lg border border-amber-200 bg-amber-50 px-4 py-2.5 text-xs text-amber-900">
          <Lock className="w-3.5 h-3.5 shrink-0" />
          <span>A document is issued for this job, so costing is locked to keep issued totals accurate. Create a new job to revise pricing.</span>
        </div>
      )}

      {/* Controls Strip: Markup Selector (E3) + Internal VAT Toggle (E4) + Freight Rates */}
      <div className="rounded-lg border border-slate-200 bg-white p-5 space-y-5">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-end">
          {/* Active Markup Selector (E3) — Presets + Any Custom % */}
          <div className="lg:col-span-6 space-y-2">
            <label className="flex items-center gap-1.5 text-xs font-medium text-slate-600">
              <span>Active Markup % (Cell E3)</span>
              <HelpTip text="Tap any quick preset (10%, 15%, 18%, 20%) or type any custom percentage (0%–500%). Feeds customer Unit Prices and Freight automatically." />
            </label>
            <div className="flex flex-wrap sm:flex-nowrap items-center gap-2">
              <div className="flex-1 grid grid-cols-4 gap-1.5 p-1 rounded-lg bg-slate-100 min-w-[210px]">
                {MARKUP_RATES.map((rate) => {
                  const isSelected = Math.abs(activeMarkup - rate) < 0.0001;
                  return (
                    <button
                      key={rate}
                      type="button"
                      onClick={() => guardFreight({ selectedMarkupPct: rate })}
                      className={`min-h-[40px] rounded-md text-sm font-semibold font-mono-tabular transition-colors cursor-pointer whitespace-nowrap ${
                        isSelected
                          ? 'text-white shadow-sm'
                          : 'text-slate-600 hover:text-slate-900 hover:bg-white/60'
                      }`}
                      style={isSelected ? { backgroundColor: theme.primary } : undefined}
                    >
                      {(rate * 100).toFixed(0)}%
                    </button>
                  );
                })}
              </div>

              <div className="flex items-center gap-1.5 shrink-0">
                <div
                  className={`flex items-center rounded-lg border px-2.5 min-h-[48px] transition-colors ${
                    !isPresetMarkup
                      ? 'border-slate-900 bg-white ring-1 ring-slate-900'
                      : 'border-slate-200 bg-slate-50/60 focus-within:border-slate-900 focus-within:bg-white'
                  }`}
                >
                  <span className="text-[11px] font-medium text-slate-500 mr-1.5 whitespace-nowrap">
                    Custom
                  </span>
                  <input
                    type="number"
                    inputMode="decimal"
                    step="0.5"
                    min="0"
                    max="500"
                    aria-label="Custom Markup Percentage"
                    value={activeMarkupPctDisplay}
                    onChange={(e) => {
                      const raw = e.target.value;
                      if (raw === '') {
                        guardFreight({ selectedMarkupPct: 0 });
                        return;
                      }
                      const num = Number(raw);
                      if (Number.isFinite(num) && num >= 0 && num <= 500) {
                        guardFreight({ selectedMarkupPct: num / 100 });
                      }
                    }}
                    className="w-16 bg-transparent text-right text-sm font-mono-tabular font-bold text-blue-700 focus:outline-none"
                  />
                  <span className="text-xs font-semibold text-slate-500 ml-1">%</span>
                </div>
              </div>
            </div>
          </div>

          {/* Internal VAT Provision (Cell E4 — Baked into Unit Price, Hidden from End User) */}
          <div className="lg:col-span-6 space-y-2">
            <label className="flex items-center gap-1.5 text-xs font-medium text-slate-600">
              <span>Internal VAT Provision (Cell E4)</span>
              <HelpTip text="Included inside every customer Unit Price (Col R) and never shown as a separate VAT line on PI, QT, or CI. Toggle Remove VAT (0%) for VAT-exempt orders." />
            </label>
            <div className="flex flex-wrap sm:flex-nowrap items-center gap-2">
              <div className="inline-flex rounded-lg bg-slate-100 p-1 shrink-0">
                <button
                  type="button"
                  onClick={() =>
                    guardControls({
                      includeVat: true,
                      vatRate: job.controls.vatRate > 0 ? job.controls.vatRate : 0.05,
                    })
                  }
                  className={`px-3.5 min-h-[40px] rounded-md text-xs font-semibold transition-colors cursor-pointer whitespace-nowrap ${
                    includeVat
                      ? 'text-white shadow-sm'
                      : 'text-slate-600 hover:text-slate-900 hover:bg-white/60'
                  }`}
                  style={includeVat ? { backgroundColor: theme.primary } : undefined}
                >
                  Include VAT
                </button>
                <button
                  type="button"
                  onClick={() => guardControls({ includeVat: false })}
                  className={`px-3.5 min-h-[40px] rounded-md text-xs font-semibold transition-colors cursor-pointer whitespace-nowrap ${
                    !includeVat
                      ? 'bg-slate-900 text-white shadow-sm'
                      : 'text-slate-600 hover:text-slate-900 hover:bg-white/60'
                  }`}
                >
                  Remove VAT (0%)
                </button>
              </div>

              <div
                className={`flex items-center rounded-lg border px-3 min-h-[48px] transition-colors ${
                  includeVat
                    ? 'border-slate-300 bg-white focus-within:border-slate-900'
                    : 'border-slate-200 bg-slate-100 opacity-60'
                }`}
              >
                <span className="text-[11px] font-medium text-slate-500 mr-1.5 whitespace-nowrap">
                  VAT Rate
                </span>
                <input
                  type="number"
                  inputMode="decimal"
                  step="0.5"
                  min="0"
                  max="100"
                  disabled={!includeVat}
                  aria-label="Internal VAT Percentage"
                  value={includeVat ? vatRateDisplay : 0}
                  onChange={(e) => {
                    const raw = e.target.value;
                    if (raw === '') {
                      guardControls({ vatRate: 0 });
                      return;
                    }
                    const num = Number(raw);
                    if (Number.isFinite(num) && num >= 0 && num <= 100) {
                      guardControls({
                        includeVat: true,
                        vatRate: Math.round(num * 100) / 10000,
                      });
                    }
                  }}
                  className="w-16 bg-transparent text-right text-sm font-mono-tabular font-bold text-blue-700 focus:outline-none disabled:text-slate-400 disabled:cursor-not-allowed"
                />
                <span className="text-xs font-semibold text-slate-500 ml-1">%</span>
              </div>
            </div>
          </div>
        </div>

        {/* Freight Rates & Doc Charge Row */}
        {job.freight.mode !== 'sea' && (
        <div className="pt-4 border-t border-slate-100 grid grid-cols-1 sm:grid-cols-3 gap-4">
          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1.5">
              Rate / CBM ({curr})
            </label>
            <input
              type="number"
              inputMode="decimal"
              step="0.01"
              min="0"
              value={job.freight.ratePerCbm}
              onChange={(e) =>
                guardFreight({ ratePerCbm: Math.max(0, Number(e.target.value) || 0) })
              }
              className="w-full min-h-[42px] rounded-lg border border-slate-200 bg-slate-50/50 px-3 py-2 text-sm font-mono-tabular text-blue-700 font-medium focus:border-slate-900 focus:bg-white focus:outline-none"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1.5">
              Rate / KG ({curr})
            </label>
            <input
              type="number"
              inputMode="decimal"
              step="0.01"
              min="0"
              value={job.freight.ratePerKg}
              onChange={(e) =>
                guardFreight({ ratePerKg: Math.max(0, Number(e.target.value) || 0) })
              }
              className="w-full min-h-[42px] rounded-lg border border-slate-200 bg-slate-50/50 px-3 py-2 text-sm font-mono-tabular text-blue-700 font-medium focus:border-slate-900 focus:bg-white focus:outline-none"
            />
          </div>
        </div>
        )}
        <FreightPanel
          freight={job.freight}
          calc={calc}
          curr={curr}
          disabled={isIssued}
          primary={theme.primary}
          onChange={guardFreight}
        />
      </div>

      {/* Active Custom Markup Card when using a non-preset % */}
      {!isPresetMarkup && (
        <div
          className="rounded-lg border-2 bg-white p-4 flex flex-wrap items-center justify-between gap-4"
          style={{ borderColor: theme.primary }}
        >
          <div>
            <div className="text-xs font-semibold uppercase tracking-wide" style={{ color: theme.primary }}>
              Active Custom Markup (Cell E3)
            </div>
            <div className="text-lg font-bold font-mono-tabular text-slate-900 mt-0.5">
              {formatPct(activeMarkup)} Markup
              {effectiveVat > 0 ? ` + ${formatPct(effectiveVat)} Internal VAT` : ' (VAT Removed)'} → Grand Total {formatMoney(calc.grandTotal, curr)}
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-5 text-xs font-mono-tabular text-slate-700">
            <div>
              <span className="text-slate-400 block font-sans">Goods (Excl. VAT)</span>
              <span className="font-semibold text-slate-900">{formatMoney(calc.goodsSubtotalExVat, curr)}</span>
            </div>
            <div>
              <span className="text-slate-400 block font-sans">
                {effectiveVat > 0 ? `Embedded VAT (${formatPct(effectiveVat)})` : 'Internal VAT'}
              </span>
              <span className="font-semibold text-slate-900">{formatMoney(calc.vatAmount, curr)}</span>
            </div>
            <div>
              <span className="text-slate-400 block font-sans">Customer Goods</span>
              <span className="font-semibold text-slate-900">{formatMoney(calc.goodsSubtotal, curr)}</span>
            </div>
            <div>
              <span className="text-slate-400 block font-sans">Customer Freight</span>
              <span className="font-semibold text-slate-900">{formatMoney(calc.customerFreight, curr)}</span>
            </div>
            <div>
              <span className="text-slate-400 block font-sans">Net Profit (Excl. VAT)</span>
              <span className="font-bold text-emerald-700">{formatMoney(calc.grossProfit, curr)}</span>
            </div>
          </div>
        </div>
      )}

      {/* Side-by-Side 10% / 15% / 18% / 20% Scenario Comparison Strip */}
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-4">
        {MARKUP_RATES.map((rate: MarkupRate) => {
          const scen = calc.scenarioTotals[rate];
          const isActive = Math.abs(activeMarkup - rate) < 0.0001;
          return (
            <button
              key={rate}
              type="button"
              onClick={() => guardFreight({ selectedMarkupPct: rate })}
              className={`text-left rounded-lg border p-4 transition-colors cursor-pointer ${
                isActive
                  ? 'bg-white border-slate-900 ring-1 ring-slate-900'
                  : 'bg-white border-slate-200 hover:border-slate-300'
              }`}
            >
              <div className="flex items-center justify-between">
                <span className="text-sm font-semibold text-slate-900">
                  {(rate * 100).toFixed(0)}% Markup Scenario
                </span>
                <span className="text-xs font-medium text-slate-500">
                  {isActive ? 'Active (E3)' : 'Select'}
                </span>
              </div>

              <div className="mt-3 text-xl font-bold font-mono-tabular text-slate-900">
                {formatMoney(scen.grandTotal, curr)}
              </div>

              <div className="mt-3 pt-3 border-t border-slate-100 space-y-1 text-xs text-slate-600 font-mono-tabular">
                <div className="flex justify-between">
                  <span>Customer Goods {effectiveVat > 0 ? '(Incl. VAT)' : ''}</span>
                  <span className="text-slate-900">{formatMoney(scen.goodsSubtotal)}</span>
                </div>
                <div className="flex justify-between">
                  <span>
                    {effectiveVat > 0
                      ? `Embedded VAT (${formatPct(effectiveVat)})`
                      : 'Internal VAT (Removed)'}
                  </span>
                  <span className="text-slate-500">{formatMoney(scen.vatAmount)}</span>
                </div>
                <div className="flex justify-between">
                  <span>Customer Freight</span>
                  <span className="text-slate-900">{formatMoney(scen.customerFreight)}</span>
                </div>
                <div className="flex justify-between pt-1 border-t border-slate-100 font-semibold text-emerald-700">
                  <span>Net Profit (Excl. VAT)</span>
                  <span>{formatMoney(scen.grossProfit, curr)}</span>
                </div>
              </div>
            </button>
          );
        })}
      </div>

      {/* Main PRICE CALC Editor: iPad Touch Cards vs Windows Desktop Excel Matrix */}
      {platformMode === 'ipad' ? (
        <div className="space-y-3">
          <div className="flex items-center justify-between px-1">
            <h2 className="text-sm font-semibold text-slate-900">
              iPad Touch Costing Cards (Rows 17–36)
            </h2>
            <span className="text-xs font-mono-tabular text-slate-500">
              Supplier Total: {formatMoney(calc.totalSupplierCost, curr)}
            </span>
          </div>
          <IPadJobEditor
            mode="costing"
            disabled={isIssued}
            items={job.items}
            computedRows={calc.rows}
            currency={curr}
            theme={theme}
            activeMarkupPct={activeMarkup}
            effectiveVatRate={effectiveVat}
            onUpdateItem={guardItem}
            onRequestConfirm={onRequestConfirm}
          />
        </div>
      ) : (
        <div className="rounded-lg border border-slate-200 bg-white overflow-hidden">
          <div className="flex flex-wrap items-center justify-between gap-2 px-5 py-3.5 border-b border-slate-200 bg-slate-50/70">
            <div className="flex items-center gap-2">
              <h2 className="text-sm font-semibold text-slate-900">
                Costing &amp; Logistics Matrix (Rows 17–36)
              </h2>
              <HelpTip
                text={`Blue numbers are editable inputs; black numbers are live formulas. ${
                  effectiveVat > 0
                    ? `Customer Unit Prices (Col R) include ${formatPct(effectiveVat)} internal VAT.`
                    : 'Internal VAT is currently OFF (0%).'
                }`}
              />
            </div>
            <div className="text-xs text-slate-600 font-mono-tabular">
              Active Items: {calc.activeRowCount} / 20 · Supplier Total: {formatMoney(calc.totalSupplierCost, curr)}
            </div>
          </div>

          {rowSel.count > 0 && (
            <div className="no-print flex flex-wrap items-center gap-2 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-xs">
              <span className="font-semibold text-slate-900">{rowSel.count} row{rowSel.count === 1 ? '' : 's'} selected</span>
              <span className="text-slate-300">|</span>
              <input
                type="number"
                min="0"
                step="0.01"
                value={bulkCost}
                onChange={(e) => setBulkCost(e.target.value)}
                placeholder="Supplier cost"
                disabled={isIssued}
                className="w-28 min-h-[32px] rounded border border-slate-200 bg-white px-2 text-right font-mono-tabular disabled:opacity-50"
              />
              <button
                type="button"
                disabled={isIssued || bulkCost.trim() === ''}
                onClick={() => { applyBulk({ supplierUnitCost: parseOptionalNumber(bulkCost) }); setBulkCost(''); }}
                className="rounded border border-slate-200 bg-white px-2.5 py-1.5 font-medium text-slate-800 hover:bg-slate-100 disabled:opacity-50 cursor-pointer"
              >
                Set cost
              </button>
              <button
                type="button"
                disabled={isIssued}
                onClick={() => applyBulk({ supplierUnitCost: null })}
                className="rounded border border-slate-200 bg-white px-2.5 py-1.5 font-medium text-slate-800 hover:bg-slate-100 disabled:opacity-50 cursor-pointer"
              >
                Clear costs
              </button>
              <button
                type="button"
                disabled={isIssued}
                onClick={() => applyBulk({ boxLengthCm: null, boxWidthCm: null, boxHeightCm: null })}
                className="rounded border border-slate-200 bg-white px-2.5 py-1.5 font-medium text-slate-800 hover:bg-slate-100 disabled:opacity-50 cursor-pointer"
              >
                Clear box size
              </button>
              <PasteIntoRows
                disabled={isIssued}
                rowIds={rowSel.selected}
                fields={[
                  { key: 'supplierUnitCost', label: 'Supplier cost', parse: (raw: string) => parseFlexibleNumber(raw) ?? undefined },
                  { key: 'boxes', label: 'Boxes', parse: (raw: string) => parseFlexibleNumber(raw) ?? undefined },
                  { key: 'boxLengthCm', label: 'Box L (cm)', parse: (raw: string) => parseFlexibleNumber(raw) ?? undefined },
                  { key: 'boxWidthCm', label: 'Box W (cm)', parse: (raw: string) => parseFlexibleNumber(raw) ?? undefined },
                  { key: 'boxHeightCm', label: 'Box H (cm)', parse: (raw: string) => parseFlexibleNumber(raw) ?? undefined },
                  { key: 'kgPerBox', label: 'Kg per box', parse: (raw: string) => parseFlexibleNumber(raw) ?? undefined },
                ]}
                onApply={(key, pairs) => pairs.forEach((p) => guardItem(p.rowId, { [key]: p.value } as Partial<JobItemRow>))}
              />
              <button type="button" onClick={rowSel.clear} className="ml-auto text-slate-600 underline cursor-pointer">
                Deselect all
              </button>
            </div>
          )}

          <div className="overflow-x-auto sheet-scroll">
            <table className="w-full border-collapse text-xs">
              <thead>
                <tr style={{ backgroundColor: theme.primary, color: theme.primaryText }}>
                  <th className="py-2.5 px-2 text-center w-8">
                    <input
                      type="checkbox"
                      aria-label="Select all rows"
                      disabled={selectableRows.length === 0}
                      checked={rowSel.allChecked(selectableRows)}
                      onChange={() => rowSel.toggleAll(selectableRows)}
                    />
                  </th>
                  <th className="py-2.5 px-2.5 text-center font-semibold border-r border-white/15 w-11">Row</th>
                  <th className="py-2.5 px-3 text-left font-semibold min-w-[150px]">Offered Description</th>
                  <th className="py-2.5 px-2 text-right font-semibold w-14">Qty</th>
                  <th className="py-2.5 px-2 text-right font-semibold min-w-[60px]">Boxes</th>
                  <th className="py-2.5 px-1.5 text-right font-semibold min-w-[60px]">L (cm)</th>
                  <th className="py-2.5 px-1.5 text-right font-semibold min-w-[60px]">W (cm)</th>
                  <th className="py-2.5 px-1.5 text-right font-semibold min-w-[60px]">H (cm)</th>
                  <th className="py-2.5 px-2 text-right font-semibold w-16">CBM</th>
                  <th className="py-2.5 px-2 text-right font-semibold w-16">Kg/Box</th>
                  <th className="py-2.5 px-2 text-right font-semibold w-16">Gross KG</th>
                  <th className="py-2.5 px-2 text-right font-semibold min-w-[84px]">Supplier Cost</th>
                  <th className="py-2.5 px-2 text-right font-semibold min-w-[84px]">Total Cost</th>
                  <th className="hidden xl:table-cell py-2.5 px-2 text-right font-semibold w-20">10%</th>
                  <th className="hidden xl:table-cell py-2.5 px-2 text-right font-semibold w-20">15%</th>
                  <th className="hidden xl:table-cell py-2.5 px-2 text-right font-semibold w-20">18%</th>
                  <th className="hidden xl:table-cell py-2.5 px-2 text-right font-semibold w-20">20%</th>
                  <th
                    className="py-2.5 px-2.5 text-right font-semibold min-w-[120px]"
                    style={{ backgroundColor: theme.accent }}
                  >
                    Col R ({formatPct(activeMarkup)}{effectiveVat > 0 ? ` + ${formatPct(effectiveVat)}` : ''})
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 font-mono-tabular">
                {calc.rows.map((row, idx) => {
                  const raw = job.items[idx];
                  const isInactive = row.isEmpty;
                  return (
                    <tr
                      key={row.rowNumber}
                      onMouseEnter={() => rowSel.dragOver(row.rowNumber)}
                      className={`${isInactive ? 'opacity-50' : ''} hover:bg-slate-50/90 transition-colors`}
                      style={!isInactive && idx % 2 === 1 ? { backgroundColor: theme.zebraRow } : undefined}
                    >
                      <td
                        className="py-1.5 px-2 text-center select-none"
                        onMouseDown={(e) => {
                          if (e.button === 0) rowSel.beginDrag(row.rowNumber, selectableRows);
                        }}
                      >
                        <input
                          type="checkbox"
                          aria-label={`Select row ${row.rowNumber}`}
                          checked={rowSel.has(row.rowNumber)}
                          onChange={(e) => rowSel.click(row.rowNumber, selectableRows, (e.nativeEvent as MouseEvent).shiftKey)}
                        />
                      </td>
                      <td className="py-1.5 px-2.5 text-center text-slate-400 border-r border-slate-100">
                        {row.rowNumber}
                      </td>
                      <td className="py-1.5 px-3 font-sans text-slate-800 max-w-[240px] truncate" title={row.offeredDescription || row.requestedDescription}>
                        {row.offeredDescription || row.requestedDescription || <span className="text-slate-300">—</span>}
                      </td>
                      <td className="py-1.5 px-2 text-right text-slate-700">
                        {row.qty !== null ? row.qty : ''}
                      </td>
                      <td className="p-1">
                        <input
                          type="number"
                          step="1"
                          min="0"
                          disabled={isInactive}
                          value={raw.boxes ?? ''}
                          onChange={(e) =>
                            guardItem(row.rowNumber, { boxes: parseOptionalNumber(e.target.value) })
                          }
                          className="w-full min-h-[34px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white min-w-[52px] px-1.5 py-1 text-right text-blue-700 focus:outline-none"
                        />
                      </td>
                      <td className="p-1">
                        <input
                          type="number"
                          step="1"
                          min="0"
                          disabled={isInactive}
                          value={raw.boxLengthCm ?? ''}
                          onChange={(e) =>
                            guardItem(row.rowNumber, { boxLengthCm: parseOptionalNumber(e.target.value) })
                          }
                          className="w-full min-h-[34px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white min-w-[52px] px-1.5 py-1 text-right text-blue-700 focus:outline-none"
                        />
                      </td>
                      <td className="p-1">
                        <input
                          type="number"
                          step="1"
                          min="0"
                          disabled={isInactive}
                          value={raw.boxWidthCm ?? ''}
                          onChange={(e) =>
                            guardItem(row.rowNumber, { boxWidthCm: parseOptionalNumber(e.target.value) })
                          }
                          className="w-full min-h-[34px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white min-w-[52px] px-1.5 py-1 text-right text-blue-700 focus:outline-none"
                        />
                      </td>
                      <td className="p-1">
                        <input
                          type="number"
                          step="1"
                          min="0"
                          disabled={isInactive}
                          value={raw.boxHeightCm ?? ''}
                          onChange={(e) =>
                            guardItem(row.rowNumber, { boxHeightCm: parseOptionalNumber(e.target.value) })
                          }
                          className="w-full min-h-[34px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white min-w-[52px] px-1.5 py-1 text-right text-blue-700 focus:outline-none"
                        />
                      </td>
                      <td className="py-1.5 px-2 text-right text-slate-700">
                        {row.cbm !== null ? formatNumber(row.cbm, 4) : ''}
                      </td>
                      <td className="p-1">
                        <input
                          type="number"
                          step="0.1"
                          min="0"
                          disabled={isInactive}
                          value={raw.kgPerBox ?? ''}
                          onChange={(e) =>
                            guardItem(row.rowNumber, { kgPerBox: parseOptionalNumber(e.target.value) })
                          }
                          className="w-full min-h-[34px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white min-w-[52px] px-1.5 py-1 text-right text-blue-700 focus:outline-none"
                        />
                      </td>
                      <td className="py-1.5 px-2 text-right text-slate-700">
                        {row.grossKgCalc !== null ? formatNumber(row.grossKgCalc, 2) : ''}
                      </td>
                      <td className="p-1">
                        <input
                          type="number"
                          step="0.01"
                          min="0"
                          disabled={isInactive}
                          value={raw.supplierUnitCost ?? ''}
                          placeholder={isInactive ? '' : '0.00'}
                          onChange={(e) =>
                            guardItem(row.rowNumber, {
                              supplierUnitCost: parseOptionalNumber(e.target.value),
                            })
                          }
                          className="w-full min-h-[34px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-right text-blue-700 font-medium focus:outline-none disabled:cursor-not-allowed"
                        />
                      </td>
                      <td className="py-1.5 px-2 text-right text-slate-900">
                        {formatMoney(row.supplierTotalCost)}
                      </td>
                      <td className={`hidden xl:table-cell py-1.5 px-2 text-right ${Math.abs(activeMarkup - 0.1) < 0.0001 ? "font-semibold text-slate-900" : "text-slate-500"}`}>
                        {formatMoney(row.unitPrice10)}
                      </td>
                      <td className={`hidden xl:table-cell py-1.5 px-2 text-right ${Math.abs(activeMarkup - 0.15) < 0.0001 ? "font-semibold text-slate-900" : "text-slate-500"}`}>
                        {formatMoney(row.unitPrice15)}
                      </td>
                      <td className={`hidden xl:table-cell py-1.5 px-2 text-right ${Math.abs(activeMarkup - 0.18) < 0.0001 ? "font-semibold text-slate-900" : "text-slate-500"}`}>
                        {formatMoney(row.unitPrice18)}
                      </td>
                      <td className={`hidden xl:table-cell py-1.5 px-2 text-right ${Math.abs(activeMarkup - 0.2) < 0.0001 ? "font-semibold text-slate-900" : "text-slate-500"}`}>
                        {formatMoney(row.unitPrice20)}
                      </td>
                      <td
                        className="py-1.5 px-2.5 text-right font-semibold text-slate-900"
                        style={{ backgroundColor: theme.secondaryBg }}
                      >
                        {formatMoney(row.activeUnitPrice)}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr
                  className="font-semibold font-mono-tabular border-t-2"
                  style={{ backgroundColor: theme.totalBg, borderColor: theme.strongBorder }}
                >
                  <td colSpan={3} className="py-3 px-3 text-left font-sans text-slate-900">
                    TOTALS (Rows 17–36)
                  </td>
                  <td className="py-3 px-2 text-right text-slate-900">{calc.totalQty}</td>
                  <td className="py-3 px-2 text-right text-slate-900">{calc.totalBoxes}</td>
                  <td colSpan={3} />
                  <td className="py-3 px-2 text-right text-slate-900">{formatNumber(calc.totalCbm, 4)}</td>
                  <td />
                  <td className="py-3 px-2 text-right text-slate-900">{formatNumber(calc.totalGrossKgCalc, 2)}</td>
                  <td className="py-3 px-2" />
                  <td className="py-3 px-2 text-right text-slate-900">{formatMoney(calc.totalSupplierCost)}</td>
                  <td colSpan={4} className="py-3 px-2 text-right font-sans text-slate-600">
                    Customer Goods Subtotal:
                  </td>
                  <td className="py-3 px-2.5 text-right text-slate-900">
                    {formatMoney(calc.goodsSubtotal)}
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}

      {/* Freight Basis Comparison & Internal Profitability Breakdown */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="rounded-lg border border-slate-200 bg-white p-5">
          <div className="flex items-center gap-1.5">
            <h3 className="text-sm font-semibold text-slate-900">
              {calc.freightMode === 'sea' ? 'Freight Calculation Breakdown (Sea · Containers)' : 'Freight Calculation Breakdown (MAX Basis)'}
            </h3>
            <HelpTip
              text={
                calc.freightMode === 'sea'
                  ? `Sea freight: containers × rate per container, plus the documentation charge (${formatPct(job.freight.docChargePct)}) and the loading charge (${formatPct(job.freight.loadingChargePct ?? 0.05)}). Then the freight markup (${formatPct(calc.freightMarkupPct)}).`
                  : `Air freight: the higher of volume (CBM) or weight (gross KG), plus the documentation charge (${formatPct(job.freight.docChargePct)}). Then the freight markup (${formatPct(calc.freightMarkupPct)}).`
              }
            />
          </div>

          <div className="mt-4 divide-y divide-slate-100 text-sm font-mono-tabular">
            {calc.freightMode === 'sea' ? (
              <>
                {calc.seaLines.map((l) => (
                  <div key={l.key} className="flex justify-between py-2">
                    <span className="font-sans text-slate-600">
                      {l.label} × {l.qty} @ {formatMoney(l.rate)}
                    </span>
                    <span className={l.amount > 0 ? 'text-slate-900' : 'text-slate-400'}>{formatMoney(l.amount, curr)}</span>
                  </div>
                ))}
                <div className="flex justify-between py-2">
                  <span className="font-sans text-slate-700 font-medium">Container Freight (total)</span>
                  <span className="font-medium text-slate-900">{formatMoney(calc.baseFreight, curr)}</span>
                </div>
              </>
            ) : (
              <>
                <div className="flex justify-between py-2">
                  <span className="font-sans text-slate-600">
                    Volume Freight ({formatNumber(calc.totalCbm, 4)} CBM × {formatMoney(job.freight.ratePerCbm)})
                    {calc.freightBasis === 'CBM' && <span className="ml-2 font-semibold text-slate-900">· Governing</span>}
                  </span>
                  <span className={calc.freightBasis === 'CBM' ? 'font-semibold text-slate-900' : 'text-slate-500'}>
                    {formatMoney(calc.freightByCbm, curr)}
                  </span>
                </div>
                <div className="flex justify-between py-2">
                  <span className="font-sans text-slate-600">
                    Weight Freight ({formatNumber(calc.totalGrossKgCalc, 2)} KG × {formatMoney(job.freight.ratePerKg)})
                    {calc.freightBasis === 'KG' && <span className="ml-2 font-semibold text-slate-900">· Governing</span>}
                  </span>
                  <span className={calc.freightBasis === 'KG' ? 'font-semibold text-slate-900' : 'text-slate-500'}>
                    {formatMoney(calc.freightByKg, curr)}
                  </span>
                </div>
                <div className="flex justify-between py-2">
                  <span className="font-sans text-slate-700 font-medium">Base Freight (MAX)</span>
                  <span className="font-medium text-slate-900">{formatMoney(calc.baseFreight, curr)}</span>
                </div>
              </>
            )}
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-600">
                + Documentation Charge ({formatPct(job.freight.docChargePct)})
              </span>
              <span className="text-slate-900">{formatMoney(calc.docChargeAmount, curr)}</span>
            </div>
            {calc.freightMode === 'sea' && (
              <div className="flex justify-between py-2">
                <span className="font-sans text-slate-600">
                  + Loading Charge ({formatPct(job.freight.loadingChargePct ?? 0.05)})
                </span>
                <span className="text-slate-900">{formatMoney(calc.loadingChargeAmount, curr)}</span>
              </div>
            )}
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-600">
                + Markup on Freight ({formatPct(calc.freightMarkupPct)})
              </span>
              <span className="text-slate-900">{formatMoney(calc.freightMarkupAmount, curr)}</span>
            </div>
            <div className="flex justify-between py-2.5 font-semibold text-slate-900">
              <span className="font-sans">Customer Freight Line (PI / QT / CI)</span>
              <span>{formatMoney(calc.customerFreight, curr)}</span>
            </div>
          </div>
        </div>

        <div className="rounded-lg border border-slate-200 bg-white p-5">
          <div className="flex items-center gap-1.5">
            <h3 className="text-sm font-semibold text-slate-900">
              Job Profitability &amp; Internal VAT Summary
            </h3>
            <HelpTip
              text={
                effectiveVat > 0
                  ? `Internal VAT (${formatPct(effectiveVat)}) is baked into every customer Unit Price — end-user documents show Goods Subtotal + Freight only.`
                  : 'Internal VAT is removed (0%) — customer Unit Prices reflect Supplier Cost + Markup only.'
              }
            />
          </div>

          <div className="mt-4 divide-y divide-slate-100 text-sm font-mono-tabular">
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-600">Goods Revenue (Excl. VAT)</span>
              <span className="text-slate-900">{formatMoney(calc.goodsSubtotalExVat, curr)}</span>
            </div>
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-600">
                {effectiveVat > 0
                  ? `+ Embedded Internal VAT (${formatPct(effectiveVat)} in Unit Price)`
                  : 'Internal VAT (Removed)'}
              </span>
              <span className="text-slate-700">{formatMoney(calc.vatAmount, curr)}</span>
            </div>
            <div className="flex justify-between py-2 font-medium">
              <span className="font-sans text-slate-800">Customer Goods Subtotal (Sent to PI / QT / CI)</span>
              <span className="text-slate-900">{formatMoney(calc.goodsSubtotal, curr)}</span>
            </div>
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-600">+ Freight &amp; Documentation (Separate Line)</span>
              <span className="text-slate-900">{formatMoney(calc.customerFreight, curr)}</span>
            </div>
            <div className="flex justify-between py-2.5 font-bold text-slate-900">
              <span className="font-sans">Customer Grand Total</span>
              <span>{formatMoney(calc.grandTotal, curr)}</span>
            </div>
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-600">Less: Supplier Goods + Base Freight/Doc</span>
              <span className="text-slate-700">
                {formatMoney(calc.totalSupplierCost + calc.freightWithDoc, curr)}
              </span>
            </div>
            <div className="flex justify-between py-2.5 font-bold text-emerald-700">
              <span className="font-sans">Net Gross Profit (Excl. VAT)</span>
              <span>{formatMoney(calc.grossProfit, curr)}</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
