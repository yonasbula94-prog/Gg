import React from 'react';
import { Lock } from 'lucide-react';
import { EXCEL_THEMES, MARKUP_RATES } from '../constants/themes';
import {
  FreightConfig,
  JobCalculationSummary,
  JobItemRow,
  MarkupRate,
  PlatformMode,
  ProcurementJob,
  ThemeVariantId,
} from '../types/procurement';
import { formatMoney, formatNumber, formatPct } from '../utils/calculations';
import { IPadJobEditor } from './IPadJobEditor';

interface PriceCalcViewProps {
  job: ProcurementJob;
  calc: JobCalculationSummary;
  themeId: ThemeVariantId;
  platformMode: PlatformMode;
  onUpdateItem: (rowNumber: number, patch: Partial<JobItemRow>) => void;
  onUpdateFreight: (patch: Partial<FreightConfig>) => void;
}

function parseOptionalNumber(val: string): number | null {
  if (val.trim() === '') return null;
  const n = Number(val);
  return Number.isNaN(n) ? null : n;
}

export const PriceCalcView: React.FC<PriceCalcViewProps> = ({
  job,
  calc,
  themeId,
  platformMode,
  onUpdateItem,
  onUpdateFreight,
}) => {
  const theme = EXCEL_THEMES[themeId];
  const curr = job.controls.currency;
  const activeMarkup = job.freight.selectedMarkupPct;
  const isPresetMarkup = MARKUP_RATES.some((r) => Math.abs(r - activeMarkup) < 0.0001);
  const activeMarkupPctDisplay = Math.round((activeMarkup * 100 + Number.EPSILON) * 100) / 100;

  return (
    <div className="space-y-6">
      {/* Confidential internal notice (flat, no thick left border) */}
      <div className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-red-200 bg-red-50/70 px-4 sm:px-5 py-3.5">
        <div className="flex items-start sm:items-center gap-2.5 text-sm text-red-900">
          <Lock className="w-4 h-4 text-red-700 shrink-0 mt-0.5 sm:mt-0" />
          <div>
            <span className="font-semibold">Internal Cost Sheet (PRICE CALC)</span>
            <span aria-hidden="true" className="mx-1.5 text-red-300">·</span>
            <span className="text-red-800">
              Holds supplier costs &amp; markup scenarios. Automatically excluded from customer exports.
            </span>
          </div>
        </div>
        <div className="text-xs text-red-800 font-mono-tabular">
          Formula Basis: Cost × (1 + Markup %) · Col R → PI Unit Price
        </div>
      </div>

      {/* Controls Strip: Markup Selector (E3) + Custom % + Freight Rates + Doc Charge */}
      <div className="rounded-lg border border-slate-200 bg-white p-5">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-end">
          {/* Active Markup Selector (E3) — Presets + Any Custom % */}
          <div className="lg:col-span-6 space-y-2">
            <label className="block text-xs font-medium text-slate-600">
              Active Markup % (Cell E3 — Tap Preset or Type Any Custom %)
            </label>
            <div className="flex flex-wrap sm:flex-nowrap items-center gap-2">
              <div className="flex-1 grid grid-cols-4 gap-1.5 p-1 rounded-lg bg-slate-100 min-w-[210px]">
                {MARKUP_RATES.map((rate) => {
                  const isSelected = Math.abs(activeMarkup - rate) < 0.0001;
                  return (
                    <button
                      key={rate}
                      type="button"
                      onClick={() => onUpdateFreight({ selectedMarkupPct: rate })}
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
                        onUpdateFreight({ selectedMarkupPct: 0 });
                        return;
                      }
                      const num = Number(raw);
                      if (!Number.isNaN(num) && num >= 0 && num <= 500) {
                        onUpdateFreight({ selectedMarkupPct: num / 100 });
                      }
                    }}
                    className="w-16 bg-transparent text-right text-sm font-mono-tabular font-bold text-blue-700 focus:outline-none"
                  />
                  <span className="text-xs font-semibold text-slate-500 ml-1">%</span>
                </div>
              </div>
            </div>
          </div>

          {/* Freight Rates & Doc Charge */}
          <div className="lg:col-span-6 grid grid-cols-1 sm:grid-cols-3 gap-4">
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
                onChange={(e) => onUpdateFreight({ ratePerCbm: Number(e.target.value) || 0 })}
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
                onChange={(e) => onUpdateFreight({ ratePerKg: Number(e.target.value) || 0 })}
                className="w-full min-h-[42px] rounded-lg border border-slate-200 bg-slate-50/50 px-3 py-2 text-sm font-mono-tabular text-blue-700 font-medium focus:border-slate-900 focus:bg-white focus:outline-none"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1.5">
                Doc Charge % (Default 10%)
              </label>
              <input
                type="number"
                inputMode="decimal"
                step="0.5"
                min="0"
                max="100"
                value={Math.round((job.freight.docChargePct * 100 + Number.EPSILON) * 100) / 100}
                onChange={(e) =>
                  onUpdateFreight({ docChargePct: Math.max(0, Number(e.target.value) || 0) / 100 })
                }
                className="w-full min-h-[42px] rounded-lg border border-slate-200 bg-slate-50/50 px-3 py-2 text-sm font-mono-tabular text-blue-700 font-medium focus:border-slate-900 focus:bg-white focus:outline-none"
              />
            </div>
          </div>
        </div>
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
              {formatPct(activeMarkup)} Markup → Grand Total {formatMoney(calc.grandTotal, curr)}
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-5 text-xs font-mono-tabular text-slate-700">
            <div>
              <span className="text-slate-400 block font-sans">Goods Subtotal</span>
              <span className="font-semibold text-slate-900">{formatMoney(calc.goodsSubtotal, curr)}</span>
            </div>
            <div>
              <span className="text-slate-400 block font-sans">VAT ({formatPct(job.controls.vatRate)})</span>
              <span className="font-semibold text-slate-900">{formatMoney(calc.vatAmount, curr)}</span>
            </div>
            <div>
              <span className="text-slate-400 block font-sans">Customer Freight</span>
              <span className="font-semibold text-slate-900">{formatMoney(calc.customerFreight, curr)}</span>
            </div>
            <div>
              <span className="text-slate-400 block font-sans">Est. Gross Profit</span>
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
              onClick={() => onUpdateFreight({ selectedMarkupPct: rate })}
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
                  <span>Goods Subtotal</span>
                  <span className="text-slate-900">{formatMoney(scen.goodsSubtotal)}</span>
                </div>
                <div className="flex justify-between">
                  <span>VAT ({formatPct(job.controls.vatRate)} Goods)</span>
                  <span className="text-slate-900">{formatMoney(scen.vatAmount)}</span>
                </div>
                <div className="flex justify-between">
                  <span>Customer Freight</span>
                  <span className="text-slate-900">{formatMoney(scen.customerFreight)}</span>
                </div>
                <div className="flex justify-between pt-1 border-t border-slate-100 font-semibold text-emerald-700">
                  <span>Est. Gross Profit</span>
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
            items={job.items}
            computedRows={calc.rows}
            currency={curr}
            themeId={themeId}
            onUpdateItem={onUpdateItem}
          />
        </div>
      ) : (
      <div className="rounded-lg border border-slate-200 bg-white overflow-hidden">
        <div className="flex flex-wrap items-center justify-between gap-2 px-5 py-3.5 border-b border-slate-200 bg-slate-50/70">
          <div>
            <h2 className="text-sm font-semibold text-slate-900">
              Costing &amp; Logistics Matrix (Excel Rows 17–36)
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Blue values are editable inputs · Black values are live formulas · Blank rows stay blank (no phantom zeros)
            </p>
          </div>
          <div className="text-xs text-slate-600 font-mono-tabular">
            Active Items: {calc.activeRowCount} / 20 · Supplier Total: {formatMoney(calc.totalSupplierCost, curr)}
          </div>
        </div>

        <div className="overflow-x-auto sheet-scroll">
          <table className="w-full border-collapse text-xs">
            <thead>
              <tr style={{ backgroundColor: theme.primary, color: theme.primaryText }}>
                <th className="py-2.5 px-2.5 text-center font-semibold border-r border-white/15 w-11">Row</th>
                <th className="py-2.5 px-3 text-left font-semibold min-w-[200px]">Offered Description</th>
                <th className="py-2.5 px-2 text-right font-semibold w-14">Qty</th>
                <th className="py-2.5 px-2 text-right font-semibold min-w-[96px]">Supplier Cost</th>
                <th className="py-2.5 px-2 text-right font-semibold min-w-[96px]">Total Cost</th>
                <th className="py-2.5 px-2 text-right font-semibold w-16">Boxes</th>
                <th className="py-2.5 px-1.5 text-right font-semibold w-16">L (cm)</th>
                <th className="py-2.5 px-1.5 text-right font-semibold w-16">W (cm)</th>
                <th className="py-2.5 px-1.5 text-right font-semibold w-16">H (cm)</th>
                <th className="py-2.5 px-2 text-right font-semibold w-20">CBM</th>
                <th className="py-2.5 px-2 text-right font-semibold w-18">Kg/Box</th>
                <th className="py-2.5 px-2 text-right font-semibold w-20">Gross KG</th>
                <th className="py-2.5 px-2 text-right font-semibold w-20">10%</th>
                <th className="py-2.5 px-2 text-right font-semibold w-20">15%</th>
                <th className="py-2.5 px-2 text-right font-semibold w-20">18%</th>
                <th className="py-2.5 px-2 text-right font-semibold w-20">20%</th>
                <th
                  className="py-2.5 px-2.5 text-right font-semibold min-w-[104px]"
                  style={{ backgroundColor: theme.accent }}
                >
                  Col R ({formatPct(activeMarkup)})
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
                    className={`${isInactive ? 'opacity-50' : ''} hover:bg-slate-50/90 transition-colors`}
                    style={!isInactive && idx % 2 === 1 ? { backgroundColor: theme.zebraRow } : undefined}
                  >
                    <td className="py-1.5 px-2.5 text-center text-slate-400 border-r border-slate-100">
                      {row.rowNumber}
                    </td>
                    <td className="py-1.5 px-3 font-sans text-slate-800 max-w-[240px] truncate" title={row.offeredDescription || row.requestedDescription}>
                      {row.offeredDescription || row.requestedDescription || <span className="text-slate-300">—</span>}
                    </td>
                    <td className="py-1.5 px-2 text-right text-slate-700">
                      {row.qty !== null ? row.qty : ''}
                    </td>
                    {/* Editable Supplier Unit Cost */}
                    <td className="p-1">
                      <input
                        type="number"
                        step="0.01"
                        min="0"
                        disabled={isInactive}
                        value={raw.supplierUnitCost ?? ''}
                        placeholder={isInactive ? '' : '0.00'}
                        onChange={(e) =>
                          onUpdateItem(row.rowNumber, {
                            supplierUnitCost: parseOptionalNumber(e.target.value),
                          })
                        }
                        className="w-full min-h-[34px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-right text-blue-700 font-medium focus:outline-none disabled:cursor-not-allowed"
                      />
                    </td>
                    <td className="py-1.5 px-2 text-right text-slate-900">
                      {formatMoney(row.supplierTotalCost)}
                    </td>
                    {/* Boxes */}
                    <td className="p-1">
                      <input
                        type="number"
                        step="1"
                        min="0"
                        disabled={isInactive}
                        value={raw.boxes ?? ''}
                        onChange={(e) =>
                          onUpdateItem(row.rowNumber, { boxes: parseOptionalNumber(e.target.value) })
                        }
                        className="w-full min-h-[34px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-1.5 py-1 text-right text-blue-700 focus:outline-none"
                      />
                    </td>
                    {/* L / W / H cm */}
                    <td className="p-1">
                      <input
                        type="number"
                        step="1"
                        min="0"
                        disabled={isInactive}
                        value={raw.boxLengthCm ?? ''}
                        onChange={(e) =>
                          onUpdateItem(row.rowNumber, { boxLengthCm: parseOptionalNumber(e.target.value) })
                        }
                        className="w-full min-h-[34px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-1.5 py-1 text-right text-blue-700 focus:outline-none"
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
                          onUpdateItem(row.rowNumber, { boxWidthCm: parseOptionalNumber(e.target.value) })
                        }
                        className="w-full min-h-[34px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-1.5 py-1 text-right text-blue-700 focus:outline-none"
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
                          onUpdateItem(row.rowNumber, { boxHeightCm: parseOptionalNumber(e.target.value) })
                        }
                        className="w-full min-h-[34px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-1.5 py-1 text-right text-blue-700 focus:outline-none"
                      />
                    </td>
                    <td className="py-1.5 px-2 text-right text-slate-700">
                      {row.cbm !== null ? formatNumber(row.cbm, 4) : ''}
                    </td>
                    {/* Kg/Box */}
                    <td className="p-1">
                      <input
                        type="number"
                        step="0.1"
                        min="0"
                        disabled={isInactive}
                        value={raw.kgPerBox ?? ''}
                        onChange={(e) =>
                          onUpdateItem(row.rowNumber, { kgPerBox: parseOptionalNumber(e.target.value) })
                        }
                        className="w-full min-h-[34px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-1.5 py-1 text-right text-blue-700 focus:outline-none"
                      />
                    </td>
                    <td className="py-1.5 px-2 text-right text-slate-700">
                      {row.grossKgCalc !== null ? formatNumber(row.grossKgCalc, 2) : ''}
                    </td>
                    {/* Scenarios */}
                    <td className={`py-1.5 px-2 text-right ${activeMarkup === 0.1 ? 'font-semibold text-slate-900' : 'text-slate-500'}`}>
                      {formatMoney(row.unitPrice10)}
                    </td>
                    <td className={`py-1.5 px-2 text-right ${activeMarkup === 0.15 ? 'font-semibold text-slate-900' : 'text-slate-500'}`}>
                      {formatMoney(row.unitPrice15)}
                    </td>
                    <td className={`py-1.5 px-2 text-right ${activeMarkup === 0.18 ? 'font-semibold text-slate-900' : 'text-slate-500'}`}>
                      {formatMoney(row.unitPrice18)}
                    </td>
                    <td className={`py-1.5 px-2 text-right ${activeMarkup === 0.2 ? 'font-semibold text-slate-900' : 'text-slate-500'}`}>
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
                <td colSpan={2} className="py-3 px-3 text-left font-sans text-slate-900">
                  TOTALS (Rows 17–36)
                </td>
                <td className="py-3 px-2 text-right text-slate-900">{calc.totalQty}</td>
                <td className="py-3 px-2" />
                <td className="py-3 px-2 text-right text-slate-900">{formatMoney(calc.totalSupplierCost)}</td>
                <td className="py-3 px-2 text-right text-slate-900">{calc.totalBoxes}</td>
                <td colSpan={3} />
                <td className="py-3 px-2 text-right text-slate-900">{formatNumber(calc.totalCbm, 4)}</td>
                <td />
                <td className="py-3 px-2 text-right text-slate-900">{formatNumber(calc.totalGrossKgCalc, 2)}</td>
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

      {/* Freight Basis Comparison & Margin Breakdown */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="rounded-lg border border-slate-200 bg-white p-5">
          <h3 className="text-sm font-semibold text-slate-900">
            Freight Calculation Breakdown (MAX Basis)
          </h3>
          <p className="text-xs text-slate-500 mt-0.5">
            Higher of CBM or Gross KG freight → add Documentation Charge → apply selected {formatPct(activeMarkup)} markup.
          </p>

          <div className="mt-4 divide-y divide-slate-100 text-sm font-mono-tabular">
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-600">
                Volume Freight ({formatNumber(calc.totalCbm, 4)} CBM × {formatMoney(job.freight.ratePerCbm)})
                {calc.freightBasis === 'CBM' && (
                  <span className="ml-2 font-semibold text-slate-900">· Governing</span>
                )}
              </span>
              <span className={calc.freightBasis === 'CBM' ? 'font-semibold text-slate-900' : 'text-slate-500'}>
                {formatMoney(calc.freightByCbm, curr)}
              </span>
            </div>
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-600">
                Weight Freight ({formatNumber(calc.totalGrossKgCalc, 2)} KG × {formatMoney(job.freight.ratePerKg)})
                {calc.freightBasis === 'KG' && (
                  <span className="ml-2 font-semibold text-slate-900">· Governing</span>
                )}
              </span>
              <span className={calc.freightBasis === 'KG' ? 'font-semibold text-slate-900' : 'text-slate-500'}>
                {formatMoney(calc.freightByKg, curr)}
              </span>
            </div>
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-700 font-medium">Base Freight (MAX)</span>
              <span className="font-medium text-slate-900">{formatMoney(calc.baseFreight, curr)}</span>
            </div>
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-600">
                + Documentation Charge ({formatPct(job.freight.docChargePct)})
              </span>
              <span className="text-slate-900">{formatMoney(calc.docChargeAmount, curr)}</span>
            </div>
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-600">
                + Markup on Freight + Doc ({formatPct(activeMarkup)})
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
          <h3 className="text-sm font-semibold text-slate-900">
            Job Profitability &amp; Customer Invoice Summary
          </h3>
          <p className="text-xs text-slate-500 mt-0.5">
            VAT ({formatPct(job.controls.vatRate)}) is computed on goods subtotal only; freight appears as a separate line after VAT.
          </p>

          <div className="mt-4 divide-y divide-slate-100 text-sm font-mono-tabular">
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-600">Customer Goods Subtotal</span>
              <span className="text-slate-900">{formatMoney(calc.goodsSubtotal, curr)}</span>
            </div>
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-600">
                VAT ({formatPct(job.controls.vatRate)} on Goods Only)
              </span>
              <span className="text-slate-900">{formatMoney(calc.vatAmount, curr)}</span>
            </div>
            <div className="flex justify-between py-2">
              <span className="font-sans text-slate-600">Freight &amp; Documentation (Separate Line)</span>
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
