import React from 'react';
import { Plus, Trash2 } from 'lucide-react';
import { MARKUP_RATES, UOM_OPTIONS } from '../constants/themes';
import {
  ComputedRow,
  ConfirmDialogRequest,
  CustomColumnDef,
  JobItemRow,
  ThemePalette,
} from '../types/procurement';
import { formatMoney, formatNumber, formatPct } from '../utils/calculations';

interface IPadJobEditorProps {
  mode: 'entry' | 'costing' | 'logistics';
  items: JobItemRow[];
  computedRows: ComputedRow[];
  customColumns?: CustomColumnDef[];
  currency: string;
  theme: ThemePalette;
  activeMarkupPct?: number;
  effectiveVatRate?: number;
  disabled?: boolean;
  onUpdateItem: (rowNumber: number, patch: Partial<JobItemRow>) => void;
  onRequestConfirm?: (req: ConfirmDialogRequest) => void;
}

function parseOptionalNumber(val: string): number | null {
  if (val.trim() === '') return null;
  const n = Number(val);
  return Number.isNaN(n) ? null : n;
}

export const IPadJobEditor: React.FC<IPadJobEditorProps> = ({
  mode,
  items,
  computedRows,
  customColumns = [],
  currency,
  theme,
  activeMarkupPct = 0.2,
  effectiveVatRate = 0.05,
  disabled = false,
  onUpdateItem,
  onRequestConfirm,
}) => {
  const isCustomMarkup = !MARKUP_RATES.some((r) => Math.abs(r - activeMarkupPct) < 0.0001);

  // Show all populated rows (or first row if none populated yet) so iPad users never scroll through 16 blank rows
  const populatedIndices = computedRows
    .map((r, i) => (!r.isEmpty ? i : -1))
    .filter((i) => i !== -1);

  const visibleIndices =
    populatedIndices.length > 0 ? populatedIndices : [0];

  const firstEmptyIdx = computedRows.findIndex((r) => r.isEmpty);

  const handleAddRow = () => {
    if (firstEmptyIdx === -1 || disabled) return;
    const targetRow = items[firstEmptyIdx].rowNumber;
    onUpdateItem(targetRow, {
      requestedDescription: 'New Item',
      offeredDescription: '',
      uom: 'PCS',
      qty: 1,
    });
  };

  const clearRowNow = (rowNumber: number) => {
    onUpdateItem(rowNumber, {
      requestedDescription: '',
      offeredDescription: '',
      uom: '',
      qty: null,
      customFields: {},
      supplierUnitCost: null,
      boxes: null,
      boxLengthCm: null,
      boxWidthCm: null,
      boxHeightCm: null,
      kgPerBox: null,
      cartonNo: '',
      netWtKg: null,
      grossWtKg: null,
    });
  };

  const handleClearRow = (rowNumber: number, itemLabel: string) => {
    if (disabled) return;
    if (onRequestConfirm) {
      onRequestConfirm({
        title: `Clear Row ${rowNumber}?`,
        message: `This will clear "${itemLabel || `Row ${rowNumber}`}" along with its cost, boxes, and packing weights.`,
        confirmLabel: 'Clear Row',
        destructive: true,
        onConfirm: () => clearRowNow(rowNumber),
      });
    } else {
      clearRowNow(rowNumber);
    }
  };

  return (
    <div className="space-y-4">
      {visibleIndices.map((idx, displayPos) => {
        const raw = items[idx];
        const calcRow = computedRows[idx];

        return (
          <div
            key={raw.rowNumber}
            className="rounded-xl border border-slate-200 bg-white p-4 space-y-4"
          >
            {/* Card Header */}
            <div className="flex items-center justify-between gap-3 pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2 text-xs text-slate-600 font-mono-tabular">
                <span className="font-bold text-slate-900 text-sm">
                  Item #{displayPos + 1}
                </span>
                <span aria-hidden="true">·</span>
                <span>Excel Row {raw.rowNumber}</span>
              </div>

              <div className="flex items-center gap-3">
                {mode === 'entry' && calcRow.activeTotalPrice !== null && (
                  <span className="text-sm font-bold font-mono-tabular text-slate-900">
                    {formatMoney(calcRow.activeTotalPrice, currency)}
                  </span>
                )}
                {mode === 'costing' && calcRow.activeUnitPrice !== null && (
                  <span className="text-sm font-bold font-mono-tabular" style={{ color: theme.primary }}>
                    Col R ({formatPct(activeMarkupPct)}{effectiveVatRate > 0 ? ` + ${formatPct(effectiveVatRate)} VAT` : ''}): {formatMoney(calcRow.activeUnitPrice, currency)}
                  </span>
                )}
                {mode === 'entry' && !disabled && !calcRow.isEmpty && (
                  <button
                    type="button"
                    onClick={() =>
                      handleClearRow(
                        raw.rowNumber,
                        raw.offeredDescription || raw.requestedDescription
                      )
                    }
                    className="p-2 rounded-lg text-slate-400 hover:text-red-600 hover:bg-red-50 min-h-[40px] min-w-[40px] flex items-center justify-center cursor-pointer"
                    aria-label={`Clear row ${raw.rowNumber}`}
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                )}
              </div>
            </div>

            {mode === 'entry' && (
              <div className="space-y-3">
                {customColumns.length > 0 && (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    {customColumns.map((col) => (
                      <div key={col.id}>
                        <label className="block text-xs font-medium text-slate-600 mb-1">
                          {col.label}
                        </label>
                        <input
                          type="text"
                          disabled={disabled}
                          value={raw.customFields?.[col.id] || ''}
                          placeholder={`Enter ${col.label}...`}
                          onChange={(e) =>
                            onUpdateItem(raw.rowNumber, {
                              customFields: {
                                ...(raw.customFields || {}),
                                [col.id]: e.target.value,
                              },
                            })
                          }
                          className="w-full min-h-[46px] rounded-lg border border-slate-200 px-3.5 py-2.5 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                        />
                      </div>
                    ))}
                  </div>
                )}
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1">
                    Requested Description (End-User Wording)
                  </label>
                  <input
                    type="text"
                    disabled={disabled}
                    value={raw.requestedDescription}
                    placeholder="Type requested specification..."
                    onChange={(e) =>
                      onUpdateItem(raw.rowNumber, { requestedDescription: e.target.value })
                    }
                    className="w-full min-h-[46px] rounded-lg border border-slate-200 px-3.5 py-2.5 text-sm text-slate-900 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1">
                    Offered Description (Supplier Wording We Sell)
                  </label>
                  <input
                    type="text"
                    disabled={disabled}
                    value={raw.offeredDescription}
                    placeholder="Type supplier offered model/spec..."
                    onChange={(e) =>
                      onUpdateItem(raw.rowNumber, { offeredDescription: e.target.value })
                    }
                    className="w-full min-h-[46px] rounded-lg border border-slate-200 px-3.5 py-2.5 text-sm text-slate-900 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">UoM</label>
                    <select
                      disabled={disabled}
                      value={raw.uom}
                      onChange={(e) => onUpdateItem(raw.rowNumber, { uom: e.target.value })}
                      className="w-full min-h-[46px] rounded-lg border border-slate-200 bg-white px-3 py-2.5 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                    >
                      <option value="">—</option>
                      {UOM_OPTIONS.map((u) => (
                        <option key={u} value={u}>{u}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">Qty</label>
                    <input
                      type="number"
                      inputMode="decimal"
                      min="0"
                      step="any"
                      disabled={disabled}
                      value={raw.qty ?? ''}
                      onChange={(e) =>
                        onUpdateItem(raw.rowNumber, { qty: parseOptionalNumber(e.target.value) })
                      }
                      className="w-full min-h-[46px] rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-mono-tabular text-right text-blue-700 font-semibold focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-500 mb-1">
                      Unit Price (Col R)
                    </label>
                    <div className="min-h-[46px] rounded-lg border border-slate-100 bg-slate-50 px-3 py-2.5 text-sm font-mono-tabular text-right text-slate-700 flex items-center justify-end">
                      {formatMoney(calcRow.activeUnitPrice) || '—'}
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-500 mb-1">
                      Line Total
                    </label>
                    <div className="min-h-[46px] rounded-lg border border-slate-100 bg-slate-50 px-3 py-2.5 text-sm font-mono-tabular text-right font-semibold text-slate-900 flex items-center justify-end">
                      {formatMoney(calcRow.activeTotalPrice) || '—'}
                    </div>
                  </div>
                </div>
              </div>
            )}

            {mode === 'costing' && (
              <div className="space-y-3">
                <div className="text-sm font-medium text-slate-900">
                  {calcRow.offeredDescription || calcRow.requestedDescription || 'Unfilled Row'}
                  {calcRow.qty !== null && (
                    <span className="ml-2 text-xs font-mono-tabular text-slate-500">
                      · Qty {calcRow.qty} {calcRow.uom}
                    </span>
                  )}
                </div>

                <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">
                      Supplier Unit Cost ({currency})
                    </label>
                    <input
                      type="number"
                      inputMode="decimal"
                      step="0.01"
                      min="0"
                      value={raw.supplierUnitCost ?? ''}
                      onChange={(e) =>
                        onUpdateItem(raw.rowNumber, {
                          supplierUnitCost: parseOptionalNumber(e.target.value),
                        })
                      }
                      className="w-full min-h-[46px] rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-mono-tabular text-right text-blue-700 font-semibold focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">Boxes</label>
                    <input
                      type="number"
                      inputMode="numeric"
                      min="0"
                      value={raw.boxes ?? ''}
                      onChange={(e) =>
                        onUpdateItem(raw.rowNumber, { boxes: parseOptionalNumber(e.target.value) })
                      }
                      className="w-full min-h-[46px] rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-mono-tabular text-right text-blue-700 focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">Kg / Box</label>
                    <input
                      type="number"
                      inputMode="decimal"
                      step="0.1"
                      min="0"
                      value={raw.kgPerBox ?? ''}
                      onChange={(e) =>
                        onUpdateItem(raw.rowNumber, { kgPerBox: parseOptionalNumber(e.target.value) })
                      }
                      className="w-full min-h-[46px] rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-mono-tabular text-right text-blue-700 focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">L (cm)</label>
                    <input
                      type="number"
                      inputMode="decimal"
                      min="0"
                      value={raw.boxLengthCm ?? ''}
                      onChange={(e) =>
                        onUpdateItem(raw.rowNumber, { boxLengthCm: parseOptionalNumber(e.target.value) })
                      }
                      className="w-full min-h-[46px] rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-mono-tabular text-right text-blue-700 focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">W (cm)</label>
                    <input
                      type="number"
                      inputMode="decimal"
                      min="0"
                      value={raw.boxWidthCm ?? ''}
                      onChange={(e) =>
                        onUpdateItem(raw.rowNumber, { boxWidthCm: parseOptionalNumber(e.target.value) })
                      }
                      className="w-full min-h-[46px] rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-mono-tabular text-right text-blue-700 focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">H (cm)</label>
                    <input
                      type="number"
                      inputMode="decimal"
                      min="0"
                      value={raw.boxHeightCm ?? ''}
                      onChange={(e) =>
                        onUpdateItem(raw.rowNumber, { boxHeightCm: parseOptionalNumber(e.target.value) })
                      }
                      className="w-full min-h-[46px] rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-mono-tabular text-right text-blue-700 focus:border-slate-900 focus:outline-none"
                    />
                  </div>
                </div>

                <div className="pt-2 border-t border-slate-100 grid grid-cols-2 sm:grid-cols-6 gap-2 text-xs font-mono-tabular">
                  <div>
                    <span className="text-slate-400 block">CBM</span>
                    <span className="font-semibold text-slate-800">{calcRow.cbm !== null ? formatNumber(calcRow.cbm, 4) : '—'}</span>
                  </div>
                  <div>
                    <span className="text-slate-400 block">Gross KG</span>
                    <span className="font-semibold text-slate-800">{calcRow.grossKgCalc !== null ? formatNumber(calcRow.grossKgCalc, 2) : '—'}</span>
                  </div>
                  <div>
                    <span className="text-slate-400 block">10%</span>
                    <span className="text-slate-700">{formatMoney(calcRow.unitPrice10) || '—'}</span>
                  </div>
                  <div>
                    <span className="text-slate-400 block">15%</span>
                    <span className="text-slate-700">{formatMoney(calcRow.unitPrice15) || '—'}</span>
                  </div>
                  <div>
                    <span className="text-slate-400 block">18%</span>
                    <span className="text-slate-700">{formatMoney(calcRow.unitPrice18) || '—'}</span>
                  </div>
                  <div>
                    <span className="text-slate-400 block">
                      {isCustomMarkup ? `Custom ${formatPct(activeMarkupPct)}` : '20%'}
                    </span>
                    <span className={isCustomMarkup ? 'font-bold text-slate-900' : 'text-slate-700'}>
                      {formatMoney(isCustomMarkup ? calcRow.activeUnitPrice : calcRow.unitPrice20) || '—'}
                    </span>
                  </div>
                </div>
              </div>
            )}

            {mode === 'logistics' && (
              <div className="space-y-3">
                <div className="text-sm font-medium text-slate-900">
                  {calcRow.offeredDescription || calcRow.requestedDescription}
                  <span className="ml-2 text-xs font-mono-tabular text-slate-500">
                    · Qty {calcRow.qty} {calcRow.uom}
                  </span>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">Carton No.</label>
                    <input
                      type="text"
                      disabled={disabled}
                      value={raw.cartonNo}
                      placeholder="CTN 01-04"
                      onChange={(e) => onUpdateItem(raw.rowNumber, { cartonNo: e.target.value })}
                      className="w-full min-h-[46px] rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">Net Wt (KG)</label>
                    <input
                      type="number"
                      inputMode="decimal"
                      step="0.1"
                      min="0"
                      disabled={disabled}
                      value={raw.netWtKg ?? ''}
                      onChange={(e) =>
                        onUpdateItem(raw.rowNumber, { netWtKg: parseOptionalNumber(e.target.value) })
                      }
                      className="w-full min-h-[46px] rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-mono-tabular text-right text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">Gross Wt (KG)</label>
                    <input
                      type="number"
                      inputMode="decimal"
                      step="0.1"
                      min="0"
                      disabled={disabled}
                      value={raw.grossWtKg ?? ''}
                      onChange={(e) =>
                        onUpdateItem(raw.rowNumber, { grossWtKg: parseOptionalNumber(e.target.value) })
                      }
                      className="w-full min-h-[46px] rounded-lg border border-slate-200 px-3 py-2.5 text-sm font-mono-tabular text-right text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                    />
                  </div>
                </div>
              </div>
            )}
          </div>
        );
      })}

      {mode === 'entry' && !disabled && firstEmptyIdx !== -1 && (
        <button
          type="button"
          onClick={handleAddRow}
          className="w-full min-h-[48px] flex items-center justify-center gap-2 rounded-xl border-2 border-dashed border-slate-300 bg-white text-sm font-semibold text-slate-700 hover:border-slate-900 hover:text-slate-900 transition-colors cursor-pointer"
        >
          <Plus className="w-4 h-4" />
          <span>Add Line Item (Row {items[firstEmptyIdx].rowNumber} of 36)</span>
        </button>
      )}
    </div>
  );
};
