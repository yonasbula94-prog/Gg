import React, { useState } from 'react';
import {
  ArrowRight,
  Calculator,
  FileUp,
  Lock,
  Sparkles,
  Trash2,
} from 'lucide-react';
import { CURRENCY_OPTIONS, INCOTERM_OPTIONS, UOM_OPTIONS } from '../constants/themes';
import { isJobLocked } from '../utils/amendments';
import { useRowSelection } from '../hooks/useRowSelection';
import {
  ConfirmDialogRequest,
  JobCalculationSummary,
  JobControls,
  JobItemRow,
  ProcurementJob,
  SupplierQuoteData,
  ThemePalette,
  WorkspaceTab,
} from '../types/procurement';
import { formatMoney, isRowEmpty } from '../utils/calculations';
import { HelpTip } from './HelpTip';

interface InternalQuoteViewProps {
  job: ProcurementJob;
  amending: boolean;
  calc: JobCalculationSummary;
  theme: ThemePalette;
  onUpdateSupplierQuote: (patch: Partial<SupplierQuoteData>) => void;
  onUpdateControls: (patch: Partial<JobControls>) => void;
  onUpdateItem: (rowNumber: number, patch: Partial<JobItemRow>) => void;
  onOpenImportFile: () => void;
  onNavigateTab: (tab: WorkspaceTab) => void;
  onRequestConfirm: (req: ConfirmDialogRequest) => void;
  onNotify: (msg: string) => void;
}

function parseOptionalNumber(val: string): number | null {
  if (val.trim() === '') return null;
  const n = Number(val);
  if (!Number.isFinite(n) || n < 0) return null;
  return n;
}

export const InternalQuoteView: React.FC<InternalQuoteViewProps> = ({
  job,
  amending,
  calc,
  theme,
  onUpdateSupplierQuote,
  onUpdateControls,
  onUpdateItem,
  onOpenImportFile,
  onNavigateTab,
  onRequestConfirm,
  onNotify,
}) => {
  const sq = job.supplierQuote;
  const curr = sq.currency || job.controls.currency || 'USD';
  const filledCount = job.items.filter((it) => !isRowEmpty(it)).length;

  const handleClearRow = (row: JobItemRow) => {
    const clearNow = () =>
      onUpdateItem(row.rowNumber, {
        requestedDescription: '',
        offeredDescription: '',
        uom: '',
        qty: null,
        supplierUnitCost: null,
      });

    if (row.requestedDescription || row.offeredDescription || row.supplierUnitCost !== null) {
      onRequestConfirm({
        title: `Clear Row ${row.rowNumber - 16}?`,
        message: 'This clears the item description, quantity, and supplier unit cost for this row.',
        confirmLabel: 'Clear Row',
        destructive: true,
        onConfirm: clearNow,
      });
    } else {
      clearNow();
    }
  };

  const handleSyncTermsToQuotation = () => {
    onUpdateControls({
      currency: sq.currency || job.controls.currency,
      deliveryLeadTime: sq.deliveryLeadTime || job.controls.deliveryLeadTime,
    });
    onNotify('Synced currency and lead time from Internal Supplier Quote to Quotation.');
  };

  // Issued jobs are locked: a disabled fieldset disables every control inside, unless the job is open for amendment
  const locked = isJobLocked(job) && !amending;
  const rowSel = useRowSelection<number>();
  const [bulkUom, setBulkUom] = useState('');
  const itemRowIds = job.items.filter((it) => it.requestedDescription || it.offeredDescription || (it.qty ?? 0) > 0).map((it) => it.rowNumber);
  const applyBulkItems = (patch: Partial<JobItemRow>) => rowSel.selected.forEach((rn) => onUpdateItem(rn, patch));
  return (
    <fieldset disabled={locked} className="space-y-6 min-w-0 border-0 p-0 m-0">
      {locked && (
        <div className="rounded-lg border border-amber-200 bg-amber-50 px-4 py-2.5 text-xs text-amber-900">
          This job has an issued document, so supplier details and costs are locked. Use Unlock on the document to amend.
        </div>
      )}
      {/* Internal Confidentiality Strip & Primary Flow Actions */}
      <div className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-slate-200 bg-white px-5 py-3.5">
        <div className="flex items-center gap-2.5 min-w-0">
          <Lock className="w-4 h-4 text-slate-600 shrink-0" />
          <div className="min-w-0">
            <div className="flex items-center gap-1.5">
              <h2 className="text-sm font-semibold text-slate-900">
                Internal Quotation (Supplier Offer &amp; Terms)
              </h2>
              <HelpTip text="Confidential supplier intake sheet. Supplier details automatically pre-fill Supplier PO, and supplier unit costs & offered items flow directly into Price Calc and Customer Quotation." />
            </div>
            <p className="text-xs text-slate-500">
              Internal Only · Feeds Price Calc &amp; Supplier PO · Never Sent to Client
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={onOpenImportFile}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-xs font-medium text-slate-800 hover:bg-slate-50 cursor-pointer min-h-[40px] whitespace-nowrap"
          >
            <FileUp className="w-3.5 h-3.5 text-slate-600" />
            <span>Import Supplier Quote (.pdf / .xlsx)</span>
          </button>

          <button
            type="button"
            onClick={() => onNavigateTab('PRICE_CALC')}
            className="inline-flex items-center gap-1.5 rounded-lg px-4 py-2 text-xs font-semibold text-white hover:opacity-90 cursor-pointer min-h-[40px] whitespace-nowrap"
            style={{ backgroundColor: theme.primary, color: theme.primaryText }}
          >
            <Calculator className="w-3.5 h-3.5" />
            <span>Open Price Calc</span>
            <ArrowRight className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* Supplier Identity & Commercial Terms */}
      <div className="rounded-lg border border-slate-200 bg-white p-5 space-y-5">
        <div className="flex flex-wrap items-center justify-between gap-2 pb-3 border-b border-slate-100">
          <div className="flex items-center gap-1.5">
            <h3 className="text-sm font-semibold text-slate-900">
              Supplier Details &amp; Commercial Terms
            </h3>
            <HelpTip text="Changes to Supplier Name, Address, and Quote Reference automatically sync to Supplier PO (PO) without re-typing." />
          </div>
          <button
            type="button"
            onClick={handleSyncTermsToQuotation}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[34px] whitespace-nowrap"
          >
            <Sparkles className="w-3.5 h-3.5 text-slate-500" />
            <span>Apply Currency &amp; Lead Time to Quotation</span>
          </button>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
          <div className="lg:col-span-2">
            <label className="block text-xs font-medium text-slate-600 mb-1.5">
              Supplier Company Name
            </label>
            <input
              type="text"
              value={sq.supplierName}
              onChange={(e) => onUpdateSupplierQuote({ supplierName: e.target.value })}
              placeholder="e.g. Al-Futtaim Industrial Supplies LLC"
              className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-blue-700 focus:border-slate-900 focus:outline-none"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1.5">
              Supplier Quote Ref.
            </label>
            <input
              type="text"
              value={sq.supplierQuoteRef}
              onChange={(e) => onUpdateSupplierQuote({ supplierQuoteRef: e.target.value })}
              placeholder="e.g. SQ-2026-084"
              className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1.5">
              Quote Date
            </label>
            <input
              type="date"
              value={sq.quoteDate}
              onChange={(e) => onUpdateSupplierQuote({ quoteDate: e.target.value })}
              className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none"
            />
          </div>

          <div className="lg:col-span-2">
            <label className="block text-xs font-medium text-slate-600 mb-1.5">
              Supplier Address
            </label>
            <input
              type="text"
              value={sq.supplierAddress}
              onChange={(e) => onUpdateSupplierQuote({ supplierAddress: e.target.value })}
              placeholder="e.g. Jebel Ali Free Zone, Dubai, UAE"
              className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1.5">
              Supplier Contact Person
            </label>
            <input
              type="text"
              value={sq.supplierContact}
              onChange={(e) => onUpdateSupplierQuote({ supplierContact: e.target.value })}
              placeholder="e.g. Ahmed Al-Mansoor"
              className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1.5">
              Quote Valid Until
            </label>
            <input
              type="date"
              value={sq.validUntil}
              onChange={(e) => onUpdateSupplierQuote({ validUntil: e.target.value })}
              className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1.5">
              Supplier Currency
            </label>
            <select
              value={sq.currency}
              onChange={(e) => onUpdateSupplierQuote({ currency: e.target.value })}
              className="w-full min-h-[40px] rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none"
            >
              {CURRENCY_OPTIONS.map((c) => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1.5">
              Supplier Incoterm
            </label>
            <select
              value={sq.incoterm}
              onChange={(e) => onUpdateSupplierQuote({ incoterm: e.target.value })}
              className="w-full min-h-[40px] rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none"
            >
              {INCOTERM_OPTIONS.map((inc) => (
                <option key={inc} value={inc}>
                  {inc}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1.5">
              Supplier Payment Terms
            </label>
            <input
              type="text"
              value={sq.paymentTerms}
              onChange={(e) => onUpdateSupplierQuote({ paymentTerms: e.target.value })}
              placeholder="e.g. 100% TT Prior to Dispatch"
              className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1.5">
              Supplier Lead Time
            </label>
            <input
              type="text"
              value={sq.deliveryLeadTime}
              onChange={(e) => onUpdateSupplierQuote({ deliveryLeadTime: e.target.value })}
              placeholder="e.g. 2–3 Weeks Ex-Stock"
              className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none"
            />
          </div>

          <div className="md:col-span-2 lg:col-span-4">
            <label className="block text-xs font-medium text-slate-600 mb-1.5">
              Internal Supplier Notes / Deviations / Warranty
            </label>
            <input
              type="text"
              value={sq.notes}
              onChange={(e) => onUpdateSupplierQuote({ notes: e.target.value })}
              placeholder="Optional internal evaluation notes, warranty terms, packing details..."
              className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-800 focus:border-slate-900 focus:outline-none"
            />
          </div>
        </div>
      </div>

      {/* Supplier Offer Line Items Schedule (Direct Sync to Price Calc & Quotation) */}
      <div className="rounded-lg border border-slate-200 bg-white overflow-hidden">
        <div className="px-5 py-3.5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <h3 className="text-sm font-semibold text-slate-900">
              Supplier Offered Items &amp; Net Unit Costs
            </h3>
            <HelpTip text="Every item and supplier unit cost entered here flows straight into Price Calc (Rows 17–36) to compute selling prices, margins, and freight." />
          </div>
          <div className="flex items-center gap-4 text-xs font-mono-tabular text-slate-600">
            <span>{filledCount} / 20 Active Lines</span>
            <span>·</span>
            <span className="font-semibold text-slate-900">
              Total Supplier Cost: {formatMoney(calc.totalSupplierCost, curr)}
            </span>
          </div>
        </div>

        {rowSel.count > 0 && (
          <div className="no-print flex flex-wrap items-center gap-2 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-xs">
            <span className="font-semibold text-slate-900">{rowSel.count} row{rowSel.count === 1 ? '' : 's'} selected</span>
            <select value={bulkUom} onChange={(e) => setBulkUom(e.target.value)} className="min-h-[32px] rounded border border-slate-200 bg-white px-2">
              <option value="">Set UoM…</option>
              {UOM_OPTIONS.map((u) => (<option key={u} value={u}>{u}</option>))}
            </select>
            <button type="button" disabled={!bulkUom} onClick={() => { applyBulkItems({ uom: bulkUom }); setBulkUom(''); }} className="rounded border border-slate-200 bg-white px-2.5 py-1.5 font-medium text-slate-800 hover:bg-slate-100 disabled:opacity-50 cursor-pointer">Apply UoM</button>
            <button type="button" onClick={() => applyBulkItems({ offeredDescription: '', uom: '', qty: null, supplierUnitCost: null })} className="rounded border border-red-200 bg-white px-2.5 py-1.5 font-medium text-red-700 hover:bg-red-50 cursor-pointer">Clear selected offers</button>
            <button type="button" onClick={rowSel.clear} className="ml-auto text-slate-600 underline cursor-pointer">Deselect all</button>
          </div>
        )}

        <div className="overflow-x-auto sheet-scroll">
          <table className="w-full border-collapse text-xs">
            <thead>
              <tr style={{ backgroundColor: theme.primary, color: theme.primaryText }}>
                <th className="py-2.5 px-2 text-center w-8">
                  <input type="checkbox" aria-label="Select all rows" disabled={itemRowIds.length === 0 || locked} checked={rowSel.allChecked(itemRowIds)} onChange={() => rowSel.toggleAll(itemRowIds)} />
                </th>
                <th className="py-2.5 px-2.5 text-center font-semibold w-10">#</th>
                <th className="py-2.5 px-3 text-left font-semibold min-w-[210px]">
                  Client Requested Item
                </th>
                <th className="py-2.5 px-3 text-left font-semibold min-w-[210px]">
                  Supplier Offered Spec / Model
                </th>
                <th className="py-2.5 px-2.5 text-center font-semibold w-24">UoM</th>
                <th className="py-2.5 px-2.5 text-right font-semibold w-24">Qty</th>
                <th className="py-2.5 px-2.5 text-right font-semibold w-36">
                  Supplier Unit Cost ({curr})
                </th>
                <th className="py-2.5 px-3 text-right font-semibold w-36">
                  Supplier Total ({curr})
                </th>
                <th className="py-2.5 px-2 text-center font-semibold w-10"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {job.items.map((item, idx) => {
                const comp = calc.rows[idx];
                const isAlt = idx % 2 === 1;
                return (
                  <tr
                    key={item.rowNumber}
                    style={isAlt ? { backgroundColor: theme.zebraRow } : undefined}
                    className="hover:bg-slate-50/70"
                  >
                    <td className="py-1.5 px-2 text-center">
                      <input type="checkbox" aria-label={`Select row ${item.rowNumber}`} disabled={locked || !itemRowIds.includes(item.rowNumber)} checked={rowSel.has(item.rowNumber)} onChange={() => rowSel.toggle(item.rowNumber)} />
                    </td>
                    <td className="py-1.5 px-2 text-center font-mono-tabular text-slate-400">
                      {idx + 1}
                    </td>
                    <td className="py-1 px-1.5">
                      <input
                        type="text"
                        value={item.requestedDescription}
                        onChange={(e) =>
                          onUpdateItem(item.rowNumber, {
                            requestedDescription: e.target.value,
                          })
                        }
                        placeholder={idx === 0 ? 'e.g. LED Runway Edge Light' : ''}
                        className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-xs text-slate-900 focus:outline-none"
                      />
                    </td>
                    <td className="py-1 px-1.5">
                      <input
                        type="text"
                        value={item.offeredDescription}
                        onChange={(e) =>
                          onUpdateItem(item.rowNumber, {
                            offeredDescription: e.target.value,
                          })
                        }
                        placeholder={idx === 0 ? 'Supplier offered part / brand...' : ''}
                        className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-xs text-blue-700 focus:outline-none"
                      />
                    </td>
                    <td className="py-1 px-1.5">
                      <select
                        value={item.uom}
                        onChange={(e) => onUpdateItem(item.rowNumber, { uom: e.target.value })}
                        className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-1.5 py-1 text-xs font-mono-tabular text-center text-slate-800 focus:outline-none"
                      >
                        <option value="">—</option>
                        {UOM_OPTIONS.map((u) => (
                          <option key={u} value={u}>
                            {u}
                          </option>
                        ))}
                      </select>
                    </td>
                    <td className="py-1 px-1.5">
                      <input
                        type="number"
                        min="0"
                        step="any"
                        value={item.qty ?? ''}
                        onChange={(e) =>
                          onUpdateItem(item.rowNumber, {
                            qty: parseOptionalNumber(e.target.value),
                          })
                        }
                        className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-right font-mono-tabular text-blue-700 font-medium focus:outline-none"
                      />
                    </td>
                    <td className="py-1 px-1.5">
                      <input
                        type="number"
                        min="0"
                        step="0.01"
                        value={item.supplierUnitCost ?? ''}
                        onChange={(e) =>
                          onUpdateItem(item.rowNumber, {
                            supplierUnitCost: parseOptionalNumber(e.target.value),
                          })
                        }
                        placeholder="0.00"
                        className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-right font-mono-tabular text-blue-700 font-semibold focus:outline-none"
                      />
                    </td>
                    <td className="py-1.5 px-3 text-right font-mono-tabular font-semibold text-slate-900">
                      {comp.supplierTotalCost !== null
                        ? formatMoney(comp.supplierTotalCost)
                        : '—'}
                    </td>
                    <td className="py-1 px-1 text-center">
                      {!isRowEmpty(item) && (
                        <button
                          type="button"
                          onClick={() => handleClearRow(item)}
                          title={`Clear row ${idx + 1}`}
                          aria-label={`Clear row ${idx + 1}`}
                          className="p-1.5 rounded text-slate-400 hover:text-red-600 hover:bg-red-50 cursor-pointer"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr
                className="border-t-2 font-semibold"
                style={{
                  backgroundColor: theme.totalBg,
                  borderColor: theme.strongBorder,
                  color: theme.primary,
                }}
              >
                <td colSpan={5} className="py-3 px-4 text-right text-xs uppercase tracking-wide">
                  Total Supplier Offer Cost
                </td>
                <td className="py-3 px-2.5 text-right font-mono-tabular">{calc.totalQty || '—'}</td>
                <td className="py-3 px-2.5 text-right font-mono-tabular text-slate-500">
                  Markup {Math.round(job.freight.selectedMarkupPct * 100)}%
                </td>
                <td className="py-3 px-3 text-right font-mono-tabular text-sm font-bold">
                  {formatMoney(calc.totalSupplierCost, curr)}
                </td>
                <td></td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      {/* Bottom Flow Navigation Strip */}
      <div className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-slate-200 bg-white px-5 py-3.5">
        <div className="text-xs text-slate-600">
          Estimated Customer Grand Total (from Price Calc):{' '}
          <strong className="font-mono-tabular text-slate-900">
            {formatMoney(calc.grandTotal, job.controls.currency)}
          </strong>{' '}
          · Est. Gross Profit:{' '}
          <strong className="font-mono-tabular text-emerald-700">
            {formatMoney(calc.grossProfit, job.controls.currency)}
          </strong>
        </div>
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={() => onNavigateTab('PRICE_CALC')}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-xs font-medium text-slate-800 hover:bg-slate-50 cursor-pointer min-h-[38px] whitespace-nowrap"
          >
            <Calculator className="w-3.5 h-3.5" />
            <span>Adjust Markup &amp; Freight</span>
          </button>
          <button
            type="button"
            onClick={() => onNavigateTab('QT')}
            className="inline-flex items-center gap-1.5 rounded-lg px-4 py-2 text-xs font-semibold text-white hover:opacity-90 cursor-pointer min-h-[38px] whitespace-nowrap"
            style={{ backgroundColor: theme.primary, color: theme.primaryText }}
          >
            <span>Proceed to Customer Quotation (QT)</span>
            <ArrowRight className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>
    </fieldset>
  );
};
