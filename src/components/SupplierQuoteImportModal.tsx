import React, { useState } from 'react';
import {
  AlertTriangle,
  Check,
  ClipboardPaste,
  FileSpreadsheet,
  FileText,
  Plus,
  Trash2,
  X,
} from 'lucide-react';
import { CURRENCY_OPTIONS, EXCEL_THEMES, INCOTERM_OPTIONS, UOM_OPTIONS } from '../constants/themes';
import { CustomColumnPosition, ThemeVariantId } from '../types/procurement';
import {
  ParsedQuoteItem,
  ParsedSupplierQuote,
  parsePastedSupplierQuote,
} from '../utils/supplierQuoteParser';

export interface QuoteApplyOptions {
  target: 'active_job' | 'new_job';
  overwriteDescriptions: boolean;
  addCodeColumn: boolean;
  codeColumnLabel: string;
  codeColumnPosition: CustomColumnPosition;
}

interface SupplierQuoteImportModalProps {
  initialQuote: ParsedSupplierQuote;
  activeJobName: string;
  themeId: ThemeVariantId;
  onApply: (quote: ParsedSupplierQuote, options: QuoteApplyOptions) => void;
  onClose: () => void;
}

export const SupplierQuoteImportModal: React.FC<SupplierQuoteImportModalProps> = ({
  initialQuote,
  activeJobName,
  themeId,
  onApply,
  onClose,
}) => {
  const theme = EXCEL_THEMES[themeId];
  const [quote, setQuote] = useState<ParsedSupplierQuote>(initialQuote);
  const [pasteMode, setPasteMode] = useState(
    Boolean(initialQuote.isScannedPdf || (initialQuote.sourceType === 'text' && initialQuote.items.length === 0))
  );
  const [pastedText, setPastedText] = useState(initialQuote.rawExtractedText || '');
  const [overwriteDescriptions, setOverwriteDescriptions] = useState(false);
  const [addCodeColumn, setAddCodeColumn] = useState(
    initialQuote.items.some((it) => it.code.trim().length > 0)
  );
  const [codeColumnLabel, setCodeColumnLabel] = useState(initialQuote.detectedCodeLabel || 'Part No.');
  const [codeColumnPosition, setCodeColumnPosition] = useState<CustomColumnPosition>('after_no');

  const selectedCount = quote.items.filter((it) => it.selected).length;

  const handleReparseText = () => {
    const parsed = parsePastedSupplierQuote(pastedText, quote.filename || 'Pasted Supplier Quote');
    setQuote((prev) => ({
      ...parsed,
      sourceType: prev.sourceType,
      supplierName: parsed.supplierName || prev.supplierName,
      supplierRef: parsed.supplierRef || prev.supplierRef,
      currency: parsed.currency || prev.currency,
      incoterm: parsed.incoterm || prev.incoterm,
      paymentTerms: parsed.paymentTerms || prev.paymentTerms,
      deliveryLeadTime: parsed.deliveryLeadTime || prev.deliveryLeadTime,
    }));
    if (parsed.items.some((it) => it.code.trim())) {
      setAddCodeColumn(true);
      setCodeColumnLabel(parsed.detectedCodeLabel || 'Part No.');
    }
    setPasteMode(false);
  };

  const updateItem = (id: string, patch: Partial<ParsedQuoteItem>) => {
    setQuote((prev) => ({
      ...prev,
      items: prev.items.map((it) => (it.id === id ? { ...it, ...patch } : it)),
    }));
  };

  const removeItem = (id: string) => {
    setQuote((prev) => ({
      ...prev,
      items: prev.items.filter((it) => it.id !== id),
    }));
  };

  const addBlankRow = () => {
    setQuote((prev) => ({
      ...prev,
      items: [
        ...prev.items,
        {
          id: `manual-${Date.now()}`,
          selected: true,
          code: '',
          description: '',
          uom: 'PCS',
          qty: 1,
          unitPrice: null,
        },
      ],
    }));
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/55 p-3 sm:p-5 overflow-y-auto">
      <div className="w-full max-w-4xl rounded-xl border border-slate-200 bg-white shadow-2xl overflow-hidden my-auto max-h-[92vh] flex flex-col">
        {/* Header */}
        <div
          className="px-5 py-4 flex items-center justify-between gap-3 shrink-0"
          style={{ backgroundColor: theme.primary, color: theme.primaryText }}
        >
          <div className="flex items-center gap-2.5 min-w-0">
            {quote.sourceType === 'pdf' ? (
              <FileText className="w-5 h-5 shrink-0" />
            ) : (
              <FileSpreadsheet className="w-5 h-5 shrink-0" />
            )}
            <div className="min-w-0">
              <h2 className="text-sm sm:text-base font-bold truncate">
                Supplier Quote Importer — {quote.filename}
              </h2>
              <p className="text-xs opacity-85 truncate">
                Review captured items, prices, quantities, and terms before sending into Price Calc
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 rounded-lg hover:bg-white/15 cursor-pointer shrink-0"
            aria-label="Close supplier quote importer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Scrollable Body */}
        <div className="p-5 overflow-y-auto space-y-5 flex-1">
          {quote.isScannedPdf && (
            <div className="rounded-lg border border-amber-200 bg-amber-50 p-3.5 text-xs text-amber-900 flex items-start gap-2.5">
              <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
              <div>
                <strong>Scanned / Image-Only PDF Detected:</strong> This PDF contains a photo rather than selectable text, so offline text extraction cannot read pixels directly. Open the PDF on your iPad/iPhone or PC, select &amp; copy the quote table (Apple Live Text works great), paste it below, and click <strong>Extract from Pasted Text</strong>.
              </div>
            </div>
          )}

          {/* Toggle Paste Mode */}
          <div className="flex flex-wrap items-center justify-between gap-2">
            <span className="text-xs font-semibold text-slate-700">
              Detected Supplier Details &amp; Terms (Editable)
            </span>
            <button
              type="button"
              onClick={() => setPasteMode((v) => !v)}
              className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer"
            >
              <ClipboardPaste className="w-3.5 h-3.5" />
              <span>{pasteMode ? 'Back to Table Review' : 'Paste / Edit Raw Quote Text'}</span>
            </button>
          </div>

          {pasteMode ? (
            <div className="space-y-3">
              <p className="text-xs text-slate-500">
                Paste copied lines from any supplier PDF, email, Excel sheet, or WhatsApp quote (e.g. Description, Qty, UoM, Unit Price):
              </p>
              <textarea
                rows={9}
                value={pastedText}
                onChange={(e) => setPastedText(e.target.value)}
                placeholder={`Example:\nPN-4402   Hydraulic Seal Kit   PCS   10   42.50   425.00\nPN-8810   Pressure Gauge 0-16 Bar   PCS   5   68.00   340.00\nPayment Terms: 50% Advance, 50% Before Dispatch`}
                className="w-full rounded-lg border border-slate-300 p-3 text-xs font-mono-tabular text-slate-900 focus:border-slate-900 focus:outline-none"
              />
              <button
                type="button"
                onClick={handleReparseText}
                className="inline-flex items-center gap-2 rounded-lg px-4 py-2 text-xs font-semibold text-white cursor-pointer"
                style={{ backgroundColor: theme.primary }}
              >
                <Check className="w-3.5 h-3.5" />
                <span>Extract Items, Prices &amp; Terms</span>
              </button>
            </div>
          ) : (
            <>
              {/* Supplier & Terms Metadata */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 rounded-lg border border-slate-200 bg-slate-50/70 p-4">
                <div>
                  <label className="block text-[11px] font-medium text-slate-600 mb-1">Supplier Name</label>
                  <input
                    type="text"
                    value={quote.supplierName}
                    onChange={(e) => setQuote({ ...quote, supplierName: e.target.value })}
                    className="w-full min-h-[38px] rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-900 focus:border-slate-900 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-medium text-slate-600 mb-1">Supplier Quote Ref.</label>
                  <input
                    type="text"
                    value={quote.supplierRef}
                    onChange={(e) => setQuote({ ...quote, supplierRef: e.target.value })}
                    placeholder="SQ-2026-..."
                    className="w-full min-h-[38px] rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-mono-tabular text-slate-900 focus:border-slate-900 focus:outline-none"
                  />
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="block text-[11px] font-medium text-slate-600 mb-1">Currency</label>
                    <select
                      value={quote.currency}
                      onChange={(e) => setQuote({ ...quote, currency: e.target.value })}
                      className="w-full min-h-[38px] rounded-md border border-slate-200 bg-white px-2 py-1.5 text-xs font-mono-tabular text-slate-900"
                    >
                      {CURRENCY_OPTIONS.map((c) => (
                        <option key={c} value={c}>{c}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-[11px] font-medium text-slate-600 mb-1">Incoterm</label>
                    <select
                      value={quote.incoterm}
                      onChange={(e) => setQuote({ ...quote, incoterm: e.target.value })}
                      className="w-full min-h-[38px] rounded-md border border-slate-200 bg-white px-2 py-1.5 text-xs font-mono-tabular text-slate-900"
                    >
                      <option value="">Keep current</option>
                      {INCOTERM_OPTIONS.map((inc) => (
                        <option key={inc} value={inc}>{inc}</option>
                      ))}
                    </select>
                  </div>
                </div>
                <div className="sm:col-span-2">
                  <label className="block text-[11px] font-medium text-slate-600 mb-1">Detected Payment Terms</label>
                  <input
                    type="text"
                    value={quote.paymentTerms}
                    onChange={(e) => setQuote({ ...quote, paymentTerms: e.target.value })}
                    placeholder="Leave blank to keep existing job payment terms"
                    className="w-full min-h-[38px] rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-900 focus:border-slate-900 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-medium text-slate-600 mb-1">Delivery / Lead Time</label>
                  <input
                    type="text"
                    value={quote.deliveryLeadTime}
                    onChange={(e) => setQuote({ ...quote, deliveryLeadTime: e.target.value })}
                    placeholder="Leave blank to keep current"
                    className="w-full min-h-[38px] rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-900 focus:border-slate-900 focus:outline-none"
                  />
                </div>
              </div>

              {/* Extracted Items Table */}
              <div className="space-y-2">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <span className="text-xs font-semibold text-slate-800">
                    Captured Line Items ({selectedCount} selected · Max 20 rows)
                  </span>
                  <button
                    type="button"
                    onClick={addBlankRow}
                    className="inline-flex items-center gap-1 rounded-md border border-slate-200 px-2.5 py-1 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>Add Row</span>
                  </button>
                </div>

                <div className="overflow-x-auto rounded-lg border border-slate-200">
                  <table className="w-full border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-100 text-slate-700 border-b border-slate-200">
                        <th className="py-2 px-2 w-8 text-center">Use</th>
                        <th className="py-2 px-2 text-left w-28">Part / CAS</th>
                        <th className="py-2 px-2 text-left min-w-[200px]">Supplier Offered Description</th>
                        <th className="py-2 px-2 text-center w-20">UoM</th>
                        <th className="py-2 px-2 text-right w-20">Qty</th>
                        <th className="py-2 px-2 text-right w-28">Unit Cost ({quote.currency})</th>
                        <th className="py-2 px-2 w-8"></th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {quote.items.length === 0 ? (
                        <tr>
                          <td colSpan={7} className="py-8 text-center text-slate-500">
                            No line items auto-detected. Click <strong>Paste / Edit Raw Quote Text</strong> above or <strong>Add Row</strong>.
                          </td>
                        </tr>
                      ) : (
                        quote.items.map((item) => (
                          <tr key={item.id} className={item.selected ? 'bg-white' : 'bg-slate-50 opacity-50'}>
                            <td className="p-1.5 text-center">
                              <input
                                type="checkbox"
                                checked={item.selected}
                                onChange={(e) => updateItem(item.id, { selected: e.target.checked })}
                                className="h-4 w-4 accent-slate-900 cursor-pointer"
                              />
                            </td>
                            <td className="p-1">
                              <input
                                type="text"
                                value={item.code}
                                onChange={(e) => updateItem(item.id, { code: e.target.value })}
                                placeholder="Optional"
                                className="w-full rounded border border-slate-200 px-2 py-1 text-xs font-mono-tabular"
                              />
                            </td>
                            <td className="p-1">
                              <input
                                type="text"
                                value={item.description}
                                onChange={(e) => updateItem(item.id, { description: e.target.value })}
                                className="w-full rounded border border-slate-200 px-2 py-1 text-xs"
                              />
                            </td>
                            <td className="p-1">
                              <select
                                value={item.uom}
                                onChange={(e) => updateItem(item.id, { uom: e.target.value })}
                                className="w-full rounded border border-slate-200 px-1.5 py-1 text-xs font-mono-tabular bg-white"
                              >
                                {UOM_OPTIONS.map((u) => (
                                  <option key={u} value={u}>{u}</option>
                                ))}
                              </select>
                            </td>
                            <td className="p-1">
                              <input
                                type="number"
                                step="any"
                                value={item.qty ?? ''}
                                onChange={(e) =>
                                  updateItem(item.id, {
                                    qty: e.target.value === '' ? null : Number(e.target.value),
                                  })
                                }
                                className="w-full rounded border border-slate-200 px-2 py-1 text-xs font-mono-tabular text-right"
                              />
                            </td>
                            <td className="p-1">
                              <input
                                type="number"
                                step="0.01"
                                value={item.unitPrice ?? ''}
                                onChange={(e) =>
                                  updateItem(item.id, {
                                    unitPrice: e.target.value === '' ? null : Number(e.target.value),
                                  })
                                }
                                className="w-full rounded border border-slate-200 px-2 py-1 text-xs font-mono-tabular text-right font-semibold text-blue-700"
                              />
                            </td>
                            <td className="p-1 text-center">
                              <button
                                type="button"
                                onClick={() => removeItem(item.id)}
                                className="p-1 text-slate-400 hover:text-red-600 cursor-pointer"
                                aria-label="Remove item"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Import Behavior Checkboxes */}
              <div className="flex flex-wrap items-center gap-5 text-xs text-slate-700 pt-1">
                <label className="inline-flex items-center gap-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={addCodeColumn}
                    onChange={(e) => setAddCodeColumn(e.target.checked)}
                    className="h-4 w-4 accent-slate-900"
                  />
                  <span>Add custom column across all docs:</span>
                  <input
                    type="text"
                    value={codeColumnLabel}
                    onChange={(e) => setCodeColumnLabel(e.target.value)}
                    disabled={!addCodeColumn}
                    className="w-28 rounded border border-slate-200 px-2 py-1 text-xs font-medium disabled:opacity-50"
                  />
                  <select
                    value={codeColumnPosition}
                    onChange={(e) => setCodeColumnPosition(e.target.value as CustomColumnPosition)}
                    disabled={!addCodeColumn}
                    aria-label="Custom column placement"
                    className="rounded border border-slate-200 bg-white px-2 py-1 text-xs font-medium text-blue-700 disabled:opacity-50"
                  >
                    <option value="after_no">After No.</option>
                    <option value="after_requested">After Requested Desc</option>
                    <option value="after_offered">After Offered Desc</option>
                    <option value="after_qty">After Qty</option>
                    <option value="end">End of Table</option>
                  </select>
                </label>

                <label className="inline-flex items-center gap-2 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={overwriteDescriptions}
                    onChange={(e) => setOverwriteDescriptions(e.target.checked)}
                    className="h-4 w-4 accent-slate-900"
                  />
                  <span>Also copy into Requested Description (if loading into current job)</span>
                </label>
              </div>
            </>
          )}
        </div>

        {/* Footer Actions */}
        {!pasteMode && (
          <div className="px-5 py-3.5 border-t border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-2 shrink-0">
            <button
              type="button"
              onClick={onClose}
              className="rounded-lg border border-slate-200 bg-white px-4 py-2 text-xs font-medium text-slate-700 hover:bg-slate-100 cursor-pointer min-h-[40px]"
            >
              Cancel
            </button>
            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                disabled={selectedCount === 0}
                onClick={() =>
                  onApply(quote, {
                    target: 'active_job',
                    overwriteDescriptions,
                    addCodeColumn,
                    codeColumnLabel: codeColumnLabel.trim() || 'Part No.',
                    codeColumnPosition,
                  })
                }
                className="rounded-lg border border-slate-300 bg-white px-4 py-2 text-xs font-semibold text-slate-900 hover:bg-slate-100 cursor-pointer min-h-[40px] disabled:opacity-50"
              >
                Load into &ldquo;{activeJobName.slice(0, 24)}&rdquo; Costing
              </button>
              <button
                type="button"
                disabled={selectedCount === 0}
                onClick={() =>
                  onApply(quote, {
                    target: 'new_job',
                    overwriteDescriptions: true,
                    addCodeColumn,
                    codeColumnLabel: codeColumnLabel.trim() || 'Part No.',
                    codeColumnPosition,
                  })
                }
                className="rounded-lg px-4 py-2 text-xs font-semibold text-white hover:opacity-90 cursor-pointer min-h-[40px] disabled:opacity-50"
                style={{ backgroundColor: theme.primary }}
              >
                Create New Job from Supplier Quote
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
