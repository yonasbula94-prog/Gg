import React, { useMemo, useState } from 'react';
import {
  AlertTriangle,
  Check,
  ClipboardPaste,
  Columns,
  FileSpreadsheet,
  FileText,
  Plus,
  Sparkles,
  Table as TableIcon,
  Trash2,
  X,
} from 'lucide-react';
import { CURRENCY_OPTIONS, INCOTERM_OPTIONS, UOM_OPTIONS } from '../constants/themes';
import { CustomColumnPosition, ThemePalette } from '../types/procurement';
import {
  ColumnRole,
  extractItemsFromMappedTable,
  guessColumnRolesForHeaderRow,
  parseDelimitedTextToTable,
  ParsedQuoteItem,
  ParsedSupplierQuote,
  parsePastedSupplierQuote,
  RawExtractedTable,
} from '../utils/supplierQuoteParser';
import { HelpTip } from './HelpTip';

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
  theme: ThemePalette;
  onApply: (quote: ParsedSupplierQuote, options: QuoteApplyOptions) => void;
  onClose: () => void;
}

const ROLE_OPTIONS: { value: ColumnRole; label: string }[] = [
  { value: 'ignore', label: '— Ignore —' },
  { value: 'description', label: 'Description' },
  { value: 'code', label: 'Part / CAS No.' },
  { value: 'uom', label: 'UoM' },
  { value: 'qty', label: 'Qty' },
  { value: 'unitPrice', label: 'Unit Price' },
  { value: 'lineTotal', label: 'Line Total (÷ Qty)' },
];

type ModalMode = 'review' | 'mapper' | 'paste';

export const SupplierQuoteImportModal: React.FC<SupplierQuoteImportModalProps> = ({
  initialQuote,
  activeJobName,
  theme,
  onApply,
  onClose,
}) => {
  const [quote, setQuote] = useState<ParsedSupplierQuote>(initialQuote);
  const hasTables = Boolean(initialQuote.rawTables && initialQuote.rawTables.length > 0);

  // If auto-capture found 0 items on an Excel/text-PDF file that has raw tables, open Column Mapper directly!
  const initialMode: ModalMode = initialQuote.isScannedPdf
    ? 'paste'
    : initialQuote.sourceType === 'text' && initialQuote.items.length === 0
    ? 'paste'
    : initialQuote.items.length === 0 && hasTables
    ? 'mapper'
    : 'review';

  const [mode, setMode] = useState<ModalMode>(initialMode);
  const [pastedText, setPastedText] = useState(initialQuote.rawExtractedText || '');
  const [pasteDelimiter, setPasteDelimiter] = useState<'auto' | 'tab' | 'spaces' | 'pipe' | 'comma'>('auto');

  // Interactive Column Mapper state
  const [tables, setTables] = useState<RawExtractedTable[]>(initialQuote.rawTables || []);
  const [activeTableIdx, setActiveTableIdx] = useState<number>(
    initialQuote.detectedMapping?.tableIndex ?? 0
  );
  const [headerRowIdx, setHeaderRowIdx] = useState<number>(
    initialQuote.detectedMapping?.headerRowIndex ?? 0
  );
  const [columnRoles, setColumnRoles] = useState<ColumnRole[]>(() => {
    if (initialQuote.detectedMapping?.columnRoles?.length) {
      return initialQuote.detectedMapping.columnRoles;
    }
    const firstRows = initialQuote.rawTables?.[0]?.rows;
    if (firstRows && firstRows.length > 0) {
      return guessColumnRolesForHeaderRow(firstRows[0]).roles;
    }
    return [];
  });
  const [mergeWrappedLines, setMergeWrappedLines] = useState(true);

  const [overwriteDescriptions, setOverwriteDescriptions] = useState(false);
  const [addCodeColumn, setAddCodeColumn] = useState(
    initialQuote.items.some((it) => it.code.trim().length > 0)
  );
  const [codeColumnLabel, setCodeColumnLabel] = useState(initialQuote.detectedCodeLabel || 'Part No.');
  const [codeColumnPosition, setCodeColumnPosition] = useState<CustomColumnPosition>('after_no');

  const activeTable = tables[activeTableIdx] || null;
  const activeColCount = activeTable?.rows?.[0]?.length || 0;

  const normalizedRoles = useMemo<ColumnRole[]>(() => {
    if (!activeTable) return [];
    const out = columnRoles.slice(0, activeColCount);
    while (out.length < activeColCount) out.push('ignore');
    return out;
  }, [activeTable, activeColCount, columnRoles]);

  const mappedPreview = useMemo(() => {
    if (!activeTable) return { items: [], detectedCodeLabel: 'Part No.' };
    return extractItemsFromMappedTable(
      activeTable.rows,
      headerRowIdx,
      normalizedRoles,
      mergeWrappedLines
    );
  }, [activeTable, headerRowIdx, normalizedRoles, mergeWrappedLines]);

  const selectedCount = quote.items.filter((it) => it.selected).length;

  const handleSelectSheet = (idx: number) => {
    setActiveTableIdx(idx);
    const tbl = tables[idx];
    if (!tbl || tbl.rows.length === 0) {
      setHeaderRowIdx(0);
      setColumnRoles([]);
      return;
    }
    // Auto-guess best header row for the newly chosen sheet
    let bestR = 0;
    let bestScore = -1;
    let bestRoles: ColumnRole[] = tbl.rows[0].map(() => 'ignore');
    for (let r = 0; r < Math.min(tbl.rows.length, 30); r++) {
      const g = guessColumnRolesForHeaderRow(tbl.rows[r]);
      if (g.score > bestScore) {
        bestScore = g.score;
        bestR = r;
        bestRoles = g.roles;
      }
    }
    setHeaderRowIdx(bestR);
    setColumnRoles(bestRoles);
  };

  const handleSetHeaderRow = (rIdx: number) => {
    setHeaderRowIdx(rIdx);
    if (activeTable && activeTable.rows[rIdx]) {
      const guess = guessColumnRolesForHeaderRow(activeTable.rows[rIdx]);
      if (guess.score >= 3) {
        setColumnRoles(guess.roles);
      }
    }
  };

  const handleSetColumnRole = (colIdx: number, role: ColumnRole) => {
    setColumnRoles(() => {
      const next = [...normalizedRoles];
      // Ensure single assignment for non-ignore roles so columns don't accidentally collide
      if (role !== 'ignore') {
        for (let i = 0; i < next.length; i++) {
          if (next[i] === role) next[i] = 'ignore';
        }
      }
      next[colIdx] = role;
      return next;
    });
  };

  const handleApplyColumnMapping = () => {
    if (mappedPreview.items.length === 0) return;
    setQuote((prev) => ({
      ...prev,
      items: mappedPreview.items,
      detectedCodeLabel: mappedPreview.detectedCodeLabel,
    }));
    if (mappedPreview.items.some((it) => it.code.trim())) {
      setAddCodeColumn(true);
      setCodeColumnLabel(mappedPreview.detectedCodeLabel || 'Part No.');
    }
    setMode('review');
  };

  const handleReparseText = (openMapperAfter: boolean) => {
    const parsed = parsePastedSupplierQuote(pastedText, quote.filename || 'Pasted Supplier Quote');
    const splitGrid = parseDelimitedTextToTable(pastedText, pasteDelimiter);
    const nextTables: RawExtractedTable[] =
      splitGrid.length > 0 ? [{ name: 'Pasted Text Table', rows: splitGrid }] : [];

    setTables(nextTables);
    setActiveTableIdx(0);

    if (splitGrid.length > 0) {
      let bestR = 0;
      let bestScore = -1;
      let bestRoles: ColumnRole[] = splitGrid[0].map(() => 'ignore');
      for (let r = 0; r < Math.min(splitGrid.length, 20); r++) {
        const g = guessColumnRolesForHeaderRow(splitGrid[r]);
        if (g.score > bestScore) {
          bestScore = g.score;
          bestR = r;
          bestRoles = g.roles;
        }
      }
      setHeaderRowIdx(bestScore >= 3 ? bestR : -1);
      setColumnRoles(bestRoles);
    }

    setQuote((prev) => ({
      ...parsed,
      sourceType: prev.sourceType,
      supplierName: parsed.supplierName || prev.supplierName,
      supplierRef: parsed.supplierRef || prev.supplierRef,
      currency: parsed.currency || prev.currency,
      incoterm: parsed.incoterm || prev.incoterm,
      paymentTerms: parsed.paymentTerms || prev.paymentTerms,
      deliveryLeadTime: parsed.deliveryLeadTime || prev.deliveryLeadTime,
      rawTables: nextTables,
    }));

    if (parsed.items.some((it) => it.code.trim())) {
      setAddCodeColumn(true);
      setCodeColumnLabel(parsed.detectedCodeLabel || 'Part No.');
    }

    setMode(openMapperAfter && nextTables.length > 0 ? 'mapper' : 'review');
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

  const colLetter = (idx: number) => {
    let n = idx + 1;
    let s = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/55 p-3 sm:p-5 overflow-y-auto">
      <div className="w-full max-w-5xl rounded-xl border border-slate-200 bg-white shadow-2xl overflow-hidden my-auto max-h-[92vh] flex flex-col">
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

        {/* Mode Switcher Bar */}
        <div className="px-5 py-2.5 border-b border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-2 shrink-0">
          <div className="inline-flex items-center p-1 rounded-lg bg-slate-200/80">
            <button
              type="button"
              onClick={() => setMode('review')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold transition-colors cursor-pointer ${
                mode === 'review' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <TableIcon className="w-3.5 h-3.5" />
              <span>Review Items ({selectedCount})</span>
            </button>
            <button
              type="button"
              disabled={tables.length === 0}
              onClick={() => setMode('mapper')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold transition-colors cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed ${
                mode === 'mapper' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Columns className="w-3.5 h-3.5" />
              <span>Column Mapper Assistant</span>
            </button>
            <button
              type="button"
              onClick={() => setMode('paste')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold transition-colors cursor-pointer ${
                mode === 'paste' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <ClipboardPaste className="w-3.5 h-3.5" />
              <span>Paste / Split Text</span>
            </button>
          </div>

          <div className="flex items-center gap-1.5 text-xs text-slate-500">
            <span>Auto-capture missed columns? Use Column Mapper</span>
            <HelpTip
              align="right"
              text="If a supplier uses non-standard headings (e.g. Artikel, Menge, Net Rate) or a complex PDF table, open Column Mapper Assistant to click the header row and assign each column manually in seconds."
            />
          </div>
        </div>

        {/* Scrollable Body */}
        <div className="p-5 overflow-y-auto space-y-5 flex-1">
          {quote.isScannedPdf && (
            <div className="rounded-lg border border-amber-200 bg-amber-50 p-3.5 text-xs text-amber-900 flex items-start gap-2.5">
              <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
              <div>
                <strong>Scanned / Image-Only PDF Detected:</strong> Copy the table text from your PDF viewer (Apple Live Text works on iPad/iPhone), paste below, and click <strong>Send to Column Mapper</strong>.
              </div>
            </div>
          )}

          {/* ==================== MODE 1: INTERACTIVE COLUMN MAPPER ASSISTANT ==================== */}
          {mode === 'mapper' && activeTable && (
            <div className="space-y-4">
              <div className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-slate-200 bg-slate-50/70 p-3.5">
                <div className="flex flex-wrap items-center gap-3">
                  {tables.length > 1 && (
                    <div className="flex items-center gap-1.5">
                      <label className="text-xs font-semibold text-slate-700">Sheet:</label>
                      <select
                        value={activeTableIdx}
                        onChange={(e) => handleSelectSheet(Number(e.target.value))}
                        className="rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-900"
                      >
                        {tables.map((t, i) => (
                          <option key={t.name + i} value={i}>
                            {t.name} ({t.rows.length} rows)
                          </option>
                        ))}
                      </select>
                    </div>
                  )}

                  <div className="flex items-center gap-1.5 text-xs text-slate-700">
                    <span className="font-semibold">Header Row:</span>
                    <span className="font-mono-tabular rounded bg-white border border-slate-200 px-2 py-1 font-semibold text-blue-700">
                      {headerRowIdx >= 0 ? `Row #${headerRowIdx + 1}` : 'None (Start at Row 1)'}
                    </span>
                    <button
                      type="button"
                      onClick={() => setHeaderRowIdx(-1)}
                      className="text-[11px] text-slate-500 hover:text-slate-900 underline cursor-pointer"
                    >
                      No Header
                    </button>
                    <HelpTip text="Click any row in the table below to mark it as the Header Row — data capture starts on the very next row. Choose each column's role from the dropdowns at the top." />
                  </div>

                  <label className="inline-flex items-center gap-1.5 text-xs text-slate-700 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={mergeWrappedLines}
                      onChange={(e) => setMergeWrappedLines(e.target.checked)}
                      className="h-3.5 w-3.5 accent-slate-900"
                    />
                    <span>Merge wrapped PDF lines</span>
                  </label>
                </div>

                <div className="flex items-center gap-2.5">
                  <span className="text-xs font-mono-tabular font-semibold text-emerald-700">
                    {mappedPreview.items.length} items matched
                  </span>
                  <button
                    type="button"
                    disabled={mappedPreview.items.length === 0}
                    onClick={handleApplyColumnMapping}
                    className="inline-flex items-center gap-1.5 rounded-lg px-3.5 py-2 text-xs font-semibold text-white cursor-pointer disabled:opacity-40"
                    style={{ backgroundColor: theme.primary }}
                  >
                    <Sparkles className="w-3.5 h-3.5" />
                    <span>Apply Mapping ({mappedPreview.items.length})</span>
                  </button>
                </div>
              </div>

              {/* Raw Table Preview with Column Role Dropdowns */}
              <div className="overflow-x-auto rounded-lg border border-slate-200 max-h-[380px]">
                <table className="w-full border-collapse text-xs">
                  <thead className="sticky top-0 z-10 bg-slate-100 border-b border-slate-300">
                    <tr>
                      <th className="py-2 px-2 text-center font-semibold text-slate-500 w-16 border-r border-slate-200">
                        Row
                      </th>
                      {Array.from({ length: activeColCount }, (_, cIdx) => {
                        const role = normalizedRoles[cIdx] || 'ignore';
                        return (
                          <th
                            key={cIdx}
                            className="p-1.5 text-left min-w-[130px] border-r border-slate-200 last:border-r-0"
                          >
                            <div className="text-[10px] font-mono-tabular text-slate-400 mb-1">
                              Col {colLetter(cIdx)}
                            </div>
                            <select
                              aria-label={`Map Column ${colLetter(cIdx)}`}
                              value={role}
                              onChange={(e) =>
                                handleSetColumnRole(cIdx, e.target.value as ColumnRole)
                              }
                              className={`w-full rounded border px-2 py-1 text-xs font-semibold cursor-pointer focus:outline-none ${
                                role === 'ignore'
                                  ? 'border-slate-200 bg-white text-slate-500'
                                  : 'border-blue-600 bg-blue-50 text-blue-800'
                              }`}
                            >
                              {ROLE_OPTIONS.map((opt) => (
                                <option key={opt.value} value={opt.value}>
                                  {opt.label}
                                </option>
                              ))}
                            </select>
                          </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {activeTable.rows.slice(0, 60).map((row, rIdx) => {
                      const isHeader = rIdx === headerRowIdx;
                      const isSkippedAbove = headerRowIdx >= 0 && rIdx < headerRowIdx;
                      return (
                        <tr
                          key={rIdx}
                          onClick={() => handleSetHeaderRow(rIdx)}
                          title="Click to mark this row as the table header"
                          className={`cursor-pointer transition-colors ${
                            isHeader
                              ? 'bg-amber-100/90 font-semibold text-slate-900'
                              : isSkippedAbove
                              ? 'bg-slate-50 text-slate-400'
                              : 'hover:bg-blue-50/40 text-slate-800'
                          }`}
                        >
                          <td className="py-1.5 px-2 text-center font-mono-tabular text-[11px] border-r border-slate-100 whitespace-nowrap">
                            {isHeader ? (
                              <span className="text-amber-900 font-bold">HDR #{rIdx + 1}</span>
                            ) : (
                              <span>#{rIdx + 1}</span>
                            )}
                          </td>
                          {Array.from({ length: activeColCount }, (_, cIdx) => (
                            <td
                              key={cIdx}
                              className={`py-1.5 px-2 border-r border-slate-100 last:border-r-0 max-w-[240px] truncate ${
                                normalizedRoles[cIdx] !== 'ignore' && !isSkippedAbove && !isHeader
                                  ? 'bg-blue-50/25'
                                  : ''
                              }`}
                            >
                              {row[cIdx] || ''}
                            </td>
                          ))}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* ==================== MODE 2: PASTE / SPLIT RAW TEXT ==================== */}
          {mode === 'paste' && (
            <div className="space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-700">
                  <span>Paste Quote Lines</span>
                  <HelpTip text="Paste copied lines from any PDF, email, WhatsApp, or Excel sheet. Choose how columns are separated, then either auto-extract or open the Column Mapper Assistant." />
                </div>
                <div className="flex items-center gap-1.5 text-xs">
                  <span className="text-slate-500">Split by:</span>
                  <select
                    value={pasteDelimiter}
                    onChange={(e) =>
                      setPasteDelimiter(e.target.value as 'auto' | 'tab' | 'spaces' | 'pipe' | 'comma')
                    }
                    className="rounded-md border border-slate-200 bg-white px-2 py-1 text-xs font-medium text-slate-800"
                  >
                    <option value="auto">Auto-Detect</option>
                    <option value="tab">Tab (Excel / PDF Table)</option>
                    <option value="spaces">2+ Spaces</option>
                    <option value="pipe">Pipe ( | )</option>
                    <option value="comma">Comma (CSV)</option>
                  </select>
                </div>
              </div>
              <textarea
                rows={9}
                value={pastedText}
                onChange={(e) => setPastedText(e.target.value)}
                placeholder={`Example:\nPart No.\tDescription\tUoM\tQty\tUnit Price\nPN-4402\tHydraulic Seal Kit\tPCS\t10\t42.50\nPN-8810\tPressure Gauge 0-16 Bar\tPCS\t5\t68.00\nPayment Terms: 50% Advance, 50% Before Dispatch`}
                className="w-full rounded-lg border border-slate-300 p-3 text-xs font-mono-tabular text-slate-900 focus:border-slate-900 focus:outline-none"
              />
              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={() => handleReparseText(false)}
                  className="inline-flex items-center gap-2 rounded-lg px-4 py-2 text-xs font-semibold text-white cursor-pointer"
                  style={{ backgroundColor: theme.primary }}
                >
                  <Check className="w-3.5 h-3.5" />
                  <span>Auto-Extract Items &amp; Terms</span>
                </button>
                <button
                  type="button"
                  onClick={() => handleReparseText(true)}
                  className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-4 py-2 text-xs font-semibold text-slate-800 hover:bg-slate-50 cursor-pointer"
                >
                  <Columns className="w-3.5 h-3.5" />
                  <span>Split &amp; Open Column Mapper</span>
                </button>
              </div>
            </div>
          )}

          {/* ==================== MODE 3: REVIEW CAPTURED ITEMS ==================== */}
          {mode === 'review' && (
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
                  <label className="block text-[11px] font-medium text-slate-600 mb-1">Payment Terms</label>
                  <input
                    type="text"
                    value={quote.paymentTerms}
                    onChange={(e) => setQuote({ ...quote, paymentTerms: e.target.value })}
                    placeholder="Leave blank to keep existing"
                    className="w-full min-h-[38px] rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-900 focus:border-slate-900 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-medium text-slate-600 mb-1">Delivery Lead Time</label>
                  <input
                    type="text"
                    value={quote.deliveryLeadTime}
                    onChange={(e) => setQuote({ ...quote, deliveryLeadTime: e.target.value })}
                    placeholder="Leave blank to keep existing"
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
                  <div className="flex items-center gap-2">
                    {tables.length > 0 && (
                      <button
                        type="button"
                        onClick={() => setMode('mapper')}
                        className="inline-flex items-center gap-1 rounded-md border border-blue-200 bg-blue-50 px-2.5 py-1 text-xs font-medium text-blue-800 hover:bg-blue-100 cursor-pointer"
                      >
                        <Columns className="w-3.5 h-3.5" />
                        <span>Remap Columns</span>
                      </button>
                    )}
                    <button
                      type="button"
                      onClick={addBlankRow}
                      className="inline-flex items-center gap-1 rounded-md border border-slate-200 px-2.5 py-1 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      <span>Add Row</span>
                    </button>
                  </div>
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
                            No line items auto-detected. Open <strong>Column Mapper Assistant</strong> above to select columns manually, or click <strong>Add Row</strong>.
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
                  <span>Add custom column:</span>
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
                  <span>Also fill Requested Description</span>
                </label>
              </div>
            </>
          )}
        </div>

        {/* Footer Actions */}
        {mode === 'review' && (
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
