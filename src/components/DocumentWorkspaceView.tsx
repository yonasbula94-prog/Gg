import React, { useState } from 'react';
import {
  AlertCircle,
  CheckCircle2,
  FileSpreadsheet,
  Lock,
  Plus,
  Printer,
  Copy,
  Eye,
  Table as TableIcon,
  X,
} from 'lucide-react';
import {
  CURRENCY_OPTIONS,
  DOC_META,
  EXCEL_THEMES,
  INCOTERM_OPTIONS,
  UOM_OPTIONS,
} from '../constants/themes';
import {
  CompanyProfile,
  CustomColumnDef,
  CustomColumnPosition,
  DocPrefix,
  JobCalculationSummary,
  JobControls,
  JobItemRow,
  PlatformMode,
  ProcurementJob,
  ThemeVariantId,
  ValidationIssue,
} from '../types/procurement';
import {
  formatMoney,
  formatNumber,
  formatPct,
  getResolvedEndUser,
  validateForIssue,
} from '../utils/calculations';
import { exportIssuedDocumentXlsx } from '../utils/excelEngine';
import { IPadJobEditor } from './IPadJobEditor';

interface DocumentWorkspaceViewProps {
  job: ProcurementJob;
  prefix: DocPrefix;
  calc: JobCalculationSummary;
  themeId: ThemeVariantId;
  company: CompanyProfile;
  platformMode: PlatformMode;
  nextSeq: number;
  onUpdateJobName: (name: string) => void;
  onUpdateControls: (patch: Partial<JobControls>) => void;
  onUpdateCustomColumns: (cols: CustomColumnDef[]) => void;
  onUpdateItem: (rowNumber: number, patch: Partial<JobItemRow>) => void;
  onIssueDocument: (prefix: DocPrefix) => void;
  onDuplicateAsNewJob: () => void;
  onNotify: (msg: string) => void;
}

const PRESET_CUSTOM_COLUMNS = [
  'Part No.',
  'CAS No.',
  'Brand / Mfr',
  'HS Code',
  'Batch / Expiry',
];

const CUSTOM_COLUMN_POSITIONS: { value: CustomColumnPosition; label: string }[] = [
  { value: 'after_no', label: 'After No.' },
  { value: 'after_requested', label: 'After Requested Desc' },
  { value: 'after_offered', label: 'After Offered Desc' },
  { value: 'after_qty', label: 'After Qty' },
  { value: 'end', label: 'End of Table' },
];

function parseOptionalNumber(val: string): number | null {
  if (val.trim() === '') return null;
  const n = Number(val);
  return Number.isNaN(n) ? null : n;
}

export const DocumentWorkspaceView: React.FC<DocumentWorkspaceViewProps> = ({
  job,
  prefix,
  calc,
  themeId,
  company,
  platformMode,
  nextSeq,
  onUpdateJobName,
  onUpdateControls,
  onUpdateCustomColumns,
  onUpdateItem,
  onIssueDocument,
  onDuplicateAsNewJob,
  onNotify,
}) => {
  const [subView, setSubView] = useState<'edit' | 'preview'>(prefix === 'PI' ? 'edit' : 'preview');
  const [validationErrors, setValidationErrors] = useState<ValidationIssue[]>([]);
  const [exporting, setExporting] = useState(false);
  const [newColLabel, setNewColLabel] = useState('');
  const [newColPosition, setNewColPosition] = useState<CustomColumnPosition>('after_no');

  const customColumns = job.customColumns || [];
  const colsAt = (pos: CustomColumnPosition) =>
    customColumns.filter((c) => (c.position || 'after_no') === pos);

  const handleAddCustomColumn = (labelRaw: string, position: CustomColumnPosition = newColPosition) => {
    const label = labelRaw.trim();
    if (!label || isLocked) return;
    if (customColumns.some((c) => c.label.toLowerCase() === label.toLowerCase())) {
      onNotify(`Column "${label}" is already active on this job.`);
      return;
    }
    if (customColumns.length >= 4) {
      onNotify('Up to 4 extra columns can be added per job to keep A4 printouts clean.');
      return;
    }
    const id = `col_${label.toLowerCase().replace(/[^a-z0-9]+/g, '_')}_${Date.now().toString(36).slice(-3)}`;
    onUpdateCustomColumns([...customColumns, { id, label, position }]);
    setNewColLabel('');
    const posLabel = CUSTOM_COLUMN_POSITIONS.find((p) => p.value === position)?.label || 'After No.';
    onNotify(`Added "${label}" (${posLabel}) across all 6 documents and Excel exports.`);
  };

  const handleMoveCustomColumn = (id: string, position: CustomColumnPosition) => {
    if (isLocked) return;
    onUpdateCustomColumns(
      customColumns.map((c) => (c.id === id ? { ...c, position } : c))
    );
  };

  const handleRemoveCustomColumn = (id: string) => {
    if (isLocked) return;
    onUpdateCustomColumns(customColumns.filter((c) => c.id !== id));
  };

  const theme = EXCEL_THEMES[themeId];
  const meta = DOC_META[prefix];
  const issued = job.issuedDocs[prefix];
  const isLocked = Boolean(issued && !issued.isVoided);
  const isLogistics = prefix === 'PL' || prefix === 'DN';
  const isPO = prefix === 'PO';
  const endUser = getResolvedEndUser(job);
  const curr = job.controls.currency;
  const activeRows = calc.rows.filter((r) => !r.isEmpty);

  const handleTryIssue = () => {
    const errors = validateForIssue(job, prefix, calc);
    if (errors.length > 0) {
      setValidationErrors(errors);
      return;
    }
    setValidationErrors([]);
    onIssueDocument(prefix);
  };

  const handleExportXlsx = async () => {
    setExporting(true);
    try {
      const filename = await exportIssuedDocumentXlsx(job, prefix, themeId, company);
      onNotify(`Exported values-only Excel file: ${filename}`);
    } catch {
      onNotify('Failed to generate Excel file. Please try again.');
    } finally {
      setExporting(false);
    }
  };

  const handlePrint = () => {
    if (subView !== 'preview') {
      setSubView('preview');
      setTimeout(() => window.print(), 120);
    } else {
      window.print();
    }
  };

  const nextPreviewDocNo = `${prefix}-GG-${String(nextSeq).padStart(3, '0')}-${new Date().getFullYear()}`;
  const displayDocNo = issued ? issued.docNumber : `DRAFT (Next: ${nextPreviewDocNo})`;
  const displayDate = issued
    ? issued.issuedAt.slice(0, 10)
    : new Date().toISOString().slice(0, 10);

  return (
    <div className="space-y-5">
      {/* Document Action Strip (hidden during print) */}
      <div className="no-print flex flex-wrap items-center justify-between gap-3 rounded-lg border border-slate-200 bg-white px-5 py-3.5">
        <div className="flex flex-wrap items-center gap-3">
          {/* Segmented Edit vs A4 Print Preview */}
          <div className="inline-flex items-center p-1 rounded-lg bg-slate-100">
            <button
              type="button"
              onClick={() => setSubView('edit')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap min-h-[34px] ${
                subView === 'edit'
                  ? 'bg-white text-slate-900 shadow-sm'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <TableIcon className="w-3.5 h-3.5" />
              <span>{prefix === 'PI' ? 'Job Entry & Controls' : 'Document Inputs'}</span>
            </button>
            <button
              type="button"
              onClick={() => setSubView('preview')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap min-h-[34px] ${
                subView === 'preview'
                  ? 'bg-white text-slate-900 shadow-sm'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Eye className="w-3.5 h-3.5" />
              <span>A4 Print &amp; Excel Preview</span>
            </button>
          </div>

          {/* Clean unboxed status metadata */}
          <div className="text-xs text-slate-600 flex items-center gap-2 font-mono-tabular">
            {isLocked ? (
              <span className="inline-flex items-center gap-1.5 font-semibold text-emerald-700">
                <CheckCircle2 className="w-3.5 h-3.5" />
                Issued · {issued?.docNumber}
              </span>
            ) : (
              <span className="text-slate-500">
                Draft · Next Registry No. {nextPreviewDocNo}
              </span>
            )}
          </div>
        </div>

        {/* Right Actions: Issue, Export Values-Only .xlsx, Print PDF */}
        <div className="flex flex-wrap items-center gap-2">
          {isLocked ? (
            <button
              type="button"
              onClick={onDuplicateAsNewJob}
              className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-xs font-medium text-slate-700 hover:bg-slate-50 transition-colors cursor-pointer min-h-[40px] whitespace-nowrap"
            >
              <Copy className="w-3.5 h-3.5" />
              <span>New Job to Re-Issue</span>
            </button>
          ) : (
            <button
              type="button"
              onClick={handleTryIssue}
              className="inline-flex items-center gap-1.5 rounded-lg px-4 py-2 text-xs font-semibold text-white transition-opacity hover:opacity-90 cursor-pointer min-h-[40px] whitespace-nowrap"
              style={{ backgroundColor: theme.primary }}
            >
              <Lock className="w-3.5 h-3.5" />
              <span>Issue {meta.shortName}</span>
            </button>
          )}

          <button
            type="button"
            disabled={exporting}
            onClick={handleExportXlsx}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-xs font-medium text-slate-800 hover:bg-slate-50 transition-colors cursor-pointer min-h-[40px] whitespace-nowrap disabled:opacity-50"
          >
            <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-700" />
            <span>{exporting ? 'Exporting...' : 'Values-Only .xlsx'}</span>
          </button>

          <button
            type="button"
            onClick={handlePrint}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-xs font-medium text-slate-800 hover:bg-slate-50 transition-colors cursor-pointer min-h-[40px] whitespace-nowrap"
          >
            <Printer className="w-3.5 h-3.5 text-slate-600" />
            <span>Print / PDF</span>
          </button>
        </div>
      </div>

      {/* Validation Alert Panel */}
      {validationErrors.length > 0 && (
        <div className="no-print rounded-lg border border-red-200 bg-red-50 p-4">
          <div className="flex items-start gap-2.5">
            <AlertCircle className="w-4 h-4 text-red-600 mt-0.5 shrink-0" />
            <div className="space-y-1">
              <p className="text-sm font-semibold text-red-900">
                Pre-Issue Validation Blocked — No Registry Number Was Consumed
              </p>
              <ul className="text-xs text-red-800 list-disc pl-4 space-y-0.5">
                {validationErrors.map((err, i) => (
                  <li key={i}>{err.message}</li>
                ))}
              </ul>
            </div>
          </div>
        </div>
      )}

      {/* Locked Notice when Issued */}
      {isLocked && subView === 'edit' && (
        <div className="no-print rounded-lg border border-emerald-200 bg-emerald-50/70 px-4 py-3 text-xs text-emerald-900 flex items-center justify-between gap-3">
          <span>
            <strong>{issued?.docNumber}</strong> is frozen and logged in the Master Registry. Each document tab can be issued once per job.
          </span>
          <button
            type="button"
            onClick={() => setSubView('preview')}
            className="font-semibold underline hover:text-emerald-700 cursor-pointer whitespace-nowrap"
          >
            View Issued A4 Document
          </button>
        </div>
      )}

      {/* ==================== EDIT / CONTROLS SUBVIEW ==================== */}
      {subView === 'edit' && (
        <div className="no-print space-y-6">
          {/* Job Header & Controls Block (K:L equivalent + Stage-specific fields) */}
          <div className="rounded-lg border border-slate-200 bg-white p-5 space-y-5">
            <div className="flex flex-wrap items-center justify-between gap-2 pb-3 border-b border-slate-100">
              <div>
                <h2 className="text-sm font-semibold text-slate-900">
                  {prefix === 'PI'
                    ? 'Master Job Controls & End-User Selection (PI Cell C8 & Block K:L)'
                    : `${meta.shortName} Specific Fields`}
                </h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Changes propagate automatically across all linked documents.
                </p>
              </div>
              <span className="text-xs text-slate-500 font-mono-tabular">
                {meta.stageLabel} · {meta.sheetName}
              </span>
            </div>

            {prefix === 'PO' ? (
              /* Supplier PO specific header controls */
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Supplier Name (Required for PO)
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.supplierName}
                    onChange={(e) => onUpdateControls({ supplierName: e.target.value })}
                    placeholder="e.g. Al-Futtaim Industrial Supplies"
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Supplier Quote Reference
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.supplierRef}
                    onChange={(e) => onUpdateControls({ supplierRef: e.target.value })}
                    placeholder="SQ-2026-..."
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Supplier Address
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.supplierAddress}
                    onChange={(e) => onUpdateControls({ supplierAddress: e.target.value })}
                    placeholder="Dubai, UAE"
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
              </div>
            ) : prefix === 'PL' ? (
              /* Packing List specific controls */
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Ship Date (Required for PL)
                  </label>
                  <input
                    type="date"
                    disabled={isLocked}
                    value={job.controls.shipDate}
                    onChange={(e) => onUpdateControls({ shipDate: e.target.value })}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Port of Loading
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.portOfLoading}
                    onChange={(e) => onUpdateControls({ portOfLoading: e.target.value })}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Destination
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.destination}
                    onChange={(e) => onUpdateControls({ destination: e.target.value })}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
              </div>
            ) : prefix === 'DN' ? (
              /* Delivery Note specific controls */
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Delivery Date (Required for DN)
                  </label>
                  <input
                    type="date"
                    disabled={isLocked}
                    value={job.controls.deliveryDate}
                    onChange={(e) => onUpdateControls({ deliveryDate: e.target.value })}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Transport / BL / Truck Ref. (Required)
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.transportRef}
                    onChange={(e) => onUpdateControls({ transportRef: e.target.value })}
                    placeholder="BL-MSC-... / Truck Plate"
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Received By (Consignee Contact)
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.receivedBy}
                    onChange={(e) => onUpdateControls({ receivedBy: e.target.value })}
                    placeholder="Warehouse Manager / Receiver"
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
              </div>
            ) : (
              /* Primary PI / QT / CI Job Controls */
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="lg:col-span-2">
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Job Title / Internal Reference
                  </label>
                  <input
                    type="text"
                    value={job.jobName}
                    onChange={(e) => onUpdateJobName(e.target.value)}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-900 focus:border-slate-900 focus:outline-none"
                  />
                </div>

                <div className="lg:col-span-2">
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    TO — End User (PI Cell C8 · {company.endUsers.length} End Users)
                  </label>
                  <select
                    disabled={isLocked}
                    value={job.controls.endUserSelect}
                    onChange={(e) => onUpdateControls({ endUserSelect: e.target.value })}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm font-medium text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  >
                    {company.endUsers.map((opt) => (
                      <option key={opt} value={opt}>
                        {opt}
                      </option>
                    ))}
                  </select>
                </div>

                {job.controls.endUserSelect === 'OTHER' && (
                  <div className="lg:col-span-2">
                    <label className="block text-xs font-medium text-slate-600 mb-1.5">
                      Custom End-User Name (OTHER)
                    </label>
                    <input
                      type="text"
                      disabled={isLocked}
                      value={job.controls.customEndUser}
                      onChange={(e) => onUpdateControls({ customEndUser: e.target.value })}
                      placeholder="Type end-user organization name..."
                      className="w-full min-h-[42px] rounded-lg border border-slate-900 px-3 py-2 text-sm text-blue-700 focus:outline-none"
                    />
                  </div>
                )}

                <div className="lg:col-span-2">
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    End-User Contact / ATTN
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.endUserContact}
                    onChange={(e) => onUpdateControls({ endUserContact: e.target.value })}
                    placeholder="e.g. Procurement Department / Contact Name"
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>

                <div className="lg:col-span-2">
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    End-User Address
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.endUserAddress}
                    onChange={(e) => onUpdateControls({ endUserAddress: e.target.value })}
                    placeholder="e.g. Asmara, Eritrea"
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Currency (K3:L3)
                  </label>
                  <select
                    disabled={isLocked}
                    value={job.controls.currency}
                    onChange={(e) => onUpdateControls({ currency: e.target.value })}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  >
                    {CURRENCY_OPTIONS.map((c) => (
                      <option key={c} value={c}>{c}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Incoterm (K4:L4)
                  </label>
                  <select
                    disabled={isLocked}
                    value={job.controls.incoterm}
                    onChange={(e) => onUpdateControls({ incoterm: e.target.value })}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  >
                    {INCOTERM_OPTIONS.map((inc) => (
                      <option key={inc} value={inc}>{inc}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    VAT Rate (PI!L5 · Preset or Any %)
                  </label>
                  <div className="flex items-center gap-1.5">
                    <div className="inline-flex rounded-lg bg-slate-100 p-0.5 shrink-0">
                      {[
                        { label: '5%', value: 0.05 },
                        { label: '0%', value: 0 },
                      ].map((opt) => {
                        const active = Math.abs(job.controls.vatRate - opt.value) < 0.0001;
                        return (
                          <button
                            key={opt.label}
                            type="button"
                            disabled={isLocked}
                            onClick={() => onUpdateControls({ vatRate: opt.value })}
                            className={`px-2.5 min-h-[38px] rounded-md text-xs font-mono-tabular font-semibold transition-colors cursor-pointer disabled:opacity-50 ${
                              active
                                ? 'bg-slate-900 text-white'
                                : 'text-slate-600 hover:text-slate-900'
                            }`}
                          >
                            {opt.label}
                          </button>
                        );
                      })}
                    </div>
                    <div className="relative flex-1 min-w-[78px]">
                      <input
                        type="number"
                        inputMode="decimal"
                        step="0.5"
                        min="0"
                        max="100"
                        disabled={isLocked}
                        value={Number((job.controls.vatRate * 100).toFixed(2))}
                        onChange={(e) => {
                          if (e.target.value === '') {
                            onUpdateControls({ vatRate: 0 });
                            return;
                          }
                          const pct = Number(e.target.value);
                          if (!Number.isNaN(pct) && pct >= 0 && pct <= 100) {
                            onUpdateControls({ vatRate: Math.round(pct * 100) / 10000 });
                          }
                        }}
                        aria-label="Custom VAT percentage"
                        className="w-full min-h-[42px] rounded-lg border border-slate-200 bg-white pl-2.5 pr-6 py-2 text-sm font-mono-tabular text-right text-blue-700 font-semibold focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                      />
                      <span className="pointer-events-none absolute right-2.5 top-1/2 -translate-y-1/2 text-xs font-mono-tabular text-slate-500">
                        %
                      </span>
                    </div>
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Destination (K6:L6)
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.destination}
                    onChange={(e) => onUpdateControls({ destination: e.target.value })}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    RFQ / Tender Ref.
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.rfqRef}
                    onChange={(e) => onUpdateControls({ rfqRef: e.target.value })}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Customer PO Ref.
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.customerPoRef}
                    onChange={(e) => onUpdateControls({ customerPoRef: e.target.value })}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Payment Terms
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.paymentTerms}
                    onChange={(e) => onUpdateControls({ paymentTerms: e.target.value })}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>

                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Delivery Lead Time
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.deliveryLeadTime}
                    onChange={(e) => onUpdateControls({ deliveryLeadTime: e.target.value })}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
              </div>
            )}

            {/* Editable Per-Job Remarks & Signature Lines (Available on ALL 6 Documents — Leave Blank to Omit) */}
            <div className="pt-4 border-t border-slate-100 space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="text-xs font-semibold text-slate-800">
                  Job Remarks &amp; Signature Lines (Editable per job — clear any field to leave blank)
                </span>
                {!isLocked && (
                  <button
                    type="button"
                    onClick={() =>
                      onUpdateControls({
                        remarks: '',
                        authorizedSignatory: '',
                        counterpartySignatory: '',
                      })
                    }
                    className="text-[11px] font-medium text-slate-500 hover:text-red-600 cursor-pointer"
                  >
                    Clear Remarks &amp; Signatures
                  </button>
                )}
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="md:col-span-2">
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Job Remarks / Special Notes (Shown above signatures; leave empty to hide)
                  </label>
                  <textarea
                    rows={2}
                    disabled={isLocked}
                    value={job.controls.remarks}
                    onChange={(e) => onUpdateControls({ remarks: e.target.value })}
                    placeholder="Optional job-specific remarks, packaging notes, or leave blank..."
                    className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Left Signature Label / Signer Name (Leave empty to hide)
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.authorizedSignatory ?? ''}
                    onChange={(e) => onUpdateControls({ authorizedSignatory: e.target.value })}
                    placeholder={`e.g. Authorized Signature — ${company.name} (or leave blank)`}
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1.5">
                    Right Signature Label / Receiver Name (Leave empty to hide)
                  </label>
                  <input
                    type="text"
                    disabled={isLocked}
                    value={job.controls.counterpartySignatory ?? ''}
                    onChange={(e) => onUpdateControls({ counterpartySignatory: e.target.value })}
                    placeholder="e.g. Received in Good Order — Name & Stamp (or leave blank)"
                    className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Per-Job Custom Columns Bar (Part No., CAS No., Brand, HS Code — Flexible Placement Across All 6 Docs) */}
          <div className="rounded-lg border border-slate-200 bg-white px-5 py-3.5 space-y-3">
            <div className="flex flex-wrap items-start justify-between gap-2">
              <div>
                <span className="text-xs font-semibold text-slate-900">
                  Extra Order Columns (Choose Exact Position · Reflected Across QT, PI, PO, PL, DN, CI &amp; Excel)
                </span>
                <p className="text-[11px] text-slate-500">
                  Add Part No., CAS No., HS Code, or any custom column and place or move it anywhere in the table with zero formula impact.
                </p>
              </div>
              {customColumns.length > 0 && (
                <div className="flex flex-wrap items-center gap-2">
                  {customColumns.map((col) => (
                    <span
                      key={col.id}
                      className="inline-flex items-center gap-1.5 rounded-md border border-slate-300 bg-slate-50 pl-2.5 pr-1.5 py-1 text-xs font-semibold text-slate-800"
                    >
                      <span>{col.label}</span>
                      <select
                        disabled={isLocked}
                        value={col.position || 'after_no'}
                        onChange={(e) =>
                          handleMoveCustomColumn(col.id, e.target.value as CustomColumnPosition)
                        }
                        aria-label={`Position for ${col.label} column`}
                        className="rounded border border-slate-200 bg-white px-1.5 py-0.5 text-[11px] font-medium text-blue-700 focus:border-slate-900 focus:outline-none disabled:opacity-60 cursor-pointer"
                      >
                        {CUSTOM_COLUMN_POSITIONS.map((p) => (
                          <option key={p.value} value={p.value}>
                            {p.label}
                          </option>
                        ))}
                      </select>
                      {!isLocked && (
                        <button
                          type="button"
                          onClick={() => handleRemoveCustomColumn(col.id)}
                          className="p-0.5 text-slate-400 hover:text-red-600 cursor-pointer"
                          aria-label={`Remove ${col.label} column`}
                        >
                          <X className="w-3.5 h-3.5" />
                        </button>
                      )}
                    </span>
                  ))}
                </div>
              )}
            </div>

            {!isLocked && (
              <div className="flex flex-wrap items-center gap-2 pt-1">
                <div className="inline-flex items-center gap-1.5 rounded-md border border-slate-200 bg-slate-50 px-2 py-1">
                  <span className="text-[11px] font-semibold text-slate-600 whitespace-nowrap">
                    Place:
                  </span>
                  <select
                    value={newColPosition}
                    onChange={(e) => setNewColPosition(e.target.value as CustomColumnPosition)}
                    aria-label="Choose where to place new extra column"
                    className="bg-transparent text-xs font-semibold text-blue-700 focus:outline-none cursor-pointer"
                  >
                    {CUSTOM_COLUMN_POSITIONS.map((p) => (
                      <option key={p.value} value={p.value}>
                        {p.label}
                      </option>
                    ))}
                  </select>
                </div>

                {PRESET_CUSTOM_COLUMNS.filter(
                  (p) => !customColumns.some((c) => c.label.toLowerCase() === p.toLowerCase())
                ).map((preset) => (
                  <button
                    key={preset}
                    type="button"
                    onClick={() => handleAddCustomColumn(preset, newColPosition)}
                    className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-700 hover:border-slate-900 hover:text-slate-900 cursor-pointer"
                  >
                    <Plus className="w-3 h-3" />
                    <span>{preset}</span>
                  </button>
                ))}
                <div className="inline-flex items-center gap-1.5">
                  <input
                    type="text"
                    value={newColLabel}
                    onChange={(e) => setNewColLabel(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') handleAddCustomColumn(newColLabel, newColPosition);
                    }}
                    placeholder="Other column (e.g. Model #)..."
                    className="min-h-[32px] w-44 rounded-md border border-slate-200 px-2.5 py-1 text-xs text-slate-900 focus:border-slate-900 focus:outline-none"
                  />
                  <button
                    type="button"
                    onClick={() => handleAddCustomColumn(newColLabel, newColPosition)}
                    className="inline-flex items-center gap-1 rounded-md bg-slate-900 px-2.5 py-1.5 text-xs font-medium text-white hover:bg-slate-800 cursor-pointer"
                  >
                    <Plus className="w-3 h-3" />
                    <span>Add</span>
                  </button>
                </div>
              </div>
            )}
          </div>

          {/* Line Items Entry: iPad Touch Cards vs Windows Desktop Grid */}
          {platformMode === 'ipad' && (prefix === 'PI' || isLogistics) ? (
            <div className="space-y-3">
              <div className="flex items-center justify-between px-1">
                <h3 className="text-sm font-semibold text-slate-900">
                  {isLogistics
                    ? `${meta.shortName} Carton & Weight Cards`
                    : 'Single-Entry Item Cards (Excel Rows 17–36)'}
                </h3>
                <span className="text-xs font-mono-tabular text-slate-500">
                  {calc.activeRowCount} / 20 Rows
                </span>
              </div>
              <IPadJobEditor
                mode={isLogistics ? 'logistics' : 'entry'}
                items={job.items}
                computedRows={calc.rows}
                customColumns={customColumns}
                currency={curr}
                themeId={themeId}
                disabled={isLocked}
                onUpdateItem={onUpdateItem}
              />
            </div>
          ) : (
          <div className="rounded-lg border border-slate-200 bg-white overflow-hidden">
            <div className="flex flex-wrap items-center justify-between gap-2 px-5 py-3.5 border-b border-slate-200 bg-slate-50/70">
              <div>
                <h3 className="text-sm font-semibold text-slate-900">
                  {isLogistics
                    ? `${meta.shortName} Carton & Weight Entry (Fixed Rows 17–36)`
                    : isPO
                    ? 'Supplier PO Line Items (Linked from PI & PRICE CALC)'
                    : 'Single-Entry Item Schedule (PROFORMA INVOICE Rows 17–36)'}
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  {isLogistics
                    ? 'Verify carton numbers and net/gross weights independently before issuing. Never copy supplier PL blindly.'
                    : 'Requested Description = end-user wording · Offered Description = supplier wording we sell · Unit Price linked from PRICE CALC.'}
                </p>
              </div>
              <span className="text-xs text-slate-600 font-mono-tabular">
                {calc.activeRowCount} / 20 Rows Active
              </span>
            </div>

            <div className="overflow-x-auto sheet-scroll">
              <table className="w-full border-collapse text-xs">
                <thead>
                  <tr style={{ backgroundColor: theme.primary, color: theme.primaryText }}>
                    <th className="py-2.5 px-2.5 text-center font-semibold w-12">Row</th>
                    {colsAt('after_no').map((col) => (
                      <th key={col.id} className="py-2.5 px-2.5 text-left font-semibold min-w-[125px]">
                        {col.label}
                      </th>
                    ))}
                    <th className="py-2.5 px-3 text-left font-semibold min-w-[220px]">Requested Description</th>
                    {colsAt('after_requested').map((col) => (
                      <th key={col.id} className="py-2.5 px-2.5 text-left font-semibold min-w-[125px]">
                        {col.label}
                      </th>
                    ))}
                    <th className="py-2.5 px-3 text-left font-semibold min-w-[240px]">Offered Description</th>
                    {colsAt('after_offered').map((col) => (
                      <th key={col.id} className="py-2.5 px-2.5 text-left font-semibold min-w-[125px]">
                        {col.label}
                      </th>
                    ))}
                    <th className="py-2.5 px-2 text-center font-semibold w-20">UoM</th>
                    <th className="py-2.5 px-2 text-right font-semibold w-20">Qty</th>
                    {colsAt('after_qty').map((col) => (
                      <th key={col.id} className="py-2.5 px-2.5 text-left font-semibold min-w-[125px]">
                        {col.label}
                      </th>
                    ))}
                    {isLogistics ? (
                      <>
                        <th className="py-2.5 px-2 text-center font-semibold w-28">Carton No.</th>
                        <th className="py-2.5 px-2 text-right font-semibold w-24">Net Wt (KG)</th>
                        <th className="py-2.5 px-2 text-right font-semibold w-24">Gross Wt (KG)</th>
                      </>
                    ) : isPO ? (
                      <>
                        <th className="py-2.5 px-3 text-right font-semibold w-28">Unit Cost</th>
                        <th className="py-2.5 px-3 text-right font-semibold w-32">Total Cost</th>
                      </>
                    ) : (
                      <>
                        <th className="py-2.5 px-3 text-right font-semibold w-28">Unit Price (Col R)</th>
                        <th className="py-2.5 px-3 text-right font-semibold w-32">Total Price</th>
                      </>
                    )}
                    {colsAt('end').map((col) => (
                      <th key={col.id} className="py-2.5 px-2.5 text-left font-semibold min-w-[125px]">
                        {col.label}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {calc.rows.map((row, idx) => {
                    const raw = job.items[idx];
                    const canEditBase = prefix === 'PI' && !isLocked;
                    const canEditLogistics = isLogistics && !isLocked && !row.isEmpty;

                    const renderCustomCells = (pos: CustomColumnPosition) =>
                      colsAt(pos).map((col) => (
                        <td key={col.id} className="p-1">
                          {canEditBase ? (
                            <input
                              type="text"
                              value={raw.customFields?.[col.id] || ''}
                              placeholder={idx === 0 || !row.isEmpty ? `${col.label}...` : ''}
                              onChange={(e) =>
                                onUpdateItem(row.rowNumber, {
                                  customFields: {
                                    ...(raw.customFields || {}),
                                    [col.id]: e.target.value,
                                  },
                                })
                              }
                              className="w-full min-h-[36px] rounded border border-slate-200/80 hover:border-slate-400 focus:border-slate-900 bg-white/70 focus:bg-white px-2 py-1 font-mono-tabular text-blue-700 focus:outline-none"
                            />
                          ) : (
                            <div className="px-2 py-1.5 font-mono-tabular text-slate-800">
                              {row.customFields?.[col.id] || ''}
                            </div>
                          )}
                        </td>
                      ));

                    return (
                      <tr
                        key={row.rowNumber}
                        className="hover:bg-slate-50/90 transition-colors"
                        style={idx % 2 === 1 ? { backgroundColor: theme.zebraRow } : undefined}
                      >
                        <td className="py-1.5 px-2.5 text-center font-mono-tabular text-slate-400">
                          {row.rowNumber}
                        </td>
                        {renderCustomCells('after_no')}
                        <td className="p-1">
                          {canEditBase ? (
                            <input
                              type="text"
                              value={raw.requestedDescription}
                              placeholder={idx === 0 ? 'End user RFQ wording...' : ''}
                              onChange={(e) =>
                                onUpdateItem(row.rowNumber, { requestedDescription: e.target.value })
                              }
                              className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-slate-900 focus:outline-none"
                            />
                          ) : (
                            <div className="px-2 py-1.5 text-slate-800">{row.requestedDescription}</div>
                          )}
                        </td>
                        {renderCustomCells('after_requested')}
                        <td className="p-1">
                          {canEditBase ? (
                            <input
                              type="text"
                              value={raw.offeredDescription}
                              placeholder={idx === 0 ? 'Supplier offered wording...' : ''}
                              onChange={(e) =>
                                onUpdateItem(row.rowNumber, { offeredDescription: e.target.value })
                              }
                              className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-slate-900 focus:outline-none"
                            />
                          ) : (
                            <div className="px-2 py-1.5 text-slate-800">{row.offeredDescription}</div>
                          )}
                        </td>
                        {renderCustomCells('after_offered')}
                        <td className="p-1">
                          {canEditBase ? (
                            <select
                              value={raw.uom}
                              onChange={(e) => onUpdateItem(row.rowNumber, { uom: e.target.value })}
                              className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-1.5 py-1 text-center font-mono-tabular text-blue-700 focus:outline-none"
                            >
                              <option value="">—</option>
                              {UOM_OPTIONS.map((u) => (
                                <option key={u} value={u}>{u}</option>
                              ))}
                            </select>
                          ) : (
                            <div className="px-2 py-1.5 text-center font-mono-tabular text-slate-700">{row.uom}</div>
                          )}
                        </td>
                        <td className="p-1">
                          {canEditBase ? (
                            <input
                              type="number"
                              min="0"
                              step="any"
                              value={raw.qty ?? ''}
                              onChange={(e) =>
                                onUpdateItem(row.rowNumber, { qty: parseOptionalNumber(e.target.value) })
                              }
                              className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-right font-mono-tabular text-blue-700 font-medium focus:outline-none"
                            />
                          ) : (
                            <div className="px-2 py-1.5 text-right font-mono-tabular text-slate-800">
                              {row.qty ?? ''}
                            </div>
                          )}
                        </td>
                        {renderCustomCells('after_qty')}

                        {isLogistics ? (
                          <>
                            <td className="p-1">
                              <input
                                type="text"
                                disabled={!canEditLogistics}
                                value={raw.cartonNo}
                                placeholder={row.isEmpty ? '' : 'CTN 01-02'}
                                onChange={(e) =>
                                  onUpdateItem(row.rowNumber, { cartonNo: e.target.value })
                                }
                                className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-center font-mono-tabular text-blue-700 focus:outline-none disabled:opacity-40"
                              />
                            </td>
                            <td className="p-1">
                              <input
                                type="number"
                                step="0.1"
                                min="0"
                                disabled={!canEditLogistics}
                                value={raw.netWtKg ?? ''}
                                onChange={(e) =>
                                  onUpdateItem(row.rowNumber, { netWtKg: parseOptionalNumber(e.target.value) })
                                }
                                className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-right font-mono-tabular text-blue-700 focus:outline-none disabled:opacity-40"
                              />
                            </td>
                            <td className="p-1">
                              <input
                                type="number"
                                step="0.1"
                                min="0"
                                disabled={!canEditLogistics}
                                value={raw.grossWtKg ?? ''}
                                onChange={(e) =>
                                  onUpdateItem(row.rowNumber, { grossWtKg: parseOptionalNumber(e.target.value) })
                                }
                                className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-right font-mono-tabular text-blue-700 focus:outline-none disabled:opacity-40"
                              />
                            </td>
                          </>
                        ) : isPO ? (
                          <>
                            <td className="py-1.5 px-3 text-right font-mono-tabular text-slate-700">
                              {formatMoney(row.supplierUnitCost)}
                            </td>
                            <td className="py-1.5 px-3 text-right font-mono-tabular font-medium text-slate-900">
                              {formatMoney(row.supplierTotalCost)}
                            </td>
                          </>
                        ) : (
                          <>
                            <td className="py-1.5 px-3 text-right font-mono-tabular text-slate-700">
                              {formatMoney(row.activeUnitPrice)}
                            </td>
                            <td className="py-1.5 px-3 text-right font-mono-tabular font-medium text-slate-900">
                              {formatMoney(row.activeTotalPrice)}
                            </td>
                          </>
                        )}
                        {renderCustomCells('end')}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
          )}
        </div>
      )}

      {/* ==================== A4 EXECUTIVE PRINT / EXCEL PREVIEW ==================== */}
      {subView === 'preview' && (
        <div className="flex justify-center">
          <div className="a4-sheet w-full max-w-[960px] rounded-lg border border-slate-200 bg-white shadow-sm overflow-hidden">
            {/* Top Brand & Document Title Banner */}
            <div
              className="px-8 py-5 flex flex-wrap items-center justify-between gap-4"
              style={{ backgroundColor: theme.primary, color: theme.primaryText }}
            >
              <div className="flex items-center gap-4">
                {company.logoDataUrl && (
                  <div className="w-14 h-14 rounded-md bg-white p-1.5 flex items-center justify-center shrink-0">
                    <img
                      src={company.logoDataUrl}
                      alt={company.name}
                      className="max-w-full max-h-full object-contain"
                    />
                  </div>
                )}
                <div>
                  <div className="font-display text-2xl font-bold tracking-tight">
                    {company.name}
                  </div>
                  <div className="text-xs opacity-85 mt-0.5">
                    {company.tagline}
                  </div>
                </div>
              </div>
              <div className="text-right">
                <div className="font-display text-xl font-bold tracking-wide">
                  {meta.title}
                </div>
                <div className="text-xs font-mono-tabular opacity-90 mt-1">
                  {displayDocNo}
                </div>
              </div>
            </div>

            {/* Sub-header Contact Bar with Theme Accent Rule */}
            <div
              className="px-8 py-2.5 flex flex-wrap items-center justify-between gap-2 text-xs border-b-2"
              style={{
                backgroundColor: theme.secondaryBg,
                borderColor: theme.accent,
                color: theme.mutedText,
              }}
            >
              <span>
                {company.addressLine1}, {company.addressLine2} · {company.email}
              </span>
              <span className="font-mono-tabular font-medium text-slate-800">
                Date: {displayDate} · {company.trn}
              </span>
            </div>

            <div className="p-8 space-y-6">
              {/* Two-Column Metadata Header */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 text-xs">
                <div className="space-y-1.5">
                  <div className="font-semibold text-slate-500">{meta.recipientLabel}</div>
                  <div className="text-sm font-bold text-slate-900">
                    {isPO ? job.controls.supplierName || '—' : endUser || '—'}
                  </div>
                  <div className="text-slate-600">
                    {isPO ? job.controls.supplierAddress || '—' : job.controls.endUserContact}
                  </div>
                  {!isPO && <div className="text-slate-600">{job.controls.endUserAddress}</div>}
                </div>

                <div className="space-y-1.5 md:text-right">
                  <div className="grid grid-cols-2 gap-y-1.5 text-xs">
                    <span className="text-slate-500 md:text-right md:pr-3">Incoterm &amp; Destination</span>
                    <span className="font-medium text-slate-900">
                      {job.controls.incoterm} · {job.controls.destination}
                    </span>

                    <span className="text-slate-500 md:text-right md:pr-3">Currency</span>
                    <span className="font-mono-tabular font-medium text-slate-900">{curr}</span>

                    {prefix === 'PL' ? (
                      <>
                        <span className="text-slate-500 md:text-right md:pr-3">Ship Date</span>
                        <span className="font-mono-tabular font-medium text-slate-900">
                          {job.controls.shipDate || '—'}
                        </span>
                        <span className="text-slate-500 md:text-right md:pr-3">Port of Loading</span>
                        <span className="font-medium text-slate-900">{job.controls.portOfLoading}</span>
                      </>
                    ) : prefix === 'DN' ? (
                      <>
                        <span className="text-slate-500 md:text-right md:pr-3">Delivery Date</span>
                        <span className="font-mono-tabular font-medium text-slate-900">
                          {job.controls.deliveryDate || '—'}
                        </span>
                        <span className="text-slate-500 md:text-right md:pr-3">Transport Ref.</span>
                        <span className="font-mono-tabular font-medium text-slate-900">
                          {job.controls.transportRef || '—'}
                        </span>
                      </>
                    ) : isPO ? (
                      <>
                        <span className="text-slate-500 md:text-right md:pr-3">Supplier Quote Ref.</span>
                        <span className="font-mono-tabular font-medium text-slate-900">
                          {job.controls.supplierRef || '—'}
                        </span>
                        <span className="text-slate-500 md:text-right md:pr-3">PI Cross-Ref.</span>
                        <span className="font-mono-tabular font-medium text-slate-900">
                          {job.issuedDocs.PI?.docNumber || 'Pending PI'}
                        </span>
                      </>
                    ) : (
                      <>
                        <span className="text-slate-500 md:text-right md:pr-3">RFQ / Customer PO</span>
                        <span className="font-mono-tabular font-medium text-slate-900">
                          {[job.controls.rfqRef, job.controls.customerPoRef].filter(Boolean).join(' / ') || '—'}
                        </span>
                        <span className="text-slate-500 md:text-right md:pr-3">Payment Terms</span>
                        <span className="font-medium text-slate-900">{job.controls.paymentTerms}</span>
                      </>
                    )}
                  </div>
                </div>
              </div>

              {/* Strict Document Table */}
              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-xs">
                  <thead>
                    <tr
                      style={{
                        backgroundColor: theme.primary,
                        color: theme.primaryText,
                        borderBottom: `2px solid ${theme.strongBorder}`,
                      }}
                    >
                      <th className="py-3 px-2.5 text-center font-semibold w-10">No.</th>
                      {colsAt('after_no').map((col) => (
                        <th key={col.id} className="py-3 px-2.5 text-left font-semibold">
                          {col.label}
                        </th>
                      ))}
                      <th className="py-3 px-3 text-left font-semibold">Requested Description</th>
                      {colsAt('after_requested').map((col) => (
                        <th key={col.id} className="py-3 px-2.5 text-left font-semibold">
                          {col.label}
                        </th>
                      ))}
                      <th className="py-3 px-3 text-left font-semibold">Offered Description</th>
                      {colsAt('after_offered').map((col) => (
                        <th key={col.id} className="py-3 px-2.5 text-left font-semibold">
                          {col.label}
                        </th>
                      ))}
                      <th className="py-3 px-2 text-center font-semibold w-16">UoM</th>
                      <th className="py-3 px-2 text-right font-semibold w-16">Qty</th>
                      {colsAt('after_qty').map((col) => (
                        <th key={col.id} className="py-3 px-2.5 text-left font-semibold">
                          {col.label}
                        </th>
                      ))}
                      {isLogistics ? (
                        <>
                          <th className="py-3 px-2 text-center font-semibold w-24">Carton No.</th>
                          <th className="py-3 px-2.5 text-right font-semibold w-24">Net Wt (KG)</th>
                          <th className="py-3 px-2.5 text-right font-semibold w-24">Gross Wt (KG)</th>
                        </>
                      ) : isPO ? (
                        <>
                          <th className="py-3 px-2.5 text-right font-semibold w-28">Unit Cost ({curr})</th>
                          <th className="py-3 px-2.5 text-right font-semibold w-32">Total Cost ({curr})</th>
                        </>
                      ) : (
                        <>
                          <th className="py-3 px-2.5 text-right font-semibold w-28">Unit Price ({curr})</th>
                          <th className="py-3 px-2.5 text-right font-semibold w-32">Total Price ({curr})</th>
                        </>
                      )}
                      {colsAt('end').map((col) => (
                        <th key={col.id} className="py-3 px-2.5 text-left font-semibold">
                          {col.label}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {activeRows.length === 0 ? (
                      <tr>
                        <td
                          colSpan={(isLogistics ? 8 : 7) + customColumns.length}
                          className="py-10 text-center text-slate-400 border-b border-slate-200"
                        >
                          No items entered yet. Switch to Job Entry &amp; Controls to populate rows 17–36.
                        </td>
                      </tr>
                    ) : (
                      activeRows.map((row, idx) => {
                        const renderPreviewCustom = (pos: CustomColumnPosition) =>
                          colsAt(pos).map((col) => (
                            <td
                              key={col.id}
                              className="py-2.5 px-2.5 font-mono-tabular text-slate-800 whitespace-nowrap"
                            >
                              {row.customFields?.[col.id] || '—'}
                            </td>
                          ));

                        return (
                          <tr
                            key={row.rowNumber}
                            style={{
                              backgroundColor: idx % 2 === 1 ? theme.zebraRow : '#FFFFFF',
                              borderBottom: `1px solid ${theme.border}`,
                            }}
                          >
                            <td className="py-2.5 px-2.5 text-center font-mono-tabular text-slate-500">
                              {idx + 1}
                            </td>
                            {renderPreviewCustom('after_no')}
                            <td className="py-2.5 px-3 text-slate-800">{row.requestedDescription}</td>
                            {renderPreviewCustom('after_requested')}
                            <td className="py-2.5 px-3 text-slate-900 font-medium">{row.offeredDescription}</td>
                            {renderPreviewCustom('after_offered')}
                            <td className="py-2.5 px-2 text-center font-mono-tabular text-slate-700">{row.uom}</td>
                            <td className="py-2.5 px-2 text-right font-mono-tabular text-slate-900">{row.qty}</td>
                            {renderPreviewCustom('after_qty')}
                            {isLogistics ? (
                              <>
                                <td className="py-2.5 px-2 text-center font-mono-tabular text-slate-800">
                                  {row.cartonNo || '—'}
                                </td>
                                <td className="py-2.5 px-2.5 text-right font-mono-tabular text-slate-800">
                                  {formatNumber(row.netWtKg, 2)}
                                </td>
                                <td className="py-2.5 px-2.5 text-right font-mono-tabular font-medium text-slate-900">
                                  {formatNumber(row.grossWtKg, 2)}
                                </td>
                              </>
                            ) : isPO ? (
                              <>
                                <td className="py-2.5 px-2.5 text-right font-mono-tabular text-slate-800">
                                  {formatMoney(row.supplierUnitCost)}
                                </td>
                                <td className="py-2.5 px-2.5 text-right font-mono-tabular font-medium text-slate-900">
                                  {formatMoney(row.supplierTotalCost)}
                                </td>
                              </>
                            ) : (
                              <>
                                <td className="py-2.5 px-2.5 text-right font-mono-tabular text-slate-800">
                                  {formatMoney(row.activeUnitPrice)}
                                </td>
                                <td className="py-2.5 px-2.5 text-right font-mono-tabular font-medium text-slate-900">
                                  {formatMoney(row.activeTotalPrice)}
                                </td>
                              </>
                            )}
                            {renderPreviewCustom('end')}
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>

              {/* Totals Block */}
              <div className="flex justify-end">
                {isLogistics ? (
                  <div
                    className="w-full sm:w-96 rounded-md p-4 text-xs font-mono-tabular space-y-1.5 border"
                    style={{ backgroundColor: theme.totalBg, borderColor: theme.border }}
                  >
                    <div className="flex justify-between">
                      <span className="font-sans text-slate-600">Total Line Quantity</span>
                      <span className="font-semibold text-slate-900">{calc.totalQty}</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="font-sans text-slate-600">Total Net Weight</span>
                      <span className="font-semibold text-slate-900">{formatNumber(calc.plTotalNetWt, 2)} KG</span>
                    </div>
                    <div className="flex justify-between pt-1.5 border-t border-slate-300/60 text-sm font-bold text-slate-900">
                      <span className="font-sans">Total Gross Weight</span>
                      <span>{formatNumber(calc.plTotalGrossWt, 2)} KG</span>
                    </div>
                  </div>
                ) : isPO ? (
                  <div
                    className="w-full sm:w-80 rounded-md p-4 text-sm font-mono-tabular border"
                    style={{ backgroundColor: theme.totalBg, borderColor: theme.strongBorder }}
                  >
                    <div className="flex justify-between font-bold text-slate-900">
                      <span className="font-sans">TOTAL ORDER COST ({curr})</span>
                      <span>{formatMoney(calc.totalSupplierCost)}</span>
                    </div>
                  </div>
                ) : (
                  <div className="w-full sm:w-96 text-xs font-mono-tabular">
                    <div className="flex justify-between py-2 border-b border-slate-200">
                      <span className="font-sans text-slate-600">GOODS SUBTOTAL</span>
                      <span className="font-medium text-slate-900">{formatMoney(calc.goodsSubtotal, curr)}</span>
                    </div>
                    <div className="flex justify-between py-2 border-b border-slate-200">
                      <span className="font-sans text-slate-600">
                        VAT ({formatPct(job.controls.vatRate)} on Goods Only)
                      </span>
                      <span className="font-medium text-slate-900">{formatMoney(calc.vatAmount, curr)}</span>
                    </div>
                    <div className="flex justify-between py-2 border-b border-slate-200">
                      <span className="font-sans text-slate-600">
                        FREIGHT &amp; DOCUMENTATION ({job.controls.incoterm})
                      </span>
                      <span className="font-medium text-slate-900">{formatMoney(calc.customerFreight, curr)}</span>
                    </div>
                    <div
                      className="flex justify-between py-3 px-3 mt-1 rounded-sm text-sm font-bold border-t border-b-2"
                      style={{
                        backgroundColor: theme.totalBg,
                        borderColor: theme.strongBorder,
                        color: theme.primary,
                      }}
                    >
                      <span className="font-sans">GRAND TOTAL ({curr})</span>
                      <span>{formatMoney(calc.grandTotal)}</span>
                    </div>
                  </div>
                )}
              </div>

              {/* Terms, Bank & Signatures (Fully Editable Per Job — Omitted Cleanly When Empty) */}
              <div className="pt-6 border-t border-slate-200 space-y-5 text-xs text-slate-600">
                {job.controls.remarks.trim() && (
                  <div>
                    <span className="font-semibold text-slate-800">Remarks: </span>
                    <span className="whitespace-pre-line">{job.controls.remarks}</span>
                  </div>
                )}
                {(prefix === 'PI' || prefix === 'CI') && (
                  <div className="font-mono-tabular text-[11px] text-slate-600">
                    {company.bankDetails}
                  </div>
                )}

                {((job.controls.authorizedSignatory ?? '').trim() ||
                  (job.controls.counterpartySignatory ?? '').trim()) && (
                  <div className="grid grid-cols-2 gap-10 pt-8">
                    <div>
                      {(job.controls.authorizedSignatory ?? '').trim() && (
                        <div className="border-t border-slate-400 pt-2 font-medium text-slate-800">
                          {job.controls.authorizedSignatory}
                        </div>
                      )}
                    </div>
                    <div>
                      {(job.controls.counterpartySignatory ?? '').trim() && (
                        <div className="border-t border-slate-400 pt-2 font-medium text-slate-800">
                          {job.controls.counterpartySignatory}
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
