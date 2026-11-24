import React, { useEffect, useRef, useState } from 'react';
import {
  AlertCircle,
  Save,
  Unlock,
  AlertTriangle,
  ArrowLeft,
  ArrowRight,
  CheckCircle2,
  ChevronDown,
  ChevronRight,
  Circle,
  Copy,
  Eye,
  FileSpreadsheet,
  Lock,
  Plus,
  Printer,
  RotateCcw,
  ShieldCheck,
  SlidersHorizontal,
  Table as TableIcon,
  X,
} from 'lucide-react';
import {
  CURRENCY_OPTIONS,
  DOC_META,
  INCOTERM_OPTIONS,
  resolveDocumentColumns,
  UOM_OPTIONS,
} from '../constants/themes';
import { formatDocNumber } from '../store/storage';
import {
  CompanyProfile,
  ConfirmDialogRequest,
  CustomColumnDef,
  CustomColumnPosition,
  DocPrefix,
  InquiryType,
  JobCalculationSummary,
  JobControls,
  JobItemRow,
  JobWorkflowState,
  PlatformMode,
  ProcurementJob,
  ThemePalette,
  ValidationIssue,
  WorkspaceTab,
} from '../types/procurement';
import {
  formatMoney,
  formatNumber,
  getResolvedEndUser,
  getWorkflowGateStatus,
  validateForIssue,
} from '../utils/calculations';
import { exportIssuedDocumentXlsx } from '../utils/excelEngine';
import { HelpTip } from './HelpTip';
import { IPadJobEditor } from './IPadJobEditor';
import { useRowSelection } from '../hooks/useRowSelection';
import { PasteIntoRows } from './PasteIntoRows';
import { parseFlexibleNumber } from '../utils/supplierQuoteParser';

interface DocumentWorkspaceViewProps {
  job: ProcurementJob;
  prefix: DocPrefix;
  calc: JobCalculationSummary;
  theme: ThemePalette;
  company: CompanyProfile;
  platformMode: PlatformMode;
  nextNewJobSeq: number;
  onUpdateJobName: (name: string) => void;
  onUpdateInquiryType: (inquiryType: InquiryType) => void;
  onUpdateControls: (patch: Partial<JobControls>) => void;
  onUpdateWorkflow: (patch: Partial<JobWorkflowState>) => void;
  onUpdateCustomColumns: (cols: CustomColumnDef[], nextOrder?: string[]) => void;
  onUpdateColumnOrder: (order: string[] | undefined) => void;
  onUpdateItem: (rowNumber: number, patch: Partial<JobItemRow>) => void;
  onIssueDocument: (prefix: DocPrefix) => void;
  onDuplicateAsNewJob: () => void;
  onNavigateTab: (tab: WorkspaceTab) => void;
  onRequestConfirm: (req: ConfirmDialogRequest) => void;
  onNotify: (msg: string) => void;
  amending: boolean;
  amendingPrefix: DocPrefix | null;
  onRequestUnlock: (prefix: DocPrefix) => void;
  onRequestSave: (prefix: DocPrefix) => void;
  onDiscardAmendment: () => void;
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
  if (!Number.isFinite(n) || n < 0) return null;
  return n;
}

export const DocumentWorkspaceView: React.FC<DocumentWorkspaceViewProps> = ({
  job,
  prefix,
  calc,
  theme,
  company,
  platformMode,
  nextNewJobSeq,
  onUpdateJobName,
  onUpdateInquiryType,
  onUpdateControls,
  onUpdateWorkflow,
  onUpdateCustomColumns,
  onUpdateColumnOrder,
  onUpdateItem,
  onIssueDocument,
  onDuplicateAsNewJob,
  onNavigateTab,
  onRequestConfirm,
  onNotify,
  amending,
  amendingPrefix,
  onRequestUnlock,
  onRequestSave,
  onDiscardAmendment,
}) => {
  const [subView, setSubView] = useState<'edit' | 'preview'>(prefix === 'QT' ? 'edit' : 'preview');
  const [controlsOpen, setControlsOpen] = useState(true);
  const [remarksOpen, setRemarksOpen] = useState(false);
  const [extraColsOpen, setExtraColsOpen] = useState(false);
  const [dragColKey, setDragColKey] = useState<string | null>(null);
  const [validationErrors, setValidationErrors] = useState<ValidationIssue[]>([]);
  const [exporting, setExporting] = useState(false);
  const [newColLabel, setNewColLabel] = useState('');
  const [newColPosition, setNewColPosition] = useState<CustomColumnPosition>('after_no');
  const printTimerRef = useRef<number | null>(null);

  useEffect(() => {
    return () => {
      if (printTimerRef.current !== null) {
        window.clearTimeout(printTimerRef.current);
      }
    };
  }, []);

  const meta = DOC_META[prefix];
  const issued = job.issuedDocs[prefix];
  const isIssued = Boolean(issued && !issued.isVoided);
  // Locked unless this job is open for an amendment (one document unlocked at a time)
  const isLocked = isIssued && !amending;
  const isAmendingThisDoc = amending && amendingPrefix === prefix;
  const isLogistics = prefix === 'PL' || prefix === 'DN';
  const isPO = prefix === 'PO';
  const endUser = getResolvedEndUser(job);
  const curr = job.controls.currency;
  const activeRows = calc.rows.filter((r) => !r.isEmpty);
  const rowSel = useRowSelection<number>();
  const [bulkUom, setBulkUom] = useState('');
  const selectableRowIds = activeRows.map((r) => r.rowNumber);
  const applyBulkRows = (patch: Partial<JobItemRow>) => {
    rowSel.selected.forEach((rn) => onUpdateItem(rn, patch));
  };
  const customColumns = job.customColumns || [];
  const customColMap = new Map(customColumns.map((c) => [`custom:${c.id}`, c]));

  // Unified ordered column keys for both Edit Schedule and A4 Print Preview
  const orderedColKeys = resolveDocumentColumns(job.columnOrder, customColumns, isLogistics);
  const defaultOrder = resolveDocumentColumns(undefined, customColumns, isLogistics);
  const hasCustomOrder =
    orderedColKeys.length === defaultOrder.length &&
    orderedColKeys.some((k, idx) => k !== defaultOrder[idx]);

  const spliceKeyIntoOrder = (
    baseOrder: string[],
    newKey: string,
    position: CustomColumnPosition
  ): string[] => {
    const next = baseOrder.filter((k) => k !== newKey);
    if (position === 'after_no') {
      next.unshift(newKey);
    } else if (position === 'after_requested') {
      const idx = next.indexOf('requested');
      next.splice(idx >= 0 ? idx + 1 : 0, 0, newKey);
    } else if (position === 'after_offered') {
      const idx = next.indexOf('offered');
      next.splice(idx >= 0 ? idx + 1 : next.length, 0, newKey);
    } else if (position === 'after_qty') {
      const idx = next.indexOf('qty');
      next.splice(idx >= 0 ? idx + 1 : next.length, 0, newKey);
    } else {
      next.push(newKey);
    }
    return next;
  };

  // Merge updated order back into the master cross-document columnOrder (preserving pricing/logistics-only keys)
  const commitReorderedVisibleKeys = (nextVisible: string[]) => {
    const existing = job.columnOrder || [];
    const visibleSet = new Set(nextVisible);
    const otherKeys = existing.filter((k) => !visibleSet.has(k));
    onUpdateColumnOrder([...nextVisible, ...otherKeys]);
  };

  const handleMoveColumnByOffset = (key: string, delta: -1 | 1) => {
    if (isLocked) return;
    const idx = orderedColKeys.indexOf(key);
    const target = idx + delta;
    if (idx < 0 || target < 0 || target >= orderedColKeys.length) return;
    const next = [...orderedColKeys];
    const [moved] = next.splice(idx, 1);
    next.splice(target, 0, moved);
    commitReorderedVisibleKeys(next);
  };

  const handleDropColumn = (sourceKey: string, targetKey: string) => {
    if (isLocked || !sourceKey || sourceKey === targetKey) return;
    const fromIdx = orderedColKeys.indexOf(sourceKey);
    const toIdx = orderedColKeys.indexOf(targetKey);
    if (fromIdx < 0 || toIdx < 0) return;
    const next = [...orderedColKeys];
    const [moved] = next.splice(fromIdx, 1);
    next.splice(toIdx, 0, moved);
    commitReorderedVisibleKeys(next);
  };

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
    const nextCols = [...customColumns, { id, label, position }];
    const nextOrder = job.columnOrder
      ? spliceKeyIntoOrder(job.columnOrder, `custom:${id}`, position)
      : undefined;
    onUpdateCustomColumns(nextCols, nextOrder);
    setNewColLabel('');
    const posLabel = CUSTOM_COLUMN_POSITIONS.find((p) => p.value === position)?.label || 'After No.';
    onNotify(`Added "${label}" (${posLabel}) across all 6 documents and Excel exports.`);
  };

  const handleMoveCustomColumn = (id: string, position: CustomColumnPosition) => {
    if (isLocked) return;
    const nextCols = customColumns.map((c) => (c.id === id ? { ...c, position } : c));
    const nextOrder = job.columnOrder
      ? spliceKeyIntoOrder(job.columnOrder, `custom:${id}`, position)
      : undefined;
    onUpdateCustomColumns(nextCols, nextOrder);
  };

  const handleRemoveCustomColumn = (id: string) => {
    if (isLocked) return;
    const col = customColumns.find((c) => c.id === id);
    if (!col) return;
    const populatedCount = job.items.filter(
      (it) => it.customFields?.[id] && it.customFields[id].trim() !== ''
    ).length;

    const removeNow = () => {
      const nextCols = customColumns.filter((c) => c.id !== id);
      const nextOrder = job.columnOrder
        ? job.columnOrder.filter((k) => k !== `custom:${id}`)
        : undefined;
      onUpdateCustomColumns(nextCols, nextOrder);
    };

    if (populatedCount > 0) {
      onRequestConfirm({
        title: `Remove "${col.label}" Column?`,
        message: `${populatedCount} line item${populatedCount === 1 ? '' : 's'} currently have values entered under "${col.label}". Removing this column hides it from all 6 documents and Excel exports.`,
        confirmLabel: 'Remove Column',
        destructive: true,
        onConfirm: removeNow,
      });
    } else {
      removeNow();
    }
  };

  const gateStatus = getWorkflowGateStatus(job, prefix);
  const plannedSeq = job.jobSequence || nextNewJobSeq;
  const plannedYear = job.jobSequenceYear || new Date().getFullYear();
  const plannedDocNo = formatDocNumber(prefix, plannedSeq, plannedYear);

  const handleTryIssue = () => {
    const errors = validateForIssue(job, prefix, calc);
    if (errors.length > 0) {
      setValidationErrors(errors);
      return;
    }
    setValidationErrors([]);

    if (!gateStatus.isReady) {
      const unmet = gateStatus.requirements
        .filter((r) => !r.met)
        .map((r) => `• ${r.label}`)
        .join('\n');
      onRequestConfirm({
        title: `Prerequisite Step Incomplete — Issue Anyway?`,
        message: `Recommended stage prerequisites are not yet complete:\n${unmet}\n\n${gateStatus.guidance}\n\nYou can complete those steps first, or issue ${plannedDocNo} manually now.`,
        confirmLabel: `Issue Anyway ()`,
        destructive: false,
        onConfirm: () => onIssueDocument(prefix),
      });
      return;
    }

    onIssueDocument(prefix);
  };

  const handleExportXlsx = async () => {
    setExporting(true);
    try {
      const filename = await exportIssuedDocumentXlsx(job, prefix, theme, company);
      onNotify(`Exported values-only Excel file: ${filename}`);
    } catch {
      onNotify('Failed to generate Excel file. Please try again.');
    } finally {
      setExporting(false);
    }
  };

  const handlePrint = () => {
    if (printTimerRef.current !== null) {
      window.clearTimeout(printTimerRef.current);
      printTimerRef.current = null;
    }
    if (subView !== 'preview') {
      setSubView('preview');
      requestAnimationFrame(() => {
        printTimerRef.current = window.setTimeout(() => {
          printTimerRef.current = null;
          window.print();
        }, 120);
      });
    } else {
      window.print();
    }
  };

  // Never show "DRAFT (Next: ...)" on A4 print, Excel preview, or PDF print — only render doc no. once issued
  const issuedDocNo = isLocked && issued ? issued.docNumber : '';
  const displayDate = isLocked && issued
    ? issued.issuedAt.slice(0, 10)
    : new Date().toISOString().slice(0, 10);

  const getColumnHeaderMeta = (
    key: string,
    forPreview: boolean
  ): { label: string; align: 'left' | 'center' | 'right'; widthClass: string } => {
    if (key.startsWith('custom:')) {
      const cc = customColMap.get(key);
      return {
        label: cc?.label || 'Custom',
        align: 'left',
        widthClass: forPreview ? '' : 'min-w-[125px]',
      };
    }
    switch (key) {
      case 'requested':
        return {
          label: 'Requested Description',
          align: 'left',
          widthClass: forPreview ? '' : 'min-w-[210px]',
        };
      case 'offered':
        return {
          label: 'Offered Description',
          align: 'left',
          widthClass: forPreview ? '' : 'min-w-[230px]',
        };
      case 'uom':
        return { label: 'UoM', align: 'center', widthClass: forPreview ? 'w-16' : 'w-20' };
      case 'qty':
        return { label: 'Qty', align: 'right', widthClass: forPreview ? 'w-16' : 'w-20' };
      case 'packages':
        return { label: 'Packages (boxes · L×W×H cm)', align: 'center', widthClass: forPreview ? 'w-32' : 'w-36' };
      case 'carton':
        return { label: 'Carton No.', align: 'center', widthClass: forPreview ? 'w-24' : 'w-28' };
      case 'net_wt':
        return { label: 'Net Wt (KG)', align: 'right', widthClass: 'w-24' };
      case 'gross_wt':
        return { label: 'Gross Wt (KG)', align: 'right', widthClass: 'w-24' };
      case 'unit_val':
        return isPO
          ? {
              label: forPreview ? `Unit Cost (${curr})` : 'Unit Cost',
              align: 'right',
              widthClass: 'w-28',
            }
          : {
              label: forPreview ? `Unit Price (${curr})` : 'Unit Price (Col R)',
              align: 'right',
              widthClass: 'w-28',
            };
      case 'total_val':
        return isPO
          ? {
              label: forPreview ? `Total Cost (${curr})` : 'Total Cost',
              align: 'right',
              widthClass: 'w-32',
            }
          : {
              label: forPreview ? `Total Price (${curr})` : 'Total Price',
              align: 'right',
              widthClass: 'w-32',
            };
      default:
        return { label: key, align: 'left', widthClass: '' };
    }
  };

  const controlsSummary =
    prefix === 'PO'
      ? `${job.controls.supplierName || 'No Supplier'} · ${job.controls.supplierRef || 'No Ref'}`
      : prefix === 'PL'
      ? `${job.controls.shipDate || 'No Ship Date'} · ${job.controls.portOfLoading || '—'}`
      : prefix === 'DN'
      ? `${job.controls.deliveryDate || 'No Delivery Date'} · ${job.controls.transportRef || '—'}`
      : prefix === 'PI'
      ? `${job.controls.customerPoRef || 'No PO Ref'} · ${job.controls.paymentTerms || '—'}`
      : prefix === 'CI'
      ? `${job.controls.customerPoRef || 'No PO Ref'} · ${job.controls.incoterm} ${job.controls.destination}`
      : `${endUser || 'No Client'} · ${job.inquiryType || 'RFQ'} · ${curr} · ${job.controls.incoterm} ${job.controls.destination}`;

  const remarksFilledCount = [
    job.controls.remarks.trim(),
    (job.controls.authorizedSignatory ?? '').trim(),
    (job.controls.counterpartySignatory ?? '').trim(),
  ].filter(Boolean).length;

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
              <span>{prefix === 'QT' ? 'Master Job Controls' : 'Document Inputs'}</span>
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

          {/* Show Issued badge OR Job Shared Sequence preview (strictly no-print; never shown on A4) */}
          {isLocked && issued ? (
            <div className="text-xs flex items-center gap-2 font-mono-tabular">
              <span className="inline-flex items-center gap-1.5 font-semibold text-emerald-700">
                <CheckCircle2 className="w-3.5 h-3.5" />
                Issued · {issued.docNumber}{issued.revision && issued.revision > 1 ? ` · Rev ${issued.revision}` : ""}
              </span>
            </div>
          ) : job.jobSequence ? (
            <div className="text-xs flex items-center gap-1.5 font-mono-tabular text-slate-600">
              <span>Shared Job Seq #{String(job.jobSequence).padStart(3, '0')}</span>
              <span className="text-slate-300">→</span>
              <span className="font-semibold text-slate-900">{plannedDocNo}</span>
            </div>
          ) : null}
        </div>

        {/* Right Actions: Issue, Export Values-Only .xlsx, Print PDF */}
        <div className="flex flex-wrap items-center gap-2">
          {amending && (
            <>
              {isAmendingThisDoc && (
                <button
                  type="button"
                  onClick={() => onRequestSave(prefix)}
                  className="inline-flex items-center gap-1.5 rounded-lg px-4 py-2 text-xs font-semibold text-white transition-opacity hover:opacity-90 cursor-pointer min-h-[40px] whitespace-nowrap"
                  style={{ backgroundColor: theme.primary }}
                >
                  <Save className="w-3.5 h-3.5" />
                  <span>Save as Rev {(issued?.revision ?? 1) + 1}</span>
                </button>
              )}
              <button
                type="button"
                onClick={onDiscardAmendment}
                className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[40px] whitespace-nowrap"
              >
                <span>Discard changes</span>
              </button>
            </>
          )}
          {isIssued && !amending && (
            <button
              type="button"
              onClick={() => onRequestUnlock(prefix)}
              className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3.5 py-2 text-xs font-semibold text-slate-900 hover:bg-slate-50 cursor-pointer min-h-[40px] whitespace-nowrap"
            >
              <Unlock className="w-3.5 h-3.5" />
              <span>Unlock to amend</span>
            </button>
          )}
          {amending ? null : isLocked ? (
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

      {/* Guided Staged Workflow Gate Banner (Strictly no-print; zero auto-issuing) */}
      {!isLocked && (
        <div
          className={`no-print rounded-lg border p-4 space-y-3 ${
            gateStatus.isReady
              ? 'border-emerald-200 bg-emerald-50/50'
              : 'border-amber-200 bg-amber-50/60'
          }`}
        >
          <div className="flex flex-wrap items-start justify-between gap-3">
            <div className="flex items-start gap-2.5 min-w-0">
              {gateStatus.isReady ? (
                <ShieldCheck className="w-4 h-4 text-emerald-700 mt-0.5 shrink-0" />
              ) : (
                <AlertTriangle className="w-4 h-4 text-amber-700 mt-0.5 shrink-0" />
              )}
              <div className="space-y-1 min-w-0">
                <div className="text-xs font-semibold text-slate-900">
                  {gateStatus.title}
                </div>
                <p className="text-xs text-slate-700">{gateStatus.guidance}</p>
                <div className="flex flex-wrap items-center gap-x-4 gap-y-1 pt-1">
                  {gateStatus.requirements.map((req) => (
                    <span
                      key={req.id}
                      className={`inline-flex items-center gap-1.5 text-xs ${
                        req.met ? 'text-emerald-800 font-medium' : 'text-amber-900'
                      }`}
                    >
                      {req.met ? (
                        <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
                      ) : (
                        <Circle className="w-3.5 h-3.5 text-amber-500 shrink-0" />
                      )}
                      <span>{req.label}</span>
                    </span>
                  ))}
                </div>
              </div>
            </div>

            {gateStatus.recommendedTab && gateStatus.recommendedTabLabel && (
              <button
                type="button"
                onClick={() => onNavigateTab(gateStatus.recommendedTab!)}
                className="inline-flex items-center gap-1.5 rounded-md border border-slate-300 bg-white px-3 py-1.5 text-xs font-semibold text-slate-800 hover:bg-slate-50 cursor-pointer shrink-0 whitespace-nowrap"
              >
                <span>{gateStatus.recommendedTabLabel}</span>
                <ArrowRight className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          {/* Stage 3 Inline Confirmation: Client PO Received (on PO and CI) */}
          {(prefix === 'PO' || prefix === 'CI') && (
            <div className="pt-3 border-t border-slate-200/80 grid grid-cols-1 md:grid-cols-12 gap-3 items-center">
              <label className="md:col-span-7 flex items-center gap-2.5 text-xs font-semibold text-slate-900 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={Boolean(job.workflow?.endUserPoConfirmed)}
                  onChange={(e) => {
                    const checked = e.target.checked;
                    onUpdateWorkflow({
                      endUserPoConfirmed: checked,
                      endUserPoConfirmedAt: checked ? new Date().toISOString() : undefined,
                    });
                  }}
                  className="w-4 h-4 rounded border-slate-300 accent-slate-900 cursor-pointer"
                />
                <span>Confirm: Purchase Order from Client / End-User Received</span>
              </label>
              <div className="md:col-span-5">
                <input
                  type="text"
                  value={job.controls.customerPoRef}
                  onChange={(e) => onUpdateControls({ customerPoRef: e.target.value })}
                  placeholder="Client PO Number (e.g. PO-CLIENT-2026-08)"
                  className="w-full min-h-[36px] rounded-md border border-slate-300 bg-white px-3 py-1.5 text-xs font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none"
                />
              </div>
            </div>
          )}

          {/* Stage 4 Inline Confirmations: Payment Done to Supplier + Supplier Packing List Received (on PL) */}
          {prefix === 'PL' && (
            <div className="pt-3 border-t border-slate-200/80 grid grid-cols-1 md:grid-cols-2 gap-3">
              <div className="rounded-md border border-slate-200 bg-white p-3 space-y-2">
                <label className="flex items-center gap-2.5 text-xs font-semibold text-slate-900 cursor-pointer select-none">
                  <input
                    type="checkbox"
                    checked={Boolean(job.workflow?.supplierPaidConfirmed)}
                    onChange={(e) => {
                      const checked = e.target.checked;
                      onUpdateWorkflow({
                        supplierPaidConfirmed: checked,
                        supplierPaidConfirmedAt: checked ? new Date().toISOString() : undefined,
                      });
                    }}
                    className="w-4 h-4 rounded border-slate-300 accent-slate-900 cursor-pointer"
                  />
                  <span>1. Payment Completed to Supplier</span>
                </label>
                <input
                  type="text"
                  value={job.workflow?.supplierPaymentRef || ''}
                  onChange={(e) => onUpdateWorkflow({ supplierPaymentRef: e.target.value })}
                  placeholder="TT / Wire Reference (optional)"
                  className="w-full min-h-[34px] rounded-md border border-slate-200 bg-slate-50/70 px-2.5 py-1 text-xs font-mono-tabular text-blue-700 focus:border-slate-900 focus:bg-white focus:outline-none"
                />
              </div>

              <div className="rounded-md border border-slate-200 bg-white p-3 space-y-2">
                <label className="flex items-center gap-2.5 text-xs font-semibold text-slate-900 cursor-pointer select-none">
                  <input
                    type="checkbox"
                    checked={Boolean(job.workflow?.supplierPackingListReceived)}
                    onChange={(e) => {
                      const checked = e.target.checked;
                      onUpdateWorkflow({
                        supplierPackingListReceived: checked,
                        supplierPackingListAt: checked ? new Date().toISOString() : undefined,
                      });
                    }}
                    className="w-4 h-4 rounded border-slate-300 accent-slate-900 cursor-pointer"
                  />
                  <span>2. Supplier Packing List Received</span>
                </label>
                <input
                  type="text"
                  value={job.workflow?.supplierPackingListRef || ''}
                  onChange={(e) => onUpdateWorkflow({ supplierPackingListRef: e.target.value })}
                  placeholder="Supplier PL Ref / AWB (optional)"
                  className="w-full min-h-[34px] rounded-md border border-slate-200 bg-slate-50/70 px-2.5 py-1 text-xs font-mono-tabular text-blue-700 focus:border-slate-900 focus:bg-white focus:outline-none"
                />
              </div>
            </div>
          )}
        </div>
      )}

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

      {amending && subView === "edit" && (
        <div className="no-print rounded-lg border border-amber-200 bg-amber-50 px-4 py-3 text-xs text-amber-900">
          Job is open for amendment{amendingPrefix ? ` ()` : ""}. Every change is logged; save a revision to record the reason.
        </div>
      )}
      {/* Locked Notice when Issued */}
      {isLocked && subView === 'edit' && (
        <div className="no-print rounded-lg border border-emerald-200 bg-emerald-50/70 px-4 py-3 text-xs text-emerald-900 flex items-center justify-between gap-3">
          <span>
            <strong>{issued?.docNumber}</strong> is frozen and logged in the Master Registry.
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
        <div className="no-print space-y-4">
          {/* Collapsible 1: Master Job Controls (on QT) OR Document-Specific Inputs (on PI/PO/CI/PL/DN) */}
          <div className="rounded-lg border border-slate-200 bg-white overflow-visible">
            <div className="flex flex-wrap items-center justify-between gap-2 px-5 py-3.5 hover:bg-slate-50/70 transition-colors rounded-t-lg">
              <div className="flex items-center gap-2 min-w-0">
                <button
                  type="button"
                  onClick={() => setControlsOpen((v) => !v)}
                  aria-expanded={controlsOpen}
                  className="flex items-center gap-2 text-left cursor-pointer min-w-0"
                >
                  {controlsOpen ? (
                    <ChevronDown className="w-4 h-4 text-slate-500 shrink-0" />
                  ) : (
                    <ChevronRight className="w-4 h-4 text-slate-500 shrink-0" />
                  )}
                  <h2 className="text-sm font-semibold text-slate-900">
                    {prefix === 'QT'
                      ? 'Master Job Controls'
                      : `${meta.shortName} Inputs`}
                  </h2>
                </button>
                <HelpTip
                  text={
                    prefix === 'QT'
                      ? 'Universal Master Job Controls held on Customer Quotation (QT). Changes entered here automatically propagate across Proforma Invoice, Commercial Invoice, Packing List, Delivery Note, and Excel workbooks.'
                      : `Fields specific to ${meta.shortName}. Shared client, RFQ/Tender ref, currency, and line-item controls live on the Customer Quotation (QT) master sheet.`
                  }
                />
              </div>

              <button
                type="button"
                onClick={() => setControlsOpen((v) => !v)}
                className="flex items-center gap-3 text-xs text-slate-500 min-w-0 cursor-pointer"
              >
                {!controlsOpen && (
                  <span className="truncate max-w-[320px] font-medium text-slate-600">
                    {controlsSummary}
                  </span>
                )}
                <span className="font-mono-tabular shrink-0">{meta.sheetName}</span>
              </button>
            </div>

            {controlsOpen && (
              <div className="px-5 pb-5 pt-2 border-t border-slate-100 space-y-4">
                {prefix !== 'QT' && (
                  <div className="flex flex-wrap items-center justify-between gap-2 rounded-md bg-slate-50 px-3.5 py-2 text-xs text-slate-600">
                    <span className="truncate">
                      Master Job: <strong className="text-slate-900">{job.jobName}</strong> ·{' '}
                      {endUser || 'Unassigned Client'} · {curr}
                    </span>
                    <button
                      type="button"
                      onClick={() => onNavigateTab('QT')}
                      className="inline-flex items-center gap-1 font-semibold text-blue-700 hover:underline cursor-pointer whitespace-nowrap"
                    >
                      <SlidersHorizontal className="w-3.5 h-3.5" />
                      <span>Edit Master Controls on Quotation</span>
                    </button>
                  </div>
                )}

                {prefix === 'PO' ? (
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                      <label className="flex items-center gap-1.5 text-xs font-medium text-slate-600 mb-1.5">
                        <span>Supplier Name</span>
                        <HelpTip text="Required before issuing Supplier PO." />
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
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                      <label className="flex items-center gap-1.5 text-xs font-medium text-slate-600 mb-1.5">
                        <span>Ship Date</span>
                        <HelpTip text="Required before issuing Packing List." />
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
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                      <label className="flex items-center gap-1.5 text-xs font-medium text-slate-600 mb-1.5">
                        <span>Delivery Date</span>
                        <HelpTip text="Required before issuing Delivery Note." />
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
                      <label className="flex items-center gap-1.5 text-xs font-medium text-slate-600 mb-1.5">
                        <span>Transport / BL / Truck Ref.</span>
                        <HelpTip text="Required before issuing Delivery Note." />
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
                        Received By
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
                ) : prefix === 'PI' ? (
                  /* Proforma Invoice-specific inputs (lightweight stage doc triggered by Quotation) */
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-medium text-slate-600 mb-1.5">
                        Client PO Ref.
                      </label>
                      <input
                        type="text"
                        disabled={isLocked}
                        value={job.controls.customerPoRef}
                        onChange={(e) => onUpdateControls({ customerPoRef: e.target.value })}
                        placeholder="Optional until client confirms PO"
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
                ) : prefix === 'CI' ? (
                  /* Commercial Invoice-specific inputs only (mirrors PI closely on print) */
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-medium text-slate-600 mb-1.5">
                        Client PO Ref.
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
                ) : (
                  /* QT: Universal Master Job Controls */
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                    <div className="lg:col-span-2">
                      <label className="block text-xs font-medium text-slate-600 mb-1.5">
                        Job Title / Reference
                      </label>
                      <input
                        type="text"
                        value={job.jobName}
                        onChange={(e) => onUpdateJobName(e.target.value)}
                        className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-medium text-slate-900 focus:border-slate-900 focus:outline-none"
                      />
                    </div>

                    <div className="lg:col-span-2">
                      <label className="flex items-center gap-1.5 text-xs font-medium text-slate-600 mb-1.5">
                        <span>TO — Client</span>
                        <HelpTip text={`Select from ${company.endUsers.length} directory clients. Address populates automatically (locked) and ATTN clears for fresh entry. Choose OTHER for a custom one-off client.`} />
                      </label>
                      <select
                        disabled={isLocked}
                        value={job.controls.endUserSelect}
                        onChange={(e) => {
                          const selected = e.target.value;
                          const savedAddress =
                            selected === 'OTHER'
                              ? ''
                              : company.endUserAddresses?.[selected] || '';
                          onUpdateControls({
                            endUserSelect: selected,
                            endUserAddress: savedAddress,
                            endUserContact: '',
                          });
                        }}
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
                          Custom Client Name
                        </label>
                        <input
                          type="text"
                          disabled={isLocked}
                          value={job.controls.customEndUser}
                          onChange={(e) => onUpdateControls({ customEndUser: e.target.value })}
                          placeholder="Type client organization name..."
                          className="w-full min-h-[42px] rounded-lg border border-slate-900 px-3 py-2 text-sm text-blue-700 focus:outline-none"
                        />
                      </div>
                    )}

                    <div className="lg:col-span-2">
                      <label className="block text-xs font-medium text-slate-600 mb-1.5">
                        Contact / ATTN
                      </label>
                      <input
                        type="text"
                        disabled={isLocked}
                        value={job.controls.endUserContact}
                        onChange={(e) => onUpdateControls({ endUserContact: e.target.value })}
                        placeholder="e.g. Procurement Department"
                        className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                      />
                    </div>

                    <div className="lg:col-span-2">
                      <label className="flex items-center gap-1.5 text-xs font-medium text-slate-600 mb-1.5">
                        <span>Address</span>
                        {job.controls.endUserSelect !== 'OTHER' && (
                          <span className="inline-flex items-center gap-1 text-[11px] font-medium text-slate-400">
                            <Lock className="w-3 h-3" />
                            <span>From Directory</span>
                          </span>
                        )}
                        <HelpTip
                          text={
                            job.controls.endUserSelect !== 'OTHER'
                              ? 'Locked here because this client is managed in Clients & Settings → End-User Directory. Edit saved addresses there, or choose OTHER for a one-off custom client.'
                              : 'Enter the address for this custom client.'
                          }
                        />
                      </label>
                      <input
                        type="text"
                        readOnly={job.controls.endUserSelect !== 'OTHER'}
                        disabled={isLocked || job.controls.endUserSelect !== 'OTHER'}
                        value={
                          job.controls.endUserSelect !== 'OTHER'
                            ? (company.endUserAddresses?.[job.controls.endUserSelect] ??
                                job.controls.endUserAddress)
                            : job.controls.endUserAddress
                        }
                        onChange={(e) => {
                          if (job.controls.endUserSelect === 'OTHER') {
                            onUpdateControls({ endUserAddress: e.target.value });
                          }
                        }}
                        placeholder={
                          job.controls.endUserSelect !== 'OTHER'
                            ? 'Set in Clients & Settings → End-User Directory'
                            : 'e.g. Asmara, Eritrea'
                        }
                        className={`w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm focus:border-slate-900 focus:outline-none ${
                          job.controls.endUserSelect !== 'OTHER'
                            ? 'bg-slate-50 text-slate-600 cursor-not-allowed'
                            : 'text-blue-700 disabled:bg-slate-50'
                        }`}
                      />
                    </div>

                    {/* Inquiry Type (RFQ vs Tender) for Dashboard Separate Counters */}
                    <div>
                      <label className="flex items-center gap-1.5 text-xs font-medium text-slate-600 mb-1.5">
                        <span>Inquiry Type</span>
                        <HelpTip text="Classifies this job as an RFQ or a formal Tender for the Job Tracker Dashboard counters." />
                      </label>
                      <div className="grid grid-cols-2 gap-1 p-1 rounded-lg bg-slate-100 min-h-[42px]">
                        {(['RFQ', 'TENDER'] as InquiryType[]).map((t) => {
                          const selected = (job.inquiryType || 'RFQ') === t;
                          return (
                            <button
                              key={t}
                              type="button"
                              disabled={isLocked}
                              onClick={() => onUpdateInquiryType(t)}
                              className={`rounded-md text-xs font-semibold transition-colors cursor-pointer disabled:opacity-60 ${
                                selected
                                  ? 'text-white shadow-sm'
                                  : 'text-slate-600 hover:text-slate-900'
                              }`}
                              style={selected ? { backgroundColor: theme.primary } : undefined}
                            >
                              {t === 'RFQ' ? 'RFQ' : 'Tender'}
                            </button>
                          );
                        })}
                      </div>
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
                        placeholder="e.g. RFQ-2026-041 or TND-2026-09"
                        className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono-tabular text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-medium text-slate-600 mb-1.5">
                        Currency
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
                        Incoterm
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

                    <div>
                      <label className="block text-xs font-medium text-slate-600 mb-1.5">
                        Client PO Ref.
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
              </div>
            )}
          </div>

          {/* Collapsible 2: Job Remarks & Signature Lines (One-Click Expand/Collapse) */}
          <div className="rounded-lg border border-slate-200 bg-white overflow-visible">
            <div className="flex flex-wrap items-center justify-between gap-2 px-5 py-3 hover:bg-slate-50/70 transition-colors rounded-t-lg">
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setRemarksOpen((v) => !v)}
                  aria-expanded={remarksOpen}
                  className="flex items-center gap-2 text-left cursor-pointer"
                >
                  {remarksOpen ? (
                    <ChevronDown className="w-4 h-4 text-slate-500 shrink-0" />
                  ) : (
                    <ChevronRight className="w-4 h-4 text-slate-500 shrink-0" />
                  )}
                  <span className="text-xs font-semibold text-slate-900">
                    Job Remarks &amp; Signature Lines
                  </span>
                </button>
                <HelpTip text="Editable per job. Enter custom job remarks or specific signer/receiver names, or clear any field to leave it blank on printed documents and Excel exports." />
              </div>
              <button
                type="button"
                onClick={() => setRemarksOpen((v) => !v)}
                className="text-xs text-slate-500 cursor-pointer"
              >
                {remarksFilledCount === 0 ? 'Blank' : `${remarksFilledCount} / 3 populated`}
              </button>
            </div>

            {remarksOpen && (
              <div className="px-5 pb-5 pt-3 border-t border-slate-100 space-y-3">
                {!isLocked && remarksFilledCount > 0 && (
                  <div className="flex justify-end">
                    <button
                      type="button"
                      onClick={() =>
                        onRequestConfirm({
                          title: 'Clear Remarks & Signatures?',
                          message:
                            'This will clear the job remarks and both signature labels across all documents for this job.',
                          confirmLabel: 'Clear Fields',
                          destructive: true,
                          onConfirm: () =>
                            onUpdateControls({
                              remarks: '',
                              authorizedSignatory: '',
                              counterpartySignatory: '',
                            }),
                        })
                      }
                      className="text-[11px] font-medium text-slate-500 hover:text-red-600 cursor-pointer"
                    >
                      Clear Remarks &amp; Signatures
                    </button>
                  </div>
                )}

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="md:col-span-2">
                    <label className="flex items-center gap-1.5 text-xs font-medium text-slate-600 mb-1.5">
                      <span>Job Remarks / Special Notes</span>
                      <HelpTip text="Printed above the signature lines. Leave empty to hide the Remarks line." />
                    </label>
                    <textarea
                      rows={2}
                      disabled={isLocked}
                      value={job.controls.remarks}
                      onChange={(e) => onUpdateControls({ remarks: e.target.value })}
                      placeholder="Optional remarks or leave blank..."
                      className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                    />
                  </div>
                  <div>
                    <label className="flex items-center gap-1.5 text-xs font-medium text-slate-600 mb-1.5">
                      <span>Left Signature Label</span>
                      <HelpTip text="Usually the Grand Group authorized signatory name or title. Leave empty to omit." />
                    </label>
                    <input
                      type="text"
                      disabled={isLocked}
                      value={job.controls.authorizedSignatory ?? ''}
                      onChange={(e) => onUpdateControls({ authorizedSignatory: e.target.value })}
                      placeholder={`e.g. Authorized Signature — ${company.name}`}
                      className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                    />
                  </div>
                  <div>
                    <label className="flex items-center gap-1.5 text-xs font-medium text-slate-600 mb-1.5">
                      <span>Right Signature Label</span>
                      <HelpTip text="Customer, supplier, or storekeeper receiver signature label. Leave empty to omit." />
                    </label>
                    <input
                      type="text"
                      disabled={isLocked}
                      value={job.controls.counterpartySignatory ?? ''}
                      onChange={(e) => onUpdateControls({ counterpartySignatory: e.target.value })}
                      placeholder="e.g. Received in Good Order — Name & Stamp"
                      className="w-full min-h-[42px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-blue-700 focus:border-slate-900 focus:outline-none disabled:bg-slate-50"
                    />
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Collapsible 3: Extra Order Columns (Part No., CAS No., etc.) */}
          <div className="rounded-lg border border-slate-200 bg-white overflow-visible">
            <div className="flex flex-wrap items-center justify-between gap-2 px-5 py-3 hover:bg-slate-50/70 transition-colors rounded-t-lg">
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setExtraColsOpen((v) => !v)}
                  aria-expanded={extraColsOpen}
                  className="flex items-center gap-2 text-left cursor-pointer"
                >
                  {extraColsOpen ? (
                    <ChevronDown className="w-4 h-4 text-slate-500 shrink-0" />
                  ) : (
                    <ChevronRight className="w-4 h-4 text-slate-500 shrink-0" />
                  )}
                  <span className="text-xs font-semibold text-slate-900">
                    Extra Order Columns
                  </span>
                </button>
                <HelpTip text="Add up to 4 order-specific columns (Part No., CAS No., HS Code, Brand) reflected across all 6 documents & Excel exports with zero formula impact. You can also move ANY column left or right directly in the table header below." />
              </div>
              <button
                type="button"
                onClick={() => setExtraColsOpen((v) => !v)}
                className="text-xs text-slate-500 cursor-pointer"
              >
                {customColumns.length === 0
                  ? 'None added'
                  : customColumns.map((c) => c.label).join(' · ')}
              </button>
            </div>

            {extraColsOpen && (
              <div className="px-5 pb-4 pt-3 border-t border-slate-100 space-y-3">
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

                {!isLocked && (
                  <div className="flex flex-wrap items-center gap-2">
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
                        placeholder="Other column..."
                        className="min-h-[32px] w-40 rounded-md border border-slate-200 px-2.5 py-1 text-xs text-slate-900 focus:border-slate-900 focus:outline-none"
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
            )}
          </div>

          {/* Line Items Entry: iPad Touch Cards vs Windows Desktop Grid */}
          {platformMode === 'ipad' && (prefix === 'QT' || isLogistics) ? (
            <div className="space-y-3">
              <div className="flex items-center justify-between px-1">
                <h3 className="text-sm font-semibold text-slate-900">
                  {isLogistics
                    ? `${meta.shortName} Carton & Weight Cards`
                    : 'Single-Entry Item Cards'}
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
                theme={theme}
                disabled={isLocked}
                onUpdateItem={onUpdateItem}
                onRequestConfirm={onRequestConfirm}
              />
            </div>
          ) : (
            <div className="rounded-lg border border-slate-200 bg-white overflow-hidden">
              <div className="flex flex-wrap items-center justify-between gap-2 px-5 py-3.5 border-b border-slate-200 bg-slate-50/70">
                <div className="flex items-center gap-2">
                  <h3 className="text-sm font-semibold text-slate-900">
                    {isLogistics
                      ? `${meta.shortName} Carton & Weight Schedule`
                      : isPO
                      ? 'Supplier PO Line Items'
                      : 'Single-Entry Item Schedule'}
                  </h3>
                  <HelpTip
                    text={
                      isLogistics
                        ? 'Verify carton numbers and net/gross weights independently before issuing. Fixed Excel rows 17–36.'
                        : 'Single-entry master rows (Excel rows 17–36). Requested = end-user wording; Offered = supplier wording we sell; Unit Price links from PRICE CALC Col R. Use ← / → arrows or drag column headers to reorder any column across all documents.'
                    }
                  />
                </div>
                <div className="flex items-center gap-3">
                  {hasCustomOrder && !isLocked && (
                    <button
                      type="button"
                      onClick={() => onUpdateColumnOrder(undefined)}
                      className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1 text-[11px] font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-50 cursor-pointer"
                    >
                      <RotateCcw className="w-3 h-3" />
                      <span>Reset Column Order</span>
                    </button>
                  )}
                  <span className="text-xs text-slate-600 font-mono-tabular">
                    {calc.activeRowCount} / 20 Rows
                  </span>
                </div>
              </div>

              {rowSel.count > 0 && !isLocked && (
                <div className="no-print flex flex-wrap items-center gap-2 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-xs">
                  <span className="font-semibold text-slate-900">{rowSel.count} row{rowSel.count === 1 ? '' : 's'} selected</span>
                  <span className="text-slate-300">|</span>
                  <select
                    value={bulkUom}
                    onChange={(e) => setBulkUom(e.target.value)}
                    className="min-h-[32px] rounded border border-slate-200 bg-white px-2"
                  >
                    <option value="">Set UoM…</option>
                    {UOM_OPTIONS.map((u) => (
                      <option key={u} value={u}>{u}</option>
                    ))}
                  </select>
                  <button
                    type="button"
                    disabled={!bulkUom}
                    onClick={() => { applyBulkRows({ uom: bulkUom }); setBulkUom(''); }}
                    className="rounded border border-slate-200 bg-white px-2.5 py-1.5 font-medium text-slate-800 hover:bg-slate-100 disabled:opacity-50 cursor-pointer"
                  >
                    Apply UoM
                  </button>
                  <button
                    type="button"
                    onClick={() => applyBulkRows({ qty: null })}
                    className="rounded border border-slate-200 bg-white px-2.5 py-1.5 font-medium text-slate-800 hover:bg-slate-100 cursor-pointer"
                  >
                    Clear qty
                  </button>
                  <button
                    type="button"
                    onClick={() => applyBulkRows({ requestedDescription: '', offeredDescription: '', uom: '', qty: null })}
                    className="rounded border border-red-200 bg-white px-2.5 py-1.5 font-medium text-red-700 hover:bg-red-50 cursor-pointer"
                  >
                    Clear selected rows
                  </button>
                  <PasteIntoRows
                    disabled={isLocked}
                    rowIds={rowSel.selected}
                    fields={[
                      { key: 'offeredDescription', label: 'Offered description', parse: (raw: string) => raw },
                      { key: 'requestedDescription', label: 'Requested description', parse: (raw: string) => raw },
                      { key: 'uom', label: 'UoM', parse: (raw: string) => raw },
                      { key: 'qty', label: 'Qty', parse: (raw: string) => parseFlexibleNumber(raw) ?? undefined },
                    ]}
                    onApply={(key, pairs) => pairs.forEach((p) => onUpdateItem(p.rowId, { [key]: p.value } as Partial<JobItemRow>))}
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
                          disabled={selectableRowIds.length === 0 || isLocked}
                          checked={rowSel.allChecked(selectableRowIds)}
                          onChange={() => rowSel.toggleAll(selectableRowIds)}
                        />
                      </th>
                      <th className="py-2.5 px-2.5 text-center font-semibold w-12">Row</th>
                      {orderedColKeys.map((colKey, colIdx) => {
                        const colMeta = getColumnHeaderMeta(colKey, false);
                        const alignClass =
                          colMeta.align === 'center'
                            ? 'text-center justify-center'
                            : colMeta.align === 'right'
                            ? 'text-right justify-end'
                            : 'text-left justify-start';

                        return (
                          <th
                            key={colKey}
                            draggable={!isLocked}
                            onDragStart={() => setDragColKey(colKey)}
                            onDragOver={(e) => {
                              if (!isLocked) e.preventDefault();
                            }}
                            onDrop={() => {
                              if (dragColKey) handleDropColumn(dragColKey, colKey);
                              setDragColKey(null);
                            }}
                            onDragEnd={() => setDragColKey(null)}
                            className={`py-2 px-2.5 font-semibold select-none ${colMeta.widthClass} ${
                              !isLocked ? 'cursor-grab active:cursor-grabbing' : ''
                            } ${dragColKey === colKey ? 'opacity-60' : ''}`}
                          >
                            <div className={`flex items-center gap-1.5 ${alignClass}`}>
                              {!isLocked && colIdx > 0 && (
                                <button
                                  type="button"
                                  onClick={() => handleMoveColumnByOffset(colKey, -1)}
                                  title={`Move ${colMeta.label} left`}
                                  aria-label={`Move ${colMeta.label} left`}
                                  className="p-0.5 rounded bg-white/15 hover:bg-white/30 transition-colors cursor-pointer shrink-0"
                                >
                                  <ArrowLeft className="w-3 h-3" />
                                </button>
                              )}
                              <span className="truncate">{colMeta.label}</span>
                              {!isLocked && colIdx < orderedColKeys.length - 1 && (
                                <button
                                  type="button"
                                  onClick={() => handleMoveColumnByOffset(colKey, 1)}
                                  title={`Move ${colMeta.label} right`}
                                  aria-label={`Move ${colMeta.label} right`}
                                  className="p-0.5 rounded bg-white/15 hover:bg-white/30 transition-colors cursor-pointer shrink-0"
                                >
                                  <ArrowRight className="w-3 h-3" />
                                </button>
                              )}
                            </div>
                          </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {calc.rows.map((row, idx) => {
                      const raw = job.items[idx];
                      const canEditBase = prefix === 'QT' && !isLocked;
                      const canEditLogistics = isLogistics && !isLocked && !row.isEmpty;

                      const renderEditCell = (colKey: string) => {
                        if (colKey.startsWith('custom:')) {
                          const cc = customColMap.get(colKey);
                          if (!cc) return null;
                          return (
                            <td key={colKey} className="p-1">
                              {canEditBase ? (
                                <input
                                  type="text"
                                  value={raw.customFields?.[cc.id] || ''}
                                  placeholder={idx === 0 || !row.isEmpty ? `${cc.label}...` : ''}
                                  onChange={(e) =>
                                    onUpdateItem(row.rowNumber, {
                                      customFields: {
                                        ...(raw.customFields || {}),
                                        [cc.id]: e.target.value,
                                      },
                                    })
                                  }
                                  className="w-full min-h-[36px] rounded border border-slate-200/80 hover:border-slate-400 focus:border-slate-900 bg-white/70 focus:bg-white px-2 py-1 font-mono-tabular text-blue-700 focus:outline-none"
                                />
                              ) : (
                                <div className="px-2 py-1.5 font-mono-tabular text-slate-800">
                                  {row.customFields?.[cc.id] || ''}
                                </div>
                              )}
                            </td>
                          );
                        }

                        switch (colKey) {
                          case 'requested':
                            return (
                              <td key={colKey} className="p-1">
                                {canEditBase ? (
                                  <input
                                    type="text"
                                    value={raw.requestedDescription}
                                    placeholder={idx === 0 ? 'End user RFQ wording...' : ''}
                                    onChange={(e) =>
                                      onUpdateItem(row.rowNumber, {
                                        requestedDescription: e.target.value,
                                      })
                                    }
                                    className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-slate-900 focus:outline-none"
                                  />
                                ) : (
                                  <div className="px-2 py-1.5 text-slate-800">
                                    {row.requestedDescription}
                                  </div>
                                )}
                              </td>
                            );
                          case 'offered':
                            return (
                              <td key={colKey} className="p-1">
                                {canEditBase ? (
                                  <input
                                    type="text"
                                    value={raw.offeredDescription}
                                    placeholder={idx === 0 ? 'Supplier offered wording...' : ''}
                                    onChange={(e) =>
                                      onUpdateItem(row.rowNumber, {
                                        offeredDescription: e.target.value,
                                      })
                                    }
                                    className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-slate-900 focus:outline-none"
                                  />
                                ) : (
                                  <div className="px-2 py-1.5 text-slate-800">
                                    {row.offeredDescription}
                                  </div>
                                )}
                              </td>
                            );
                          case 'uom':
                            return (
                              <td key={colKey} className="p-1">
                                {canEditBase ? (
                                  <select
                                    value={raw.uom}
                                    onChange={(e) =>
                                      onUpdateItem(row.rowNumber, { uom: e.target.value })
                                    }
                                    className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-1.5 py-1 text-center font-mono-tabular text-blue-700 focus:outline-none"
                                  >
                                    <option value="">—</option>
                                    {UOM_OPTIONS.map((u) => (
                                      <option key={u} value={u}>{u}</option>
                                    ))}
                                  </select>
                                ) : (
                                  <div className="px-2 py-1.5 text-center font-mono-tabular text-slate-700">
                                    {row.uom}
                                  </div>
                                )}
                              </td>
                            );
                          case 'qty':
                            return (
                              <td key={colKey} className="p-1">
                                {canEditBase ? (
                                  <input
                                    type="number"
                                    min="0"
                                    step="any"
                                    value={raw.qty ?? ''}
                                    onChange={(e) =>
                                      onUpdateItem(row.rowNumber, {
                                        qty: parseOptionalNumber(e.target.value),
                                      })
                                    }
                                    className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-right font-mono-tabular text-blue-700 font-medium focus:outline-none"
                                  />
                                ) : (
                                  <div className="px-2 py-1.5 text-right font-mono-tabular text-slate-800">
                                    {row.qty ?? ''}
                                  </div>
                                )}
                              </td>
                            );
                          case 'packages':
                            return (
                              <td key={colKey} className="py-2 px-2 text-center font-mono-tabular text-slate-700">
                                {row.boxes ? `${row.boxes} pkg` : '—'}
                                {row.boxDims && <div className="text-[11px] text-slate-500">{row.boxDims} cm</div>}
                              </td>
                            );
                          case 'carton':
                            return (
                              <td key={colKey} className="p-1">
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
                            );
                          case 'net_wt':
                            return (
                              <td key={colKey} className="p-1">
                                <input
                                  type="number"
                                  step="0.1"
                                  min="0"
                                  disabled={!canEditLogistics}
                                  value={raw.netWtKg ?? ''}
                                  onChange={(e) =>
                                    onUpdateItem(row.rowNumber, {
                                      netWtKg: parseOptionalNumber(e.target.value),
                                    })
                                  }
                                  className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-right font-mono-tabular text-blue-700 focus:outline-none disabled:opacity-40"
                                />
                              </td>
                            );
                          case 'gross_wt':
                            return (
                              <td key={colKey} className="p-1">
                                <input
                                  type="number"
                                  step="0.1"
                                  min="0"
                                  disabled={!canEditLogistics}
                                  value={raw.grossWtKg ?? ''}
                                  onChange={(e) =>
                                    onUpdateItem(row.rowNumber, {
                                      grossWtKg: parseOptionalNumber(e.target.value),
                                    })
                                  }
                                  className="w-full min-h-[36px] rounded border border-transparent hover:border-slate-300 focus:border-slate-900 bg-transparent focus:bg-white px-2 py-1 text-right font-mono-tabular text-blue-700 focus:outline-none disabled:opacity-40"
                                />
                              </td>
                            );
                          case 'unit_val':
                            return (
                              <td
                                key={colKey}
                                className="py-1.5 px-3 text-right font-mono-tabular text-slate-700"
                              >
                                {formatMoney(isPO ? row.supplierUnitCost : row.activeUnitPrice)}
                              </td>
                            );
                          case 'total_val':
                            return (
                              <td
                                key={colKey}
                                className="py-1.5 px-3 text-right font-mono-tabular font-medium text-slate-900"
                              >
                                {formatMoney(isPO ? row.supplierTotalCost : row.activeTotalPrice)}
                              </td>
                            );
                          default:
                            return <td key={colKey} />;
                        }
                      };

                      return (
                        <tr
                          key={row.rowNumber}
                          className="hover:bg-slate-50/90 transition-colors"
                          style={idx % 2 === 1 ? { backgroundColor: theme.zebraRow } : undefined}
                        >
                          <td className="py-1.5 px-2 text-center">
                            <input
                              type="checkbox"
                              aria-label={`Select row ${row.rowNumber}`}
                              disabled={row.isEmpty || isLocked}
                              checked={rowSel.has(row.rowNumber)}
                              onChange={() => rowSel.toggle(row.rowNumber)}
                            />
                          </td>
                          <td className="py-1.5 px-2.5 text-center font-mono-tabular text-slate-400">
                            {row.rowNumber}
                          </td>
                          {orderedColKeys.map(renderEditCell)}
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
          <div
            className="a4-sheet w-full max-w-[960px] rounded-lg border border-slate-200 bg-white shadow-sm overflow-hidden"
            style={{ fontFamily: theme.fontCss, color: theme.bodyText }}
          >
            {/* Top Brand & Document Title Banner */}
            <div
              className="px-8 py-5 flex flex-wrap items-center justify-between gap-4"
              style={{ backgroundColor: theme.primary, color: theme.primaryText }}
            >
              <div className="flex items-center gap-4">
                {company.logoDataUrl && (
                  <div className="w-24 h-24 rounded-md bg-white p-2 flex items-center justify-center shrink-0">
                    <img
                      src={company.logoDataUrl}
                      alt={company.name}
                      className="max-w-full max-h-full object-contain"
                    />
                  </div>
                )}
                <div>
                  <div className="text-2xl font-bold tracking-tight">
                    {company.name}
                  </div>
                  <div className="text-xs opacity-85 mt-0.5">
                    {company.tagline}
                  </div>
                </div>
              </div>
              <div className="text-right">
                <div className="text-xl font-bold tracking-wide">
                  {meta.title}
                </div>
                {issuedDocNo && (
                  <div className="text-xs font-mono-tabular opacity-90 mt-1">
                    {issuedDocNo}
                  </div>
                )}
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
                  <div className="text-sm font-bold">
                    {isPO ? job.controls.supplierName || '—' : endUser || '—'}
                  </div>
                  {isPO ? (
                    <div className="opacity-80">{job.controls.supplierAddress || '—'}</div>
                  ) : (
                    <>
                      {job.controls.endUserContact && (
                        <div className="opacity-85">ATTN: {job.controls.endUserContact}</div>
                      )}
                      {(job.controls.endUserSelect !== 'OTHER'
                        ? (company.endUserAddresses?.[job.controls.endUserSelect] ??
                            job.controls.endUserAddress)
                        : job.controls.endUserAddress) && (
                        <div className="opacity-80">
                          {job.controls.endUserSelect !== 'OTHER'
                            ? (company.endUserAddresses?.[job.controls.endUserSelect] ??
                                job.controls.endUserAddress)
                            : job.controls.endUserAddress}
                        </div>
                      )}
                    </>
                  )}
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
                          {job.issuedDocs.PI?.docNumber || '—'}
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

              {/* Strict Document Table (Follows exact columnOrder across all 6 documents) */}
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
                      {orderedColKeys.map((colKey) => {
                        const colMeta = getColumnHeaderMeta(colKey, true);
                        const alignCls =
                          colMeta.align === 'center'
                            ? 'text-center'
                            : colMeta.align === 'right'
                            ? 'text-right'
                            : 'text-left';
                        return (
                          <th
                            key={colKey}
                            className={`py-3 px-2.5 font-semibold ${alignCls} ${colMeta.widthClass}`}
                          >
                            {colMeta.label}
                          </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody>
                    {activeRows.length === 0 ? (
                      <tr>
                        <td
                          colSpan={1 + orderedColKeys.length}
                          className="py-10 text-center text-slate-400 border-b border-slate-200"
                        >
                          No items entered yet. Enter items on Internal Quote or Quotation (Master Job Controls) to populate rows 17–36.
                        </td>
                      </tr>
                    ) : (
                      activeRows.map((row, idx) => (
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
                          {orderedColKeys.map((colKey) => {
                            if (colKey.startsWith('custom:')) {
                              const cc = customColMap.get(colKey);
                              return (
                                <td
                                  key={colKey}
                                  className="py-2.5 px-2.5 font-mono-tabular whitespace-nowrap"
                                >
                                  {(cc && row.customFields?.[cc.id]) || '—'}
                                </td>
                              );
                            }
                            switch (colKey) {
                              case 'requested':
                                return (
                                  <td key={colKey} className="py-2.5 px-3">
                                    {row.requestedDescription}
                                  </td>
                                );
                              case 'offered':
                                return (
                                  <td key={colKey} className="py-2.5 px-3 font-medium">
                                    {row.offeredDescription}
                                  </td>
                                );
                              case 'uom':
                                return (
                                  <td
                                    key={colKey}
                                    className="py-2.5 px-2 text-center font-mono-tabular"
                                  >
                                    {row.uom}
                                  </td>
                                );
                              case 'qty':
                                return (
                                  <td
                                    key={colKey}
                                    className="py-2.5 px-2 text-right font-mono-tabular"
                                  >
                                    {row.qty}
                                  </td>
                                );
                              case 'packages':
                                return (
                                  <td key={colKey} className="py-2.5 px-2 text-center font-mono-tabular">
                                    {row.boxes ? `${row.boxes} pkg` : '—'}
                                    {row.boxDims && <div className="text-[11px] text-slate-500">{row.boxDims} cm</div>}
                                  </td>
                                );
                              case 'carton':
                                return (
                                  <td
                                    key={colKey}
                                    className="py-2.5 px-2 text-center font-mono-tabular"
                                  >
                                    {row.cartonNo || '—'}
                                  </td>
                                );
                              case 'net_wt':
                                return (
                                  <td
                                    key={colKey}
                                    className="py-2.5 px-2.5 text-right font-mono-tabular"
                                  >
                                    {formatNumber(row.netWtKg, 2)}
                                  </td>
                                );
                              case 'gross_wt':
                                return (
                                  <td
                                    key={colKey}
                                    className="py-2.5 px-2.5 text-right font-mono-tabular font-medium"
                                  >
                                    {formatNumber(row.grossWtKg, 2)}
                                  </td>
                                );
                              case 'unit_val':
                                return (
                                  <td
                                    key={colKey}
                                    className="py-2.5 px-2.5 text-right font-mono-tabular"
                                  >
                                    {formatMoney(isPO ? row.supplierUnitCost : row.activeUnitPrice)}
                                  </td>
                                );
                              case 'total_val':
                                return (
                                  <td
                                    key={colKey}
                                    className="py-2.5 px-2.5 text-right font-mono-tabular font-medium"
                                  >
                                    {formatMoney(isPO ? row.supplierTotalCost : row.activeTotalPrice)}
                                  </td>
                                );
                              default:
                                return <td key={colKey} />;
                            }
                          })}
                        </tr>
                      ))
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
                      <span className="text-slate-600">Total Line Quantity</span>
                      <span className="font-semibold text-slate-900">{calc.totalQty}</span>
                    </div>
                    <div className="flex justify-between">
                      <span className="text-slate-600">Total Net Weight</span>
                      <span className="font-semibold text-slate-900">
                        {formatNumber(calc.plTotalNetWt, 2)} KG
                      </span>
                    </div>
                    <div className="flex justify-between pt-1.5 border-t border-slate-300/60 text-sm font-bold text-slate-900">
                      <span>Total Gross Weight</span>
                      <span>{formatNumber(calc.plTotalGrossWt, 2)} KG</span>
                    </div>
                  </div>
                ) : isPO ? (
                  <div
                    className="w-full sm:w-80 rounded-md p-4 text-sm font-mono-tabular border"
                    style={{ backgroundColor: theme.totalBg, borderColor: theme.strongBorder }}
                  >
                    <div className="flex justify-between font-bold text-slate-900">
                      <span>TOTAL ORDER COST ({curr})</span>
                      <span>{formatMoney(calc.totalSupplierCost)}</span>
                    </div>
                  </div>
                ) : (
                  <div className="w-full sm:w-96 text-xs font-mono-tabular">
                    <div className="flex justify-between py-2 border-b border-slate-200">
                      <span className="text-slate-600">GOODS SUBTOTAL</span>
                      <span className="font-medium text-slate-900">
                        {formatMoney(calc.goodsSubtotal, curr)}
                      </span>
                    </div>
                    <div className="flex justify-between py-2 border-b border-slate-200">
                      <span className="text-slate-600">
                        FREIGHT &amp; DOCUMENTATION · {calc.freightMode === 'sea' ? 'SEA' : 'AIR'} FREIGHT ({job.controls.incoterm})
                      </span>
                      <span className="font-medium text-slate-900">
                        {formatMoney(calc.customerFreight, curr)}
                      </span>
                    </div>
                    <div
                      className="flex justify-between py-3 px-3 mt-1 rounded-sm text-sm font-bold border-t border-b-2"
                      style={{
                        backgroundColor: theme.totalBg,
                        borderColor: theme.strongBorder,
                        color: theme.primary,
                      }}
                    >
                      <span>GRAND TOTAL ({curr})</span>
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
