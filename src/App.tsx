import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  Calculator,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Hash,
  Menu,
  Monitor,
  Package,
  Palette,
  Plus,
  Receipt,
  ShoppingCart,
  Tablet,
  Trash2,
  Truck,
  X,
} from 'lucide-react';
import { DOC_META, EXCEL_THEMES } from './constants/themes';
import { DocumentWorkspaceView } from './components/DocumentWorkspaceView';
import { ExcelStudioView } from './components/ExcelStudioView';
import { OfflineIndicator, PWAInstallButton } from './components/PWAInstallButton';
import { PriceCalcView } from './components/PriceCalcView';
import { RegistryView } from './components/RegistryView';
import { QuoteApplyOptions, SupplierQuoteImportModal } from './components/SupplierQuoteImportModal';
import {
  createBlankJob,
  createDefaultCompanyProfile,
  formatDocNumber,
  getNextSequenceNumber,
  loadActiveJobId,
  loadCompanyProfile,
  loadJobs,
  loadRegistry,
  loadThemeVariant,
  normalizeJob,
  saveActiveJobId,
  saveCompanyProfile,
  saveJobs,
  saveRegistry,
  saveThemeVariant,
} from './store/storage';
import {
  CompanyProfile,
  CustomColumnDef,
  DocPrefix,
  FreightConfig,
  JobControls,
  JobItemRow,
  PlatformMode,
  ProcurementJob,
  RegistryEntry,
  ThemeVariantId,
  WorkspaceTab,
} from './types/procurement';
import { calculateJob, formatMoney, formatPct, getResolvedEndUser } from './utils/calculations';
import { importWorkbookFromXlsx } from './utils/excelEngine';
import { detectDefaultPlatform, savePlatformMode } from './utils/platform';
import { ParsedSupplierQuote, parseSupplierPdfFile } from './utils/supplierQuoteParser';

const TAB_LABELS: Record<WorkspaceTab, string> = {
  PI: 'Proforma Invoice (Entry)',
  PRICE_CALC: 'Price Calc (Internal)',
  QT: 'Quotation (QT)',
  PO: 'Supplier PO (PO)',
  PL: 'Packing List (PL)',
  DN: 'Delivery Note (DN)',
  CI: 'Commercial Invoice (CI)',
  REGISTRY: 'Master Registry',
  STUDIO: 'Excel Studio & Import',
};

export default function App() {
  const [jobs, setJobs] = useState<ProcurementJob[]>(() => loadJobs());
  const [activeJobId, setActiveJobId] = useState<string>(() => loadActiveJobId(loadJobs()));
  const [registry, setRegistry] = useState<RegistryEntry[]>(() => loadRegistry());
  const [themeId, setThemeId] = useState<ThemeVariantId>(() => loadThemeVariant());
  const [company, setCompany] = useState<CompanyProfile>(() => loadCompanyProfile());
  const [platformMode, setPlatformMode] = useState<PlatformMode>(() => detectDefaultPlatform());
  const [activeTab, setActiveTab] = useState<WorkspaceTab>('PI');
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [toast, setToast] = useState<string | null>(null);
  const [pendingSupplierQuote, setPendingSupplierQuote] = useState<ParsedSupplierQuote | null>(null);
  const quickImportRef = useRef<HTMLInputElement | null>(null);

  const activeJob = useMemo(
    () => jobs.find((j) => j.id === activeJobId) || jobs[0],
    [jobs, activeJobId]
  );

  const calc = useMemo(() => calculateJob(activeJob), [activeJob]);
  const nextSeq = useMemo(() => getNextSequenceNumber(registry), [registry]);
  const theme = EXCEL_THEMES[themeId];

  useEffect(() => saveJobs(jobs), [jobs]);
  useEffect(() => saveRegistry(registry), [registry]);
  useEffect(() => saveThemeVariant(themeId), [themeId]);
  useEffect(() => saveCompanyProfile(company), [company]);
  useEffect(() => savePlatformMode(platformMode), [platformMode]);
  useEffect(() => {
    if (activeJob?.id) saveActiveJobId(activeJob.id);
  }, [activeJob?.id]);

  const showToast = (msg: string) => {
    setToast(msg);
  };

  useEffect(() => {
    if (!toast) return;
    const timer = setTimeout(() => setToast(null), 4000);
    return () => clearTimeout(timer);
  }, [toast]);

  const mutateActiveJob = (updater: (job: ProcurementJob) => ProcurementJob) => {
    setJobs((prev) =>
      prev.map((j) =>
        j.id === activeJob.id
          ? { ...updater(j), updatedAt: new Date().toISOString() }
          : j
      )
    );
  };

  const handleUpdateJobName = (name: string) => {
    mutateActiveJob((j) => ({ ...j, jobName: name }));
  };

  const handleUpdateControls = (patch: Partial<JobControls>) => {
    mutateActiveJob((j) => ({
      ...j,
      controls: { ...j.controls, ...patch },
    }));
  };

  const handleUpdateFreight = (patch: Partial<FreightConfig>) => {
    mutateActiveJob((j) => ({
      ...j,
      freight: { ...j.freight, ...patch },
    }));
  };

  const handleUpdateCustomColumns = (cols: CustomColumnDef[]) => {
    mutateActiveJob((j) => ({
      ...j,
      customColumns: cols,
    }));
  };

  const handleUpdateItem = (rowNumber: number, patch: Partial<JobItemRow>) => {
    mutateActiveJob((j) => ({
      ...j,
      items: j.items.map((item) =>
        item.rowNumber === rowNumber ? { ...item, ...patch } : item
      ),
    }));
  };

  const handleCreateNewJob = () => {
    const num = jobs.length + 1;
    const fresh = createBlankJob(`Procurement Job #${num}`);
    setJobs((prev) => [fresh, ...prev]);
    setActiveJobId(fresh.id);
    setActiveTab('PI');
    setSidebarOpen(false);
    showToast(`Created fresh job from template: ${fresh.jobName}`);
  };

  const handleDuplicateAsNewJob = () => {
    const copy: ProcurementJob = {
      ...structuredClone(activeJob),
      id: `job-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      jobName: `${activeJob.jobName} (Revision)`,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      issuedDocs: {},
    };
    setJobs((prev) => [copy, ...prev]);
    setActiveJobId(copy.id);
    setActiveTab('PI');
    showToast(`Created unlocked revision job: ${copy.jobName}`);
  };

  const handleDeleteCurrentJob = () => {
    if (jobs.length <= 1) return;
    const remaining = jobs.filter((j) => j.id !== activeJob.id);
    setJobs(remaining);
    setActiveJobId(remaining[0].id);
    showToast(`Removed job "${activeJob.jobName}" from local workspace.`);
  };

  const handleIssueDocument = (prefix: DocPrefix) => {
    const seq = getNextSequenceNumber(registry);
    const year = new Date().getFullYear();
    const docNumber = formatDocNumber(prefix, seq, year);
    const now = new Date().toISOString();

    const partyName =
      prefix === 'PO'
        ? activeJob.controls.supplierName || 'Supplier'
        : getResolvedEndUser(activeJob) || 'Customer';
    const docAmount =
      prefix === 'PO'
        ? calc.totalSupplierCost
        : prefix === 'PL' || prefix === 'DN'
        ? 0
        : calc.grandTotal;

    const newRegEntry: RegistryEntry = {
      sequence: seq,
      docNumber,
      prefix,
      year,
      jobId: activeJob.id,
      jobName: activeJob.jobName,
      customer: partyName,
      currency: activeJob.controls.currency,
      amount: docAmount,
      issuedBy: activeJob.preparedBy || 'Procurement Desk',
      issuedAt: now,
      status: 'ISSUED',
    };

    setRegistry((prev) => [...prev, newRegEntry]);
    mutateActiveJob((j) => ({
      ...j,
      issuedDocs: {
        ...j.issuedDocs,
        [prefix]: {
          prefix,
          docNumber,
          sequence: seq,
          issuedAt: now,
          issuedBy: j.preparedBy || 'Procurement Desk',
          frozenSubtotal: calc.goodsSubtotal,
          frozenVat: calc.vatAmount,
          frozenFreight: calc.customerFreight,
          frozenGrandTotal: calc.grandTotal,
        },
      },
    }));

    showToast(`Claimed ${docNumber} from Master Registry and froze document.`);
  };

  const handleVoidRegistryEntry = (sequence: number, reason: string) => {
    setRegistry((prev) =>
      prev.map((e) =>
        e.sequence === sequence ? { ...e, status: 'VOID', voidReason: reason } : e
      )
    );
    setJobs((prev) =>
      prev.map((j) => {
        const nextIssued = { ...j.issuedDocs };
        (Object.keys(nextIssued) as DocPrefix[]).forEach((p) => {
          if (nextIssued[p]?.sequence === sequence) {
            nextIssued[p] = { ...nextIssued[p]!, isVoided: true };
          }
        });
        return { ...j, issuedDocs: nextIssued };
      })
    );
    showToast(`Marked sequence #${sequence} as VOID. Number permanently retired.`);
  };

  const handleSyncRegistry = (incoming: RegistryEntry[]) => {
    setRegistry((prev) => {
      const bySeq = new Map<number, RegistryEntry>();
      prev.forEach((e) => bySeq.set(e.sequence, e));
      incoming.forEach((e) => bySeq.set(e.sequence, e));
      return Array.from(bySeq.values()).sort((a, b) => a.sequence - b.sequence);
    });
  };

  const handleImportJob = (imported: ProcurementJob) => {
    const normalized = normalizeJob(imported);
    setJobs((prev) => [normalized, ...prev]);
    setActiveJobId(normalized.id);
    setActiveTab('PI');
  };

  const handleApplySupplierQuote = (
    quote: ParsedSupplierQuote,
    options: QuoteApplyOptions
  ) => {
    const selectedItems = quote.items.filter((it) => it.selected).slice(0, 20);
    if (selectedItems.length === 0) return;

    const buildUpdatedJob = (base: ProcurementJob): ProcurementJob => {
      const existingCols = base.customColumns || [];
      let nextCols = [...existingCols];
      let codeColId: string | null = null;

      if (options.addCodeColumn) {
        const matchCol = existingCols.find(
          (c) => c.label.toLowerCase() === options.codeColumnLabel.toLowerCase()
        );
        if (matchCol) {
          codeColId = matchCol.id;
        } else if (nextCols.length < 4) {
          codeColId = `col_${options.codeColumnLabel.toLowerCase().replace(/[^a-z0-9]+/g, '_')}_${Date.now().toString(36).slice(-3)}`;
          nextCols.push({
            id: codeColId,
            label: options.codeColumnLabel,
            position: options.codeColumnPosition || 'after_no',
          });
        }
      }

      const nextItems = base.items.map((row, idx) => {
        const src = selectedItems[idx];
        if (!src) return row;
        const nextCustom = { ...(row.customFields || {}) };
        if (codeColId && src.code.trim()) {
          nextCustom[codeColId] = src.code.trim();
        }
        return {
          ...row,
          requestedDescription:
            options.overwriteDescriptions || !row.requestedDescription.trim()
              ? src.description
              : row.requestedDescription,
          offeredDescription: src.description,
          uom: src.uom || row.uom || 'PCS',
          qty: src.qty ?? row.qty,
          supplierUnitCost: src.unitPrice ?? row.supplierUnitCost,
          customFields: nextCustom,
        };
      });

      return {
        ...base,
        updatedAt: new Date().toISOString(),
        customColumns: nextCols,
        controls: {
          ...base.controls,
          supplierName: quote.supplierName || base.controls.supplierName,
          supplierRef: quote.supplierRef || base.controls.supplierRef,
          currency: quote.currency || base.controls.currency,
          incoterm: quote.incoterm || base.controls.incoterm,
          paymentTerms: quote.paymentTerms || base.controls.paymentTerms,
          deliveryLeadTime: quote.deliveryLeadTime || base.controls.deliveryLeadTime,
        },
        items: nextItems,
      };
    };

    if (options.target === 'new_job') {
      const fresh = buildUpdatedJob(
        createBlankJob(quote.supplierName ? `${quote.supplierName} Quote` : 'Supplier Quote Job')
      );
      setJobs((prev) => [fresh, ...prev]);
      setActiveJobId(fresh.id);
    } else {
      mutateActiveJob(buildUpdatedJob);
    }

    setPendingSupplierQuote(null);
    setActiveTab('PRICE_CALC');
    showToast(
      `Loaded ${selectedItems.length} supplier quote items into Price Calc (${quote.currency}).`
    );
  };

  const handleQuickExcelOpen = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const isPdf =
        file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
      if (isPdf) {
        const quote = await parseSupplierPdfFile(file);
        setPendingSupplierQuote(quote);
        return;
      }

      const res = await importWorkbookFromXlsx(file);
      if (res.kind === 'registry') {
        handleSyncRegistry(res.entries);
        setActiveTab('REGISTRY');
        showToast(`Synced ${res.importedCount} entries from ${file.name}`);
      } else if (res.kind === 'supplier_quote') {
        setPendingSupplierQuote(res.quote);
      } else {
        handleImportJob(res.job);
        showToast(`Opened ${file.name} (${res.itemCount} items loaded)`);
      }
    } catch {
      showToast('Could not read file. Upload a supplier .pdf / .xlsx quote or Grand Group .xlsx workbook.');
    } finally {
      if (quickImportRef.current) quickImportRef.current.value = '';
    }
  };

  const handleUpdateCompany = (patch: Partial<CompanyProfile>) => {
    setCompany((prev) => ({ ...prev, ...patch }));
  };

  const handleResetCompany = () => {
    setCompany(createDefaultCompanyProfile());
    showToast('Reset company branding, logo, and End-User list to defaults.');
  };

  const handleRestoreBackup = (
    restoredJobs: ProcurementJob[],
    restoredRegistry: RegistryEntry[],
    restoredCompany?: CompanyProfile
  ) => {
    if (restoredJobs.length > 0) {
      const normalized = restoredJobs.map(normalizeJob);
      setJobs(normalized);
      setActiveJobId(normalized[0].id);
    }
    if (restoredRegistry.length > 0) {
      setRegistry(restoredRegistry);
    }
    if (restoredCompany) {
      setCompany((prev) => ({ ...prev, ...restoredCompany }));
    }
  };

  const navItems: {
    group: string;
    items: { id: WorkspaceTab; label: string; icon: React.ReactNode; docPrefix?: DocPrefix }[];
  }[] = [
    {
      group: 'Entry & Costing',
      items: [
        { id: 'PI', label: 'Proforma Invoice', icon: <FileText className="w-4 h-4" />, docPrefix: 'PI' },
        { id: 'PRICE_CALC', label: 'Price Calc (Internal)', icon: <Calculator className="w-4 h-4" /> },
      ],
    },
    {
      group: 'Trade Documents',
      items: [
        { id: 'QT', label: 'Quotation', icon: <FileText className="w-4 h-4" />, docPrefix: 'QT' },
        { id: 'PO', label: 'Supplier PO', icon: <ShoppingCart className="w-4 h-4" />, docPrefix: 'PO' },
        { id: 'PL', label: 'Packing List', icon: <Package className="w-4 h-4" />, docPrefix: 'PL' },
        { id: 'DN', label: 'Delivery Note', icon: <Truck className="w-4 h-4" />, docPrefix: 'DN' },
        { id: 'CI', label: 'Commercial Invoice', icon: <Receipt className="w-4 h-4" />, docPrefix: 'CI' },
      ],
    },
    {
      group: 'Registry & Excel',
      items: [
        { id: 'REGISTRY', label: 'Master Registry', icon: <Hash className="w-4 h-4" /> },
        { id: 'STUDIO', label: 'Excel Studio & Import', icon: <Palette className="w-4 h-4" /> },
      ],
    },
  ];

  const isDocTab = (tab: WorkspaceTab): tab is DocPrefix =>
    ['QT', 'PI', 'PO', 'PL', 'DN', 'CI'].includes(tab);

  const endUser = getResolvedEndUser(activeJob) || 'Unassigned End User';

  return (
    <div className="min-h-screen flex bg-slate-50 text-slate-900">
      {/* Mobile / iPad Sidebar Backdrop */}
      {sidebarOpen && (
        <div
          className="no-print fixed inset-0 z-40 bg-slate-900/40 lg:hidden"
          onClick={() => setSidebarOpen(false)}
        />
      )}

      {/* Persistent Left Sidebar (260px Desktop, Slide-Over on iPad/Mobile) */}
      <aside
        className={`no-print fixed lg:sticky top-0 left-0 z-50 h-screen w-[264px] shrink-0 border-r border-slate-200 bg-white flex flex-col transition-transform duration-200 ${
          sidebarOpen ? 'translate-x-0' : '-translate-x-full lg:translate-x-0'
        }`}
      >
        {/* Brand Wordmark & Logo */}
        <div className="h-16 px-4 flex items-center justify-between border-b border-slate-200">
          <div className="flex items-center gap-2.5 min-w-0">
            <div className="w-9 h-9 rounded-md border border-slate-200 bg-white p-1 flex items-center justify-center shrink-0">
              <img
                src={company.logoDataUrl}
                alt={company.name}
                className="max-w-full max-h-full object-contain"
              />
            </div>
            <span className="font-display text-sm font-bold tracking-tight text-slate-900 truncate">
              {company.name}
            </span>
          </div>
          <button
            type="button"
            onClick={() => setSidebarOpen(false)}
            className="lg:hidden p-1.5 rounded-md text-slate-500 hover:bg-slate-100 cursor-pointer"
            aria-label="Close navigation"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Active Job Switcher */}
        <div className="p-4 border-b border-slate-200 space-y-2.5 bg-slate-50/60">
          <div className="flex items-center justify-between">
            <label htmlFor="job-select" className="text-xs font-medium text-slate-500">
              Active Job ({jobs.length})
            </label>
            {jobs.length > 1 && (
              <button
                type="button"
                onClick={handleDeleteCurrentJob}
                title="Delete current job"
                className="text-slate-400 hover:text-red-600 transition-colors cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
          <select
            id="job-select"
            value={activeJob.id}
            onChange={(e) => setActiveJobId(e.target.value)}
            className="w-full min-h-[40px] rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-xs font-medium text-slate-900 focus:border-slate-900 focus:outline-none"
          >
            {jobs.map((j) => (
              <option key={j.id} value={j.id}>
                {j.jobName}
              </option>
            ))}
          </select>
          <button
            type="button"
            onClick={handleCreateNewJob}
            className="w-full flex items-center justify-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs font-medium text-slate-700 hover:bg-slate-100 hover:text-slate-900 transition-colors cursor-pointer min-h-[38px] whitespace-nowrap"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>New Blank Job</span>
          </button>
        </div>

        {/* Workflow Navigation */}
        <nav className="flex-1 overflow-y-auto px-3 py-4 space-y-5">
          {navItems.map((section) => (
            <div key={section.group}>
              <div className="px-2.5 mb-1.5 text-[11px] font-semibold text-slate-400">
                {section.group}
              </div>
              <div className="space-y-0.5">
                {section.items.map((item) => {
                  const isActive = activeTab === item.id;
                  const issuedInfo = item.docPrefix ? activeJob.issuedDocs[item.docPrefix] : undefined;
                  const isIssued = Boolean(issuedInfo && !issuedInfo.isVoided);

                  return (
                    <button
                      key={item.id}
                      type="button"
                      onClick={() => {
                        setActiveTab(item.id);
                        setSidebarOpen(false);
                      }}
                      className={`w-full flex items-center justify-between gap-2 px-2.5 py-2 rounded-lg text-xs font-medium transition-colors cursor-pointer min-h-[40px] whitespace-nowrap ${
                        isActive
                          ? 'text-white'
                          : 'text-slate-700 hover:bg-slate-100 hover:text-slate-900'
                      }`}
                      style={isActive ? { backgroundColor: theme.primary } : undefined}
                    >
                      <span className="flex items-center gap-2.5 min-w-0">
                        <span className={isActive ? 'text-white' : 'text-slate-400'}>{item.icon}</span>
                        <span className="truncate">{item.label}</span>
                      </span>

                      {isIssued && (
                        <span
                          className={`font-mono-tabular text-[11px] shrink-0 ${
                            isActive ? 'text-white/90' : 'text-emerald-700 font-semibold'
                          }`}
                        >
                          #{String(issuedInfo?.sequence).padStart(3, '0')}
                        </span>
                      )}
                      {item.id === 'REGISTRY' && (
                        <span
                          className={`font-mono-tabular text-[11px] shrink-0 ${
                            isActive ? 'text-white/90' : 'text-slate-400'
                          }`}
                        >
                          #{String(nextSeq).padStart(3, '0')}
                        </span>
                      )}
                    </button>
                  );
                })}
              </div>
            </div>
          ))}
        </nav>

        {/* Sidebar Footer: Offline PWA Install */}
        <div className="p-4 border-t border-slate-200 space-y-2">
          <PWAInstallButton />
          <div className="text-[11px] text-slate-400 text-center font-mono-tabular">
            Next Registry Seq · #{String(nextSeq).padStart(3, '0')}
          </div>
        </div>
      </aside>

      {/* Main Content Shell */}
      <div className="flex-1 flex flex-col min-w-0">
        {/* Clean Single-Row Top Bar Contract (Breadcrumb + 1 Primary Action) */}
        <header className="no-print sticky top-0 z-30 h-16 border-b border-slate-200 bg-white/95 backdrop-blur px-4 sm:px-6 flex items-center justify-between gap-6">
          <div className="flex items-center gap-3 min-w-0">
            <button
              type="button"
              onClick={() => setSidebarOpen(true)}
              className="lg:hidden p-2 -ml-1 rounded-lg text-slate-600 hover:bg-slate-100 cursor-pointer min-h-[40px] min-w-[40px] flex items-center justify-center"
              aria-label="Open menu"
            >
              <Menu className="w-5 h-5" />
            </button>
            <div className="flex items-center gap-2 text-sm min-w-0">
              <span className="hidden sm:inline text-slate-400 truncate max-w-[200px]">
                {activeJob.jobName}
              </span>
              <span className="hidden sm:inline text-slate-300">/</span>
              <h1 className="font-display font-semibold text-slate-900 truncate">
                {TAB_LABELS[activeTab]}
              </h1>
            </div>
          </div>

          <div className="flex items-center gap-3 shrink-0">
            <input
              ref={quickImportRef}
              type="file"
              accept=".xlsx,.pdf,application/pdf"
              onChange={handleQuickExcelOpen}
              className="hidden"
            />
            <button
              type="button"
              onClick={() => quickImportRef.current?.click()}
              className="inline-flex items-center gap-2 rounded-lg px-4 py-2 text-xs font-semibold text-white transition-opacity hover:opacity-90 cursor-pointer min-h-[40px] whitespace-nowrap shrink-0"
              style={{ backgroundColor: theme.primary }}
            >
              <FolderOpen className="w-3.5 h-3.5" />
              <span>Open .xlsx / .pdf</span>
            </button>
          </div>
        </header>

        {/* Workspace Toolbar: Unboxed Metadata Summary + 3-Variant Excel Color Switcher */}
        <div className="no-print border-b border-slate-200 bg-slate-50 px-4 sm:px-6 py-2.5 flex flex-wrap items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2 text-xs text-slate-600">
            <span className="font-medium text-slate-900">{endUser}</span>
            <span aria-hidden="true" className="text-slate-300">·</span>
            <span>
              {activeJob.controls.incoterm} {activeJob.controls.destination}
            </span>
            <span aria-hidden="true" className="text-slate-300">·</span>
            <span className="font-mono-tabular">
              Markup {formatPct(activeJob.freight.selectedMarkupPct)}
            </span>
            <span aria-hidden="true" className="text-slate-300">·</span>
            <span className="font-mono-tabular font-semibold text-slate-900">
              Grand Total {formatMoney(calc.grandTotal, activeJob.controls.currency)}
            </span>
          </div>

          {/* Platform Version Switcher (Windows vs iPad) + 3-Variant Excel Style Switcher */}
          <div className="flex flex-wrap items-center gap-3">
            <div className="inline-flex items-center p-0.5 rounded-lg bg-slate-200/80">
              <button
                type="button"
                onClick={() => setPlatformMode('windows')}
                className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap min-h-[30px] ${
                  platformMode === 'windows'
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Monitor className="w-3.5 h-3.5" />
                <span>Windows</span>
              </button>
              <button
                type="button"
                onClick={() => setPlatformMode('ipad')}
                className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap min-h-[30px] ${
                  platformMode === 'ipad'
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Tablet className="w-3.5 h-3.5" />
                <span>iPad / iOS</span>
              </button>
            </div>

            <div className="flex items-center gap-2">
              <span className="text-xs text-slate-500 hidden sm:inline whitespace-nowrap">
                Style:
              </span>
              <div className="inline-flex items-center p-0.5 rounded-lg bg-slate-200/80">
                {(['v1', 'v2', 'v3'] as ThemeVariantId[]).map((id) => {
                  const t = EXCEL_THEMES[id];
                  const isSelected = themeId === id;
                  return (
                    <button
                      key={id}
                      type="button"
                      onClick={() => setThemeId(id)}
                      className={`px-2.5 py-1 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap min-h-[30px] ${
                        isSelected
                          ? 'bg-white text-slate-900 shadow-sm'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      {t.shortLabel}
                    </button>
                  );
                })}
              </div>
            </div>
          </div>
        </div>

        {/* iPad Quick Document Selector Strip when viewing Trade Documents */}
        {platformMode === 'ipad' && isDocTab(activeTab) && (
          <div className="no-print border-b border-slate-200 bg-white px-4 py-2 overflow-x-auto">
            <div className="flex items-center gap-1.5 min-w-max">
              {(['PI', 'QT', 'PO', 'PL', 'DN', 'CI'] as DocPrefix[]).map((doc) => {
                const isSelected = activeTab === doc;
                const issuedInfo = activeJob.issuedDocs[doc];
                return (
                  <button
                    key={doc}
                    type="button"
                    onClick={() => setActiveTab(doc)}
                    className={`px-3.5 py-2 rounded-lg text-xs font-semibold transition-colors cursor-pointer min-h-[40px] whitespace-nowrap ${
                      isSelected
                        ? 'text-white'
                        : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                    }`}
                    style={isSelected ? { backgroundColor: theme.primary } : undefined}
                  >
                    {DOC_META[doc].shortName}
                    {issuedInfo && !issuedInfo.isVoided ? ` · #${String(issuedInfo.sequence).padStart(3, '0')}` : ''}
                  </button>
                );
              })}
            </div>
          </div>
        )}

        {/* Active Viewport Content */}
        <main className={`flex-1 p-4 sm:p-6 max-w-[1440px] w-full mx-auto ${platformMode === 'ipad' ? 'pb-24' : ''}`}>
          {isDocTab(activeTab) && (
            <DocumentWorkspaceView
              key={`${activeJob.id}-${activeTab}`}
              job={activeJob}
              prefix={activeTab}
              calc={calc}
              themeId={themeId}
              company={company}
              platformMode={platformMode}
              nextSeq={nextSeq}
              onUpdateJobName={handleUpdateJobName}
              onUpdateControls={handleUpdateControls}
              onUpdateCustomColumns={handleUpdateCustomColumns}
              onUpdateItem={handleUpdateItem}
              onIssueDocument={handleIssueDocument}
              onDuplicateAsNewJob={handleDuplicateAsNewJob}
              onNotify={showToast}
            />
          )}

          {activeTab === 'PRICE_CALC' && (
            <PriceCalcView
              job={activeJob}
              calc={calc}
              themeId={themeId}
              platformMode={platformMode}
              onUpdateItem={handleUpdateItem}
              onUpdateFreight={handleUpdateFreight}
            />
          )}

          {activeTab === 'REGISTRY' && (
            <RegistryView
              entries={registry}
              nextSeq={nextSeq}
              themeId={themeId}
              company={company}
              onVoidEntry={handleVoidRegistryEntry}
              onSyncRegistryEntries={handleSyncRegistry}
              onNotify={showToast}
            />
          )}

          {activeTab === 'STUDIO' && (
            <ExcelStudioView
              activeJob={activeJob}
              allJobs={jobs}
              registry={registry}
              themeId={themeId}
              company={company}
              onSelectTheme={setThemeId}
              onUpdateCompany={handleUpdateCompany}
              onResetCompany={handleResetCompany}
              onImportJob={handleImportJob}
              onOpenSupplierQuote={setPendingSupplierQuote}
              onSyncRegistry={handleSyncRegistry}
              onRestoreBackup={handleRestoreBackup}
              onNotify={showToast}
            />
          )}
        </main>
      </div>

      {/* Supplier Quote (.pdf / any .xlsx / Pasted Text) Review & Load Modal */}
      {pendingSupplierQuote && (
        <SupplierQuoteImportModal
          initialQuote={pendingSupplierQuote}
          activeJobName={activeJob.jobName}
          themeId={themeId}
          onApply={handleApplySupplierQuote}
          onClose={() => setPendingSupplierQuote(null)}
        />
      )}

      {/* iPadOS Bottom Touch Navigation Dock */}
      {platformMode === 'ipad' && (
        <nav
          aria-label="iPad Navigation"
          className="no-print fixed bottom-0 inset-x-0 z-40 border-t border-slate-200 bg-white/95 backdrop-blur px-2 pt-1.5 pb-[max(0.5rem,env(safe-area-inset-bottom))]"
        >
          <div className="max-w-2xl mx-auto grid grid-cols-5 gap-1">
            {[
              { id: 'PI' as WorkspaceTab, label: 'Entry / PI', icon: <FileText className="w-5 h-5" />, active: activeTab === 'PI' },
              { id: 'PRICE_CALC' as WorkspaceTab, label: 'Price Calc', icon: <Calculator className="w-5 h-5" />, active: activeTab === 'PRICE_CALC' },
              { id: 'QT' as WorkspaceTab, label: 'Docs', icon: <Receipt className="w-5 h-5" />, active: ['QT', 'PO', 'PL', 'DN', 'CI'].includes(activeTab) },
              { id: 'REGISTRY' as WorkspaceTab, label: 'Registry', icon: <Hash className="w-5 h-5" />, active: activeTab === 'REGISTRY' },
              { id: 'STUDIO' as WorkspaceTab, label: 'Studio', icon: <Palette className="w-5 h-5" />, active: activeTab === 'STUDIO' },
            ].map((item) => (
              <button
                key={item.id}
                type="button"
                onClick={() => setActiveTab(item.id)}
                className={`flex flex-col items-center justify-center gap-1 rounded-lg py-1.5 min-h-[48px] text-[11px] font-medium transition-colors cursor-pointer ${
                  item.active ? 'text-white' : 'text-slate-600 hover:bg-slate-100'
                }`}
                style={item.active ? { backgroundColor: theme.primary } : undefined}
              >
                {item.icon}
                <span className="whitespace-nowrap">{item.label}</span>
              </button>
            ))}
          </div>
        </nav>
      )}

      {/* Transient Notification Toast */}
      {toast && (
        <div className={`no-print fixed ${platformMode === 'ipad' ? 'bottom-20' : 'bottom-4'} left-1/2 -translate-x-1/2 z-50 flex items-center gap-2.5 rounded-lg border border-slate-800 bg-slate-900 px-4 py-2.5 text-xs font-medium text-white shadow-lg`}>
          <FileSpreadsheet className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toast}</span>
        </div>
      )}

      <OfflineIndicator />
    </div>
  );
}
