import React, { useMemo, useState } from 'react';
import {
  ArrowRight,
  BarChart3,
  Check,
  Clock,
  Eye,
  EyeOff,
  FileDown,
  FileSpreadsheet,
  Plus,
  Search,
  SlidersHorizontal,
  TrendingUp,
} from 'lucide-react';
import {
  CompanyProfile,
  DashboardPreferences,
  InquiryType,
  ProcurementJob,
  RegistryEntry,
  ThemePalette,
  WorkspaceTab,
} from '../types/procurement';
import {
  calculateJob,
  formatMoney,
  getJobCompletionMetrics,
  getResolvedEndUser,
} from '../utils/calculations';
import { DashboardPdfOrientation, exportDashboardReportPdf } from "../utils/dashboardPdf";
import { exportDashboardReportXlsx } from '../utils/excelEngine';
import { useRowSelection } from '../hooks/useRowSelection';
import { HelpTip } from './HelpTip';

interface JobTrackerDashboardProps {
  jobs: ProcurementJob[];
  activeJobId: string;
  registry: RegistryEntry[];
  theme: ThemePalette;
  company: CompanyProfile;
  prefs: DashboardPreferences;
  onUpdatePrefs: (patch: Partial<DashboardPreferences>) => void;
  onSelectJobAndTab: (jobId: string, tab: WorkspaceTab) => void;
  onCreateNewJob: () => void;
  onNotify: (msg: string) => void;
}

type StageFilter =
  | 'ALL'
  | 'DRAFT'
  | 'QT'
  | 'PI'
  | 'PO'
  | 'CI'
  | 'PL'
  | 'DN';

interface EnrichedJob {
  job: ProcurementJob;
  client: string;
  calc: ReturnType<typeof calculateJob>;
  metrics: ReturnType<typeof getJobCompletionMetrics>;
  marginPct: number;
}

function getIsoWeekKey(date: Date): string {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil(((d.getTime() - yearStart.getTime()) / 86400000 + 1) / 7);
  return `${d.getUTCFullYear()}-W${String(weekNo).padStart(2, '0')}`;
}

function getMonthKey(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

export const JobTrackerDashboard: React.FC<JobTrackerDashboardProps> = ({
  jobs,
  activeJobId,
  registry,
  theme,
  company,
  prefs,
  onUpdatePrefs,
  onSelectJobAndTab,
  onCreateNewJob,
  onNotify,
}) => {
  const [customizeOpen, setCustomizeOpen] = useState(false);
  const [search, setSearch] = useState('');
  const [inquiryFilter, setInquiryFilter] = useState<'ALL' | InquiryType>('ALL');
  const [stageFilter, setStageFilter] = useState<StageFilter>('ALL');
  const [exportingExcel, setExportingExcel] = useState(false);
  const [exportingPdf, setExportingPdf] = useState(false);
  const [pdfModalOpen, setPdfModalOpen] = useState(false);
  const [pdfOrientation, setPdfOrientation] = useState<DashboardPdfOrientation>("landscape");
  const jobSel = useRowSelection<string>();
  const exportSelectedJobs = async (kind: "xlsx" | "pdf") => {
    const chosen = jobs.filter((j) => jobSel.has(j.id));
    if (chosen.length === 0) return;
    try {
      const filename =
        kind === "xlsx"
          ? await exportDashboardReportXlsx(chosen, registry, theme, company, prefs)
          : await exportDashboardReportPdf(chosen, registry, theme, company, prefs, "landscape");
      onNotify(`Exported ${filename} (${chosen.length} selected job${chosen.length === 1 ? "" : "s"})`);
    } catch {
      onNotify("Could not export the selected jobs. Please try again.");
    }
  };

  const handleDownloadExcelReport = async () => {
    setExportingExcel(true);
    try {
      const filename = await exportDashboardReportXlsx(jobs, registry, theme, company, prefs);
      onNotify(`Downloaded Dashboard Excel report: ${filename}`);
    } catch {
      onNotify('Could not generate Dashboard Excel report. Please try again.');
    } finally {
      setExportingExcel(false);
    }
  };

  const handleDownloadPdfReport = async () => {
    setExportingPdf(true);
    try {
      const filename = await exportDashboardReportPdf(jobs, registry, theme, company, prefs, pdfOrientation);
      onNotify(`Downloaded Dashboard PDF report: ${filename}`);
      setPdfModalOpen(false);
    } catch {
      onNotify("Could not generate Dashboard PDF report. Please try again.");
    } finally {
      setExportingPdf(false);
    }
  };

  // Enrich all jobs with live calculations & stage completion percentages
  const enrichedAll: EnrichedJob[] = useMemo(() => {
    return jobs.map((job) => {
      const calc = calculateJob(job);
      const metrics = getJobCompletionMetrics(job);
      const client = getResolvedEndUser(job) || 'Unassigned Client';
      const marginPct =
        calc.goodsSubtotalExVat > 0
          ? Math.round((calc.grossProfit / calc.goodsSubtotalExVat) * 1000) / 10
          : 0;
      return { job, client, calc, metrics, marginPct };
    });
  }, [jobs]);

  const scopedJobs = useMemo(
    () =>
      prefs.includeCompleted
        ? enrichedAll
        : enrichedAll.filter((item) => !item.metrics.isComplete),
    [enrichedAll, prefs.includeCompleted]
  );

  // Separate RFQ & Tender counters
  const rfqJobs = useMemo(
    () => scopedJobs.filter((j) => j.job.inquiryType !== 'TENDER'),
    [scopedJobs]
  );
  const tenderJobs = useMemo(
    () => scopedJobs.filter((j) => j.job.inquiryType === 'TENDER'),
    [scopedJobs]
  );

  const totalCount = scopedJobs.length;
  const rfqValue = rfqJobs.reduce((s, j) => s + j.calc.grandTotal, 0);
  const tenderValue = tenderJobs.reduce((s, j) => s + j.calc.grandTotal, 0);
  const totalPipelineValue = rfqValue + tenderValue;
  const totalEstProfit = scopedJobs.reduce((s, j) => s + j.calc.grossProfit, 0);
  const avgCompletionPct =
    totalCount > 0
      ? Math.round(scopedJobs.reduce((s, j) => s + j.metrics.percent, 0) / totalCount)
      : 0;
  const completedCount = enrichedAll.filter((j) => j.metrics.isComplete).length;

  // Stage completion breakdown across all scoped jobs
  const stageBars = useMemo(() => {
    const denom = Math.max(1, scopedJobs.length);
    const count = (pred: (e: EnrichedJob) => boolean) => scopedJobs.filter(pred).length;
    const raw: {
      id: StageFilter | 'INT' | 'GATE_PO' | 'GATE_PL';
      label: string;
      count: number;
      filterTarget?: StageFilter;
    }[] = [
      { id: 'INT', label: 'Internal Supplier Quote', count: count((e) => e.metrics.hasInternalQuote) },
      { id: 'QT', label: 'Quotation (QT) Issued', count: count((e) => e.metrics.hasQT), filterTarget: 'QT' },
      { id: 'PI', label: 'Proforma Invoice (PI) Issued', count: count((e) => e.metrics.hasPI), filterTarget: 'PI' },
      { id: 'GATE_PO', label: 'Client PO Confirmed', count: count((e) => e.metrics.hasEndUserPoGate) },
      { id: 'PO', label: 'Supplier PO Issued', count: count((e) => e.metrics.hasPO), filterTarget: 'PO' },
      { id: 'CI', label: 'Commercial Invoice (CI) Issued', count: count((e) => e.metrics.hasCI), filterTarget: 'CI' },
      { id: 'GATE_PL', label: 'Supplier Paid & PL Confirmed', count: count((e) => e.metrics.hasSupplierShipmentGate) },
      { id: 'PL', label: 'Packing List (PL) Issued', count: count((e) => e.metrics.hasPL), filterTarget: 'PL' },
      { id: 'DN', label: 'Delivery Note (DN) Issued', count: count((e) => e.metrics.hasDN), filterTarget: 'DN' },
    ];
    return raw.map((r) => ({
      ...r,
      pct: scopedJobs.length === 0 ? 0 : Math.round((r.count / denom) * 100),
    }));
  }, [scopedJobs]);

  // Weekly (last 8 weeks) & Monthly (last 6 months) performance aggregation from real timestamps
  const performanceBuckets = useMemo(() => {
    const now = new Date();
    const activeRegistry = registry.filter((r) => r.status === 'ISSUED');

    if (prefs.performancePeriod === 'weekly') {
      const buckets: {
        key: string;
        label: string;
        jobsCreated: number;
        docsIssued: number;
        value: number;
      }[] = [];
      for (let i = 7; i >= 0; i--) {
        const d = new Date(now);
        d.setDate(now.getDate() - i * 7);
        const key = getIsoWeekKey(d);
        const label = `W${key.split('-W')[1]}`;
        buckets.push({ key, label, jobsCreated: 0, docsIssued: 0, value: 0 });
      }
      const map = new Map(buckets.map((b) => [b.key, b]));
      for (const ej of enrichedAll) {
        const dt = new Date(ej.job.createdAt);
        if (!Number.isNaN(dt.getTime())) {
          const b = map.get(getIsoWeekKey(dt));
          if (b) {
            b.jobsCreated += 1;
            b.value += ej.calc.grandTotal;
          }
        }
      }
      for (const r of activeRegistry) {
        const dt = new Date(r.issuedAt);
        if (!Number.isNaN(dt.getTime())) {
          const b = map.get(getIsoWeekKey(dt));
          if (b) b.docsIssued += 1;
        }
      }
      return buckets;
    } else {
      const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      const buckets: {
        key: string;
        label: string;
        jobsCreated: number;
        docsIssued: number;
        value: number;
      }[] = [];
      for (let i = 5; i >= 0; i--) {
        const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
        const key = getMonthKey(d);
        const label = `${monthNames[d.getMonth()]} ${String(d.getFullYear()).slice(2)}`;
        buckets.push({ key, label, jobsCreated: 0, docsIssued: 0, value: 0 });
      }
      const map = new Map(buckets.map((b) => [b.key, b]));
      for (const ej of enrichedAll) {
        const dt = new Date(ej.job.createdAt);
        if (!Number.isNaN(dt.getTime())) {
          const b = map.get(getMonthKey(dt));
          if (b) {
            b.jobsCreated += 1;
            b.value += ej.calc.grandTotal;
          }
        }
      }
      for (const r of activeRegistry) {
        const dt = new Date(r.issuedAt);
        if (!Number.isNaN(dt.getTime())) {
          const b = map.get(getMonthKey(dt));
          if (b) b.docsIssued += 1;
        }
      }
      return buckets;
    }
  }, [enrichedAll, registry, prefs.performancePeriod]);

  // Priority Ranking List (switchable by Estimated Profitability OR Overall Value)
  const priorityRanked = useMemo(() => {
    return [...scopedJobs]
      .sort((a, b) => {
        if (prefs.priorityMetric === 'profit') {
          if (b.calc.grossProfit !== a.calc.grossProfit) {
            return b.calc.grossProfit - a.calc.grossProfit;
          }
          return b.calc.grandTotal - a.calc.grandTotal;
        }
        if (b.calc.grandTotal !== a.calc.grandTotal) {
          return b.calc.grandTotal - a.calc.grandTotal;
        }
        return b.calc.grossProfit - a.calc.grossProfit;
      })
      .slice(0, 8);
  }, [scopedJobs, prefs.priorityMetric]);

  const maxPriorityMetric = useMemo(() => {
    const vals = priorityRanked.map((r) =>
      prefs.priorityMetric === 'profit' ? Math.max(0, r.calc.grossProfit) : Math.max(0, r.calc.grandTotal)
    );
    return Math.max(1, ...vals);
  }, [priorityRanked, prefs.priorityMetric]);

  // Filtered Matrix Table rows
  const matrixRows = useMemo(() => {
    const q = search.trim().toLowerCase();
    return scopedJobs.filter((ej) => {
      if (inquiryFilter !== 'ALL' && ej.job.inquiryType !== inquiryFilter) return false;
      if (stageFilter === 'DRAFT' && (ej.metrics.hasQT || ej.metrics.hasPI)) return false;
      if (stageFilter === 'QT' && !ej.metrics.hasQT) return false;
      if (stageFilter === 'PI' && !ej.metrics.hasPI) return false;
      if (stageFilter === 'PO' && !ej.metrics.hasPO) return false;
      if (stageFilter === 'CI' && !ej.metrics.hasCI) return false;
      if (stageFilter === 'PL' && !ej.metrics.hasPL) return false;
      if (stageFilter === 'DN' && !ej.metrics.hasDN) return false;

      if (!q) return true;
      const seqStr = ej.job.jobSequence ? String(ej.job.jobSequence).padStart(3, '0') : '';
      return (
        ej.job.jobName.toLowerCase().includes(q) ||
        ej.client.toLowerCase().includes(q) ||
        (ej.job.controls.rfqRef || '').toLowerCase().includes(q) ||
        seqStr.includes(q)
      );
    });
  }, [scopedJobs, search, inquiryFilter, stageFilter]);

  // SVG Ring math for Completion %
  const ringRadius = 26;
  const ringCirc = 2 * Math.PI * ringRadius;
  const ringOffset = ringCirc - (avgCompletionPct / 100) * ringCirc;

  const maxPerfCount = Math.max(
    1,
    ...performanceBuckets.map((b) => Math.max(b.docsIssued, b.jobsCreated))
  );

  const toggleItems: { key: keyof DashboardPreferences; label: string }[] = [
    { key: 'showRfqTenderCounters', label: 'RFQ & Tender Counters' },
    { key: 'showCompletionFunnel', label: 'Completion % & Stage Funnel' },
    { key: 'showPerformanceChart', label: 'Weekly / Monthly Performance' },
    { key: 'showPriorityRanking', label: 'Priority Ranking (Profit / Value)' },
    { key: 'showJobMatrix', label: 'Job Lifecycle Tracker Table' },
    { key: 'showFinancials', label: 'Show Values & Profit' },
    { key: 'includeCompleted', label: 'Include Delivered Jobs' },
  ];

  const renderStageCell = (
    jobId: string,
    targetTab: WorkspaceTab,
    issued: boolean,
    docNo?: string,
    gateReady = true
  ) => (
    <button
      type="button"
      onClick={() => onSelectJobAndTab(jobId, targetTab)}
      title={docNo ? `Open ${docNo}` : `Open ${targetTab}`}
      className={`w-full min-h-[32px] rounded px-1.5 py-1 inline-flex items-center justify-center gap-1 text-[11px] font-medium transition-colors cursor-pointer whitespace-nowrap ${
        issued
          ? 'text-emerald-800 hover:bg-emerald-50 font-semibold'
          : gateReady
          ? 'text-slate-600 hover:bg-slate-100'
          : 'text-amber-700 hover:bg-amber-50'
      }`}
    >
      {issued ? (
        <>
          <Check className="w-3.5 h-3.5 text-emerald-700 shrink-0" />
          <span className="font-mono-tabular">Issued</span>
        </>
      ) : gateReady ? (
        <span className="text-slate-400">—</span>
      ) : (
        <>
          <Clock className="w-3 h-3 text-amber-600 shrink-0" />
          <span>Pending</span>
        </>
      )}
    </button>
  );

  return (
    <div className="space-y-6">
      {/* Top Action & Customizer Strip */}
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <div className="flex items-center gap-2">
            <h2 className="text-lg font-semibold text-slate-900">
              Dashboard
            </h2>
            <HelpTip text="Track every procurement job from Internal Supplier Quote to Delivery Note. Use Customize View to toggle dashboard panels and financial metrics ON or OFF, or download the active report in Excel or PDF." />
          </div>
          <div className="text-xs text-slate-500 mt-0.5 flex flex-wrap items-center gap-2 font-mono-tabular">
            <span>{enrichedAll.length} Total Job{enrichedAll.length === 1 ? '' : 's'}</span>
            <span>·</span>
            <span>{completedCount} Delivered</span>
            <span>·</span>
            <span>{avgCompletionPct}% Average Completion</span>
          </div>
        </div>

        <div className="no-print flex flex-wrap items-center gap-2">
          <button
            type="button"
            disabled={exportingExcel}
            onClick={handleDownloadExcelReport}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs font-medium text-slate-800 hover:bg-slate-50 cursor-pointer min-h-[40px] whitespace-nowrap disabled:opacity-50"
          >
            <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-700" />
            <span>{exportingExcel ? 'Exporting...' : 'Excel Report'}</span>
          </button>

          <button
            type="button"
            disabled={exportingPdf}
            onClick={() => setPdfModalOpen(true)}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs font-medium text-slate-800 hover:bg-slate-50 cursor-pointer min-h-[40px] whitespace-nowrap disabled:opacity-50"
          >
            <FileDown className="w-3.5 h-3.5 text-red-700" />
            <span>{exportingPdf ? 'Exporting...' : 'PDF Report'}</span>
          </button>

          <button
            type="button"
            onClick={() => onUpdatePrefs({ showFinancials: !prefs.showFinancials })}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[40px] whitespace-nowrap"
          >
            {prefs.showFinancials ? (
              <>
                <Eye className="w-3.5 h-3.5 text-slate-500" />
                <span>Figures Visible</span>
              </>
            ) : (
              <>
                <EyeOff className="w-3.5 h-3.5 text-slate-500" />
                <span>Figures Hidden</span>
              </>
            )}
          </button>

          <button
            type="button"
            aria-expanded={customizeOpen}
            onClick={() => setCustomizeOpen((v) => !v)}
            className={`inline-flex items-center gap-1.5 rounded-lg border px-3.5 py-2 text-xs font-medium cursor-pointer min-h-[40px] whitespace-nowrap transition-colors ${
              customizeOpen
                ? 'border-slate-900 bg-slate-900 text-white'
                : 'border-slate-200 bg-white text-slate-800 hover:bg-slate-50'
            }`}
          >
            <SlidersHorizontal className="w-3.5 h-3.5" />
            <span>Customize View</span>
          </button>

          <button
            type="button"
            onClick={onCreateNewJob}
            className="inline-flex items-center gap-1.5 rounded-lg px-4 py-2 text-xs font-semibold text-white hover:opacity-90 cursor-pointer min-h-[40px] whitespace-nowrap"
            style={{ backgroundColor: theme.primary, color: theme.primaryText }}
          >
            <Plus className="w-3.5 h-3.5" />
            <span>New Job</span>
          </button>
        </div>
      </div>

      {/* Collapsible Dashboard Parameter ON/OFF Switch Panel */}
      {customizeOpen && (
        <div className="rounded-lg border border-slate-200 bg-white p-4 space-y-3">
          <div className="flex items-center justify-between gap-2">
            <span className="text-xs font-semibold text-slate-900">
              Dashboard Parameter Switches (Saved Automatically)
            </span>
            <span className="text-xs text-slate-500">
              Toggle modules on or off to tailor your view
            </span>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2.5">
            {toggleItems.map(({ key, label }) => {
              const checked = Boolean(prefs[key]);
              return (
                <button
                  key={key}
                  type="button"
                  role="switch"
                  aria-checked={checked}
                  onClick={() => onUpdatePrefs({ [key]: !checked })}
                  className="flex items-center justify-between gap-3 rounded-lg border border-slate-200 px-3 py-2 text-left hover:bg-slate-50 cursor-pointer min-h-[40px]"
                >
                  <span className="text-xs font-medium text-slate-800 truncate">{label}</span>
                  <span
                    className={`relative inline-flex h-5 w-9 shrink-0 items-center rounded-full transition-colors ${
                      checked ? 'bg-slate-900' : 'bg-slate-200'
                    }`}
                  >
                    <span
                      className={`inline-block h-3.5 w-3.5 rounded-full bg-white transition-transform ${
                        checked ? 'translate-x-4.5' : 'translate-x-1'
                      }`}
                    />
                  </span>
                </button>
              );
            })}
          </div>
        </div>
      )}

      {/* Section 1: Separate RFQ & Tender Counters + Completion Ring + Pipeline Value */}
      {prefs.showRfqTenderCounters && (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          {/* Separate RFQ Counter */}
          <button
            type="button"
            onClick={() => setInquiryFilter((f) => (f === 'RFQ' ? 'ALL' : 'RFQ'))}
            className={`text-left rounded-lg border p-4 transition-colors cursor-pointer ${
              inquiryFilter === 'RFQ'
                ? 'border-slate-900 bg-slate-50/70'
                : 'border-slate-200 bg-white hover:border-slate-300'
            }`}
          >
            <div className="flex items-center justify-between text-xs text-slate-500">
              <span>Active RFQs</span>
              <span className="font-mono-tabular">
                {totalCount > 0 ? Math.round((rfqJobs.length / totalCount) * 100) : 0}% of jobs
              </span>
            </div>
            <div className="mt-1.5 text-2xl font-bold font-mono-tabular text-slate-900">
              {rfqJobs.length}
            </div>
            <div className="mt-1 text-xs font-mono-tabular text-slate-600">
              {prefs.showFinancials ? formatMoney(rfqValue, 'USD') : 'Click to filter RFQs'}
            </div>
          </button>

          {/* Separate Tender Counter */}
          <button
            type="button"
            onClick={() => setInquiryFilter((f) => (f === 'TENDER' ? 'ALL' : 'TENDER'))}
            className={`text-left rounded-lg border p-4 transition-colors cursor-pointer ${
              inquiryFilter === 'TENDER'
                ? 'border-slate-900 bg-slate-50/70'
                : 'border-slate-200 bg-white hover:border-slate-300'
            }`}
          >
            <div className="flex items-center justify-between text-xs text-slate-500">
              <span>Active Tenders</span>
              <span className="font-mono-tabular">
                {totalCount > 0 ? Math.round((tenderJobs.length / totalCount) * 100) : 0}% of jobs
              </span>
            </div>
            <div className="mt-1.5 text-2xl font-bold font-mono-tabular text-slate-900">
              {tenderJobs.length}
            </div>
            <div className="mt-1 text-xs font-mono-tabular text-slate-600">
              {prefs.showFinancials ? formatMoney(tenderValue, 'USD') : 'Click to filter Tenders'}
            </div>
          </button>

          {/* Overall Completion % Ring */}
          <div className="rounded-lg border border-slate-200 bg-white p-4 flex items-center justify-between gap-3">
            <div>
              <div className="text-xs text-slate-500">Overall Completion</div>
              <div className="mt-1.5 text-2xl font-bold font-mono-tabular text-slate-900">
                {avgCompletionPct}%
              </div>
              <div className="mt-1 text-xs text-slate-500">
                {completedCount} of {enrichedAll.length} fully delivered
              </div>
            </div>
            <svg width="64" height="64" viewBox="0 0 64 64" className="shrink-0" aria-hidden="true">
              <circle
                cx="32"
                cy="32"
                r={ringRadius}
                fill="none"
                stroke="#E2E8F0"
                strokeWidth="6"
              />
              <circle
                cx="32"
                cy="32"
                r={ringRadius}
                fill="none"
                stroke={theme.primary}
                strokeWidth="6"
                strokeLinecap="round"
                strokeDasharray={ringCirc}
                strokeDashoffset={ringOffset}
                transform="rotate(-90 32 32)"
              />
              <text
                x="32"
                y="36"
                textAnchor="middle"
                className="text-[11px] font-bold fill-slate-900 font-mono-tabular"
              >
                {avgCompletionPct}%
              </text>
            </svg>
          </div>

          {/* Pipeline Value & Est. Gross Profit */}
          <div className="rounded-lg border border-slate-200 bg-white p-4">
            <div className="text-xs text-slate-500">
              {prefs.showFinancials ? 'Pipeline Value & Est. Profit' : 'Issued Documents'}
            </div>
            {prefs.showFinancials ? (
              <>
                <div className="mt-1.5 text-xl font-bold font-mono-tabular text-slate-900 truncate">
                  {formatMoney(totalPipelineValue, 'USD')}
                </div>
                <div className="mt-1 text-xs font-mono-tabular text-emerald-700 font-semibold">
                  Est. Profit: {formatMoney(totalEstProfit, 'USD')}
                </div>
              </>
            ) : (
              <>
                <div className="mt-1.5 text-2xl font-bold font-mono-tabular text-slate-900">
                  {registry.filter((r) => r.status === 'ISSUED').length}
                </div>
                <div className="mt-1 text-xs text-slate-500">Financial figures hidden</div>
              </>
            )}
          </div>
        </div>
      )}

      {/* Section 2: Graphical Analytics Split (Stage Completion Funnel + Weekly/Monthly Performance) */}
      {(prefs.showCompletionFunnel || prefs.showPerformanceChart) && (
        <div
          className={`grid grid-cols-1 ${
            prefs.showCompletionFunnel && prefs.showPerformanceChart ? 'lg:grid-cols-12' : ''
          } gap-5 items-stretch`}
        >
          {/* Left: Stage Completion % Funnel */}
          {prefs.showCompletionFunnel && (
            <div
              className={`${
                prefs.showPerformanceChart ? 'lg:col-span-6' : ''
              } rounded-lg border border-slate-200 bg-white p-5 flex flex-col justify-between space-y-4`}
            >
              <div className="flex items-center justify-between gap-2">
                <div className="flex items-center gap-1.5">
                  <BarChart3 className="w-4 h-4 text-slate-600" />
                  <h3 className="text-sm font-semibold text-slate-900">
                    Stage Completion Breakdown (%)
                  </h3>
                </div>
                <span className="text-xs font-mono-tabular text-slate-500">
                  {totalCount} Active Job{totalCount === 1 ? '' : 's'}
                </span>
              </div>

              <div className="space-y-2.5">
                {stageBars.map((bar) => {
                  const isActiveFilter = bar.filterTarget && stageFilter === bar.filterTarget;
                  return (
                    <button
                      key={bar.id}
                      type="button"
                      disabled={!bar.filterTarget}
                      onClick={() => {
                        if (!bar.filterTarget) return;
                        setStageFilter((cur) => (cur === bar.filterTarget ? 'ALL' : bar.filterTarget!));
                      }}
                      className={`w-full text-left space-y-1 rounded-md px-1.5 py-0.5 transition-colors ${
                        bar.filterTarget ? 'cursor-pointer hover:bg-slate-50' : 'cursor-default'
                      } ${isActiveFilter ? 'bg-slate-100' : ''}`}
                    >
                      <div className="flex items-center justify-between text-xs">
                        <span className="font-medium text-slate-700">{bar.label}</span>
                        <span className="font-mono-tabular text-slate-600">
                          <strong className="text-slate-900">{bar.count}</strong>/{totalCount} ·{' '}
                          <span className="font-semibold text-slate-900">{bar.pct}%</span>
                        </span>
                      </div>
                      <div className="h-2 w-full rounded-full bg-slate-100 overflow-hidden">
                        <div
                          className="h-full rounded-full transition-all duration-150"
                          style={{
                            width: `${bar.pct}%`,
                            backgroundColor:
                              bar.id === 'DN'
                                ? '#15803D'
                                : bar.id === 'GATE_PO' || bar.id === 'GATE_PL'
                                ? '#D97706'
                                : theme.primary,
                          }}
                        />
                      </div>
                    </button>
                  );
                })}
              </div>
            </div>
          )}

          {/* Right: Weekly / Monthly Performance SVG Chart */}
          {prefs.showPerformanceChart && (
            <div
              className={`${
                prefs.showCompletionFunnel ? 'lg:col-span-6' : ''
              } rounded-lg border border-slate-200 bg-white p-5 flex flex-col justify-between space-y-4`}
            >
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-1.5">
                  <TrendingUp className="w-4 h-4 text-slate-600" />
                  <h3 className="text-sm font-semibold text-slate-900">
                    {prefs.performancePeriod === 'weekly'
                      ? 'Weekly Performance (Last 8 Weeks)'
                      : 'Monthly Performance (Last 6 Months)'}
                  </h3>
                </div>

                <div className="inline-flex items-center p-1 rounded-lg bg-slate-100">
                  <button
                    type="button"
                    onClick={() => onUpdatePrefs({ performancePeriod: 'weekly' })}
                    className={`px-3 py-1 rounded-md text-xs font-medium cursor-pointer transition-colors whitespace-nowrap ${
                      prefs.performancePeriod === 'weekly'
                        ? 'bg-white text-slate-900 shadow-sm'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Weekly
                  </button>
                  <button
                    type="button"
                    onClick={() => onUpdatePrefs({ performancePeriod: 'monthly' })}
                    className={`px-3 py-1 rounded-md text-xs font-medium cursor-pointer transition-colors whitespace-nowrap ${
                      prefs.performancePeriod === 'monthly'
                        ? 'bg-white text-slate-900 shadow-sm'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Monthly
                  </button>
                </div>
              </div>

              {/* Clean Offline SVG Grouped Bar Chart */}
              <div className="w-full pt-2">
                <svg
                  viewBox="0 0 520 205"
                  className="w-full h-[205px] overflow-visible"
                  role="img"
                  aria-label={`${prefs.performancePeriod} performance chart`}
                >
                  {/* Horizontal Gridlines */}
                  {[0, 0.25, 0.5, 0.75, 1].map((frac, idx) => {
                    const y = 165 - frac * 140;
                    const tickVal = Math.round(maxPerfCount * frac);
                    return (
                      <g key={idx}>
                        <line
                          x1="34"
                          y1={y}
                          x2="510"
                          y2={y}
                          stroke="#E2E8F0"
                          strokeWidth="1"
                          strokeDasharray={frac === 0 ? undefined : '3 3'}
                        />
                        <text
                          x="26"
                          y={y + 4}
                          textAnchor="end"
                          className="text-[10px] fill-slate-400 font-mono-tabular"
                        >
                          {tickVal}
                        </text>
                      </g>
                    );
                  })}

                  {/* Bars */}
                  {performanceBuckets.map((b, idx) => {
                    const slotW = (510 - 44) / performanceBuckets.length;
                    const groupX = 44 + idx * slotW;
                    const barW = Math.min(18, Math.max(10, slotW * 0.28));
                    const jobsH = (b.jobsCreated / maxPerfCount) * 140;
                    const docsH = (b.docsIssued / maxPerfCount) * 140;
                    const centerX = groupX + slotW / 2;

                    return (
                      <g key={b.key}>
                        {/* Jobs Created Bar */}
                        <rect
                          x={centerX - barW - 2}
                          y={165 - jobsH}
                          width={barW}
                          height={Math.max(2, jobsH)}
                          rx="3"
                          fill="#94A3B8"
                        >
                          <title>{`${b.label}: ${b.jobsCreated} jobs started`}</title>
                        </rect>
                        {/* Documents Issued Bar */}
                        <rect
                          x={centerX + 2}
                          y={165 - docsH}
                          width={barW}
                          height={Math.max(2, docsH)}
                          rx="3"
                          fill={theme.primary}
                        >
                          <title>{`${b.label}: ${b.docsIssued} documents issued`}</title>
                        </rect>

                        {/* Value labels when non-zero */}
                        {b.docsIssued > 0 && (
                          <text
                            x={centerX + 2 + barW / 2}
                            y={160 - docsH}
                            textAnchor="middle"
                            className="text-[10px] font-semibold fill-slate-700 font-mono-tabular"
                          >
                            {b.docsIssued}
                          </text>
                        )}

                        {/* X Axis Label */}
                        <text
                          x={centerX}
                          y="186"
                          textAnchor="middle"
                          className="text-[10px] fill-slate-500 font-mono-tabular"
                        >
                          {b.label}
                        </text>
                      </g>
                    );
                  })}
                </svg>
              </div>

              {/* Clean Legend */}
              <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100 text-xs text-slate-600">
                <div className="flex items-center gap-4">
                  <span className="inline-flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-sm bg-slate-400 inline-block" />
                    <span>Jobs Started</span>
                  </span>
                  <span className="inline-flex items-center gap-1.5">
                    <span
                      className="w-2.5 h-2.5 rounded-sm inline-block"
                      style={{ backgroundColor: theme.primary }}
                    />
                    <span>Documents Issued</span>
                  </span>
                </div>
                <span className="font-mono-tabular text-slate-500">
                  Period Total:{' '}
                  <strong className="text-slate-900">
                    {performanceBuckets.reduce((s, b) => s + b.docsIssued, 0)} docs
                  </strong>
                </span>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Section 3: Priority Ranking List (Profitability vs. Overall Value) */}
      {prefs.showPriorityRanking && (
        <div className="rounded-lg border border-slate-200 bg-white p-5 space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex items-center gap-2">
              <h3 className="text-sm font-semibold text-slate-900">Priority Job Ranking</h3>
              <HelpTip text="Ranks active jobs from highest to lowest by either Estimated Gross Profit or Overall Job Value so you can prioritize high-impact RFQs and Tenders." />
            </div>

            <div className="inline-flex items-center p-1 rounded-lg bg-slate-100">
              <button
                type="button"
                onClick={() => onUpdatePrefs({ priorityMetric: 'profit' })}
                className={`px-3 py-1.5 rounded-md text-xs font-medium cursor-pointer transition-colors whitespace-nowrap ${
                  prefs.priorityMetric === 'profit'
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                By Est. Profitability
              </button>
              <button
                type="button"
                onClick={() => onUpdatePrefs({ priorityMetric: 'value' })}
                className={`px-3 py-1.5 rounded-md text-xs font-medium cursor-pointer transition-colors whitespace-nowrap ${
                  prefs.priorityMetric === 'value'
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                By Overall Value
              </button>
            </div>
          </div>

          {priorityRanked.length === 0 ? (
            <div className="py-6 text-center text-xs text-slate-500">
              No active jobs to rank yet. Create a job and add supplier pricing to see priority rankings.
            </div>
          ) : (
            <div className="divide-y divide-slate-100">
              {priorityRanked.map((item, rankIdx) => {
                const metricVal =
                  prefs.priorityMetric === 'profit'
                    ? Math.max(0, item.calc.grossProfit)
                    : Math.max(0, item.calc.grandTotal);
                const barWidth = Math.min(100, Math.round((metricVal / maxPriorityMetric) * 100));
                const curr = item.job.controls.currency;

                return (
                  <div
                    key={item.job.id}
                    className="py-3 flex flex-wrap sm:flex-nowrap items-center justify-between gap-4"
                  >
                    <div className="flex items-center gap-3 min-w-0 flex-1">
                      <span className="w-7 text-xs font-bold font-mono-tabular text-slate-500 shrink-0">
                        #{rankIdx + 1}
                      </span>
                      <div className="min-w-0 flex-1 space-y-1">
                        <div className="flex flex-wrap items-center gap-2 text-xs">
                          <span className="font-semibold text-slate-900 truncate">
                            {item.job.jobName}
                          </span>
                          <span className="text-slate-400">·</span>
                          <span className="text-slate-600 truncate">{item.client}</span>
                          <span className="text-slate-400">·</span>
                          <span className="font-mono-tabular text-slate-500">
                            {item.job.inquiryType}
                            {item.job.controls.rfqRef ? ` ${item.job.controls.rfqRef}` : ''}
                          </span>
                        </div>
                        <div className="h-1.5 w-full max-w-md rounded-full bg-slate-100 overflow-hidden">
                          <div
                            className="h-full rounded-full"
                            style={{
                              width: `${barWidth}%`,
                              backgroundColor:
                                prefs.priorityMetric === 'profit' ? '#15803D' : theme.primary,
                            }}
                          />
                        </div>
                      </div>
                    </div>

                    <div className="flex items-center gap-4 shrink-0">
                      <div className="text-right font-mono-tabular">
                        {prefs.showFinancials ? (
                          <>
                            <div className="text-xs font-bold text-slate-900">
                              {prefs.priorityMetric === 'profit'
                                ? formatMoney(item.calc.grossProfit, curr)
                                : formatMoney(item.calc.grandTotal, curr)}
                            </div>
                            <div className="text-[11px] text-slate-500">
                              {prefs.priorityMetric === 'profit'
                                ? `${item.marginPct}% margin · Value ${formatMoney(item.calc.grandTotal)}`
                                : `Profit ${formatMoney(item.calc.grossProfit)} · ${item.metrics.percent}% done`}
                            </div>
                          </>
                        ) : (
                          <div className="text-xs font-semibold text-slate-700">
                            {item.metrics.percent}% Complete
                          </div>
                        )}
                      </div>

                      <button
                        type="button"
                        onClick={() => onSelectJobAndTab(item.job.id, item.metrics.nextTab)}
                        className="inline-flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-800 hover:bg-slate-50 cursor-pointer min-h-[34px] whitespace-nowrap"
                      >
                        <span>{item.metrics.nextActionLabel}</span>
                        <ArrowRight className="w-3 h-3" />
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

      {/* Section 4: Job Lifecycle Tracker Matrix Table */}
      {prefs.showJobMatrix && (
        <div className="rounded-lg border border-slate-200 bg-white overflow-hidden">
          <div className="p-4 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-2.5 flex-1 min-w-[240px]">
              <div className="relative flex-1 max-w-xs">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Search job, client, RFQ/Tender ref..."
                  className="w-full min-h-[38px] rounded-lg border border-slate-200 pl-9 pr-3 py-1.5 text-xs text-slate-900 focus:border-slate-900 focus:outline-none"
                />
              </div>

              <div className="inline-flex items-center p-1 rounded-lg bg-slate-100">
                {(['ALL', 'RFQ', 'TENDER'] as const).map((t) => (
                  <button
                    key={t}
                    type="button"
                    onClick={() => setInquiryFilter(t)}
                    className={`px-2.5 py-1 rounded-md text-xs font-medium cursor-pointer transition-colors whitespace-nowrap ${
                      inquiryFilter === t
                        ? 'bg-white text-slate-900 shadow-sm'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    {t === 'ALL' ? 'All Types' : t === 'RFQ' ? 'RFQs' : 'Tenders'}
                  </button>
                ))}
              </div>

              <div className="inline-flex items-center p-1 rounded-lg bg-slate-100 overflow-x-auto">
                {(
                  [
                    ['ALL', 'All Stages'],
                    ['DRAFT', 'Draft'],
                    ['QT', 'QT'],
                    ['PI', 'PI'],
                    ['PO', 'PO'],
                    ['CI', 'CI'],
                    ['PL', 'PL'],
                    ['DN', 'DN'],
                  ] as [StageFilter, string][]
                ).map(([val, lbl]) => (
                  <button
                    key={val}
                    type="button"
                    onClick={() => setStageFilter(val)}
                    className={`px-2.5 py-1 rounded-md text-xs font-medium cursor-pointer transition-colors whitespace-nowrap ${
                      stageFilter === val
                        ? 'bg-white text-slate-900 shadow-sm'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    {lbl}
                  </button>
                ))}
              </div>
            </div>

            <span className="text-xs font-mono-tabular text-slate-500">
              Showing {matrixRows.length} of {scopedJobs.length}
            </span>
          </div>

          {jobSel.count > 0 && (
            <div className="no-print flex flex-wrap items-center gap-2 border-b border-slate-200 bg-slate-50 px-4 py-2.5 text-xs">
              <span className="font-semibold text-slate-900">{jobSel.count} job{jobSel.count === 1 ? '' : 's'} selected</span>
              <button type="button" onClick={() => exportSelectedJobs('xlsx')} className="rounded border border-slate-200 bg-white px-2.5 py-1.5 font-medium text-slate-800 hover:bg-slate-100 cursor-pointer">
                Export selected (.xlsx)
              </button>
              <button type="button" onClick={() => exportSelectedJobs('pdf')} className="rounded border border-slate-200 bg-white px-2.5 py-1.5 font-medium text-slate-800 hover:bg-slate-100 cursor-pointer">
                Export selected (.pdf)
              </button>
              <button type="button" onClick={jobSel.clear} className="ml-auto text-slate-600 underline cursor-pointer">Deselect all</button>
            </div>
          )}

          <div className="overflow-x-auto sheet-scroll">
            <table className="w-full border-collapse text-xs">
              <thead>
                <tr style={{ backgroundColor: theme.primary, color: theme.primaryText }}>
                  <th className="py-3 px-2 text-center w-8">
                    <input
                      type="checkbox"
                      aria-label="Select all jobs"
                      disabled={matrixRows.length === 0}
                      checked={jobSel.allChecked(matrixRows.map((m) => m.job.id))}
                      onChange={() => jobSel.toggleAll(matrixRows.map((m) => m.job.id))}
                    />
                  </th>
                  <th className="py-3 px-3 text-center font-semibold w-16">Seq</th>
                  <th className="py-3 px-3 text-left font-semibold min-w-[200px]">
                    Job &amp; Client
                  </th>
                  <th className="py-3 px-3 text-left font-semibold whitespace-nowrap">
                    Type · RFQ / Tender Ref
                  </th>
                  <th className="py-3 px-3 text-left font-semibold w-36">Completion</th>
                  <th className="py-3 px-2 text-center font-semibold w-20">Int. Q</th>
                  <th className="py-3 px-2 text-center font-semibold w-20">QT</th>
                  <th className="py-3 px-2 text-center font-semibold w-20">PI</th>
                  <th className="py-3 px-2 text-center font-semibold w-20">PO</th>
                  <th className="py-3 px-2 text-center font-semibold w-20">CI</th>
                  <th className="py-3 px-2 text-center font-semibold w-20">PL</th>
                  <th className="py-3 px-2 text-center font-semibold w-20">DN</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {matrixRows.length === 0 ? (
                  <tr>
                    <td colSpan={12} className="py-10 text-center text-slate-500">
                      No jobs match the selected filter.
                    </td>
                  </tr>
                ) : (
                  matrixRows.map((item, idx) => {
                    const { job, client, metrics } = item;
                    const isActive = job.id === activeJobId;
                    return (
                      <tr
                        key={job.id}
                        style={
                          isActive
                            ? { backgroundColor: theme.secondaryBg }
                            : idx % 2 === 1
                            ? { backgroundColor: theme.zebraRow }
                            : undefined
                        }
                        className="hover:bg-slate-50 transition-colors"
                      >
                        <td className="py-2.5 px-2 text-center">
                          <input type="checkbox" aria-label={`Select ${job.jobName}`} checked={jobSel.has(job.id)} onChange={() => jobSel.toggle(job.id)} />
                        </td>
                        <td className="py-2.5 px-3 text-center font-mono-tabular font-semibold text-slate-700">
                          {job.jobSequence
                            ? `#${String(job.jobSequence).padStart(3, '0')}`
                            : 'Draft'}
                        </td>
                        <td className="py-2.5 px-3">
                          <button
                            type="button"
                            onClick={() => onSelectJobAndTab(job.id, 'QT')}
                            className="text-left group cursor-pointer"
                          >
                            <div className="font-semibold text-slate-900 group-hover:underline">
                              {job.jobName}
                            </div>
                            <div className="text-[11px] text-slate-500">
                              {client} · {metrics.currentStageLabel}
                            </div>
                          </button>
                        </td>
                        <td className="py-2.5 px-3 whitespace-nowrap font-mono-tabular">
                          <span className="font-semibold text-slate-800">{job.inquiryType}</span>
                          {job.controls.rfqRef ? (
                            <span className="text-slate-500"> · {job.controls.rfqRef}</span>
                          ) : (
                            <span className="text-slate-400"> · —</span>
                          )}
                        </td>
                        <td className="py-2.5 px-3">
                          <div className="flex items-center gap-2">
                            <div className="flex-1 h-2 rounded-full bg-slate-200 overflow-hidden">
                              <div
                                className="h-full rounded-full"
                                style={{
                                  width: `${metrics.percent}%`,
                                  backgroundColor: metrics.isComplete ? '#15803D' : theme.primary,
                                }}
                              />
                            </div>
                            <span className="font-mono-tabular font-semibold text-slate-800 w-9 text-right">
                              {metrics.percent}%
                            </span>
                          </div>
                        </td>
                        <td className="py-2 px-1 text-center">
                          <button
                            type="button"
                            onClick={() => onSelectJobAndTab(job.id, 'INTERNAL_QUOTE')}
                            className={`w-full min-h-[32px] rounded px-1.5 py-1 inline-flex items-center justify-center gap-1 text-[11px] font-medium cursor-pointer ${
                              metrics.hasInternalQuote
                                ? 'text-emerald-800 hover:bg-emerald-50 font-semibold'
                                : 'text-slate-400 hover:bg-slate-100'
                            }`}
                          >
                            {metrics.hasInternalQuote ? (
                              <>
                                <Check className="w-3.5 h-3.5 text-emerald-700" />
                                <span>Ready</span>
                              </>
                            ) : (
                              '—'
                            )}
                          </button>
                        </td>
                        <td className="py-2 px-1 text-center">
                          {renderStageCell(job.id, 'QT', metrics.hasQT, job.issuedDocs.QT?.docNumber, true)}
                        </td>
                        <td className="py-2 px-1 text-center">
                          {renderStageCell(
                            job.id,
                            'PI',
                            metrics.hasPI,
                            job.issuedDocs.PI?.docNumber,
                            metrics.hasQT
                          )}
                        </td>
                        <td className="py-2 px-1 text-center">
                          {renderStageCell(
                            job.id,
                            'PO',
                            metrics.hasPO,
                            job.issuedDocs.PO?.docNumber,
                            metrics.hasEndUserPoGate
                          )}
                        </td>
                        <td className="py-2 px-1 text-center">
                          {renderStageCell(
                            job.id,
                            'CI',
                            metrics.hasCI,
                            job.issuedDocs.CI?.docNumber,
                            metrics.hasEndUserPoGate
                          )}
                        </td>
                        <td className="py-2 px-1 text-center">
                          {renderStageCell(
                            job.id,
                            'PL',
                            metrics.hasPL,
                            job.issuedDocs.PL?.docNumber,
                            metrics.hasSupplierShipmentGate
                          )}
                        </td>
                        <td className="py-2 px-1 text-center">
                          {renderStageCell(
                            job.id,
                            'DN',
                            metrics.hasDN,
                            job.issuedDocs.DN?.docNumber,
                            metrics.hasPL
                          )}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {pdfModalOpen && (
        <div className="no-print fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="w-full max-w-lg rounded-xl border border-slate-200 bg-white p-6 shadow-xl space-y-5">
            <div>
              <h3 className="text-base font-semibold text-slate-900">Dashboard PDF Report</h3>
              <p className="mt-1 text-xs text-slate-500">
                Choose the page layout. Margins are 42pt so no text or table column is clipped.
              </p>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              {(
                [
                  { id: 'landscape', title: 'A4 Landscape', hint: 'Recommended for the full job matrix' },
                  { id: 'portrait', title: 'A4 Portrait', hint: 'Stacked executive summary' },
                ] as { id: DashboardPdfOrientation; title: string; hint: string }[]
              ).map((opt) => (
                <button
                  key={opt.id}
                  type="button"
                  onClick={() => setPdfOrientation(opt.id)}
                  className={`rounded-lg border p-4 text-left cursor-pointer transition-colors ${
                    pdfOrientation === opt.id
                      ? 'border-slate-900 ring-1 ring-slate-900'
                      : 'border-slate-200 hover:bg-slate-50'
                  }`}
                >
                  <div className="text-sm font-semibold text-slate-900">{opt.title}</div>
                  <div className="mt-1 text-xs text-slate-500">{opt.hint}</div>
                </button>
              ))}
            </div>
            <div className="flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setPdfModalOpen(false)}
                className="rounded-lg border border-slate-200 px-4 py-2 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                disabled={exportingPdf}
                onClick={handleDownloadPdfReport}
                className="rounded-lg bg-slate-900 px-4 py-2 text-xs font-semibold text-white hover:bg-slate-800 cursor-pointer disabled:opacity-50"
              >
                {exportingPdf ? 'Preparing...' : 'Download PDF'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
