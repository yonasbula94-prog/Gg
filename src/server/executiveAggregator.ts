import { calculateJob, getJobCompletionMetrics, getResolvedEndUser, round2 } from '../utils/calculations.ts';
import type { DocPrefix, ProcurementJob, RegistryEntry } from '../types/procurement.ts';
import type {
  ExecutiveActivityEvent,
  ExecutiveJobRow,
  ExecutivePerformanceBucket,
  ExecutiveSnapshot,
  ExecutiveStageMetric,
  LicensePlanTier,
} from '../types/enterprise.ts';

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

export interface AuditLogInput {
  id: string;
  at: string;
  actorName: string;
  eventType: string;
  jobId: string;
  jobName: string;
  docNumber: string;
  prefix: string;
  revision: number;
  reason: string;
  field: string;
}

function summarizeAuditEvent(entry: AuditLogInput): string {
  switch (entry.eventType) {
    case 'DOCUMENT_ISSUED':
      return `Issued ${entry.docNumber || entry.prefix} on ${entry.jobName || 'job'}`;
    case 'DOCUMENT_VOIDED':
      return `Voided ${entry.docNumber}${entry.reason ? ` (${entry.reason})` : ''}`;
    case 'DOCUMENT_AMENDMENT':
      return `Saved ${entry.docNumber} Rev ${entry.revision} — ${entry.field || entry.reason || 'updated'}`;
    case 'JOB_CREATED':
      return `Created new job "${entry.jobName}"`;
    case 'JOB_UPDATED':
      return `Updated "${entry.jobName}"${entry.reason ? ` — ${entry.reason}` : ''}`;
    case 'JOB_DELETED':
      return `Removed job "${entry.jobName}"`;
    case 'SEAT_CHANGED':
      return entry.reason || 'Updated organization team seat';
    case 'LICENSE_ACTIVATED':
      return entry.reason || 'Activated organization license key';
    default:
      return entry.reason || `${entry.eventType} on ${entry.jobName || 'workspace'}`;
  }
}

export function buildExecutiveSnapshot(params: {
  organizationId: string;
  organizationName: string;
  planTier: LicensePlanTier;
  seatUsed: number;
  seatLimit: number;
  licenseExpiresAt: Date;
  activeListeners: number;
  jobs: ProcurementJob[];
  registry: RegistryEntry[];
  recentAudit: AuditLogInput[];
}): ExecutiveSnapshot {
  const now = new Date();
  const daysUntilExpiry = Math.ceil((params.licenseExpiresAt.getTime() - now.getTime()) / 86400000);

  const enriched = params.jobs.map((job) => {
    const calc = calculateJob(job);
    const metrics = getJobCompletionMetrics(job);
    const client = getResolvedEndUser(job) || 'Unassigned Client';
    const marginPct =
      calc.goodsSubtotalExVat > 0
        ? Math.round((calc.grossProfit / calc.goodsSubtotalExVat) * 1000) / 10
        : 0;
    const issuedPrefixes = (Object.keys(job.issuedDocs || {}) as DocPrefix[]).filter(
      (p) => Boolean(job.issuedDocs[p] && !job.issuedDocs[p]?.isVoided)
    );
    return { job, calc, metrics, client, marginPct, issuedPrefixes };
  });

  const activeItems = enriched.filter((e) => !e.metrics.isComplete);
  const completedItems = enriched.filter((e) => e.metrics.isComplete);
  const rfqItems = enriched.filter((e) => e.job.inquiryType !== 'TENDER');
  const tenderItems = enriched.filter((e) => e.job.inquiryType === 'TENDER');

  const rfqValue = round2(rfqItems.reduce((sum, e) => sum + e.calc.grandTotal, 0));
  const tenderValue = round2(tenderItems.reduce((sum, e) => sum + e.calc.grandTotal, 0));
  const totalPipelineValue = round2(rfqValue + tenderValue);
  const totalEstimatedProfit = round2(enriched.reduce((sum, e) => sum + e.calc.grossProfit, 0));
  const totalGoodsExVat = enriched.reduce((sum, e) => sum + e.calc.goodsSubtotalExVat, 0);
  const avgMarginPct =
    totalGoodsExVat > 0 ? Math.round((totalEstimatedProfit / totalGoodsExVat) * 1000) / 10 : 0;
  const avgCompletionPct =
    enriched.length > 0
      ? Math.round(enriched.reduce((sum, e) => sum + e.metrics.percent, 0) / enriched.length)
      : 0;

  const activeRegistry = params.registry.filter((r) => r.status === 'ISSUED');

  const denom = Math.max(1, enriched.length);
  const countPred = (fn: (e: (typeof enriched)[number]) => boolean) => enriched.filter(fn).length;
  const rawStages: { id: string; label: string; count: number; filterTarget?: string }[] = [
    { id: 'INT', label: 'Internal Supplier Quote', count: countPred((e) => e.metrics.hasInternalQuote) },
    { id: 'QT', label: 'Quotation (QT) Issued', count: countPred((e) => e.metrics.hasQT), filterTarget: 'QT' },
    { id: 'PI', label: 'Proforma Invoice (PI) Issued', count: countPred((e) => e.metrics.hasPI), filterTarget: 'PI' },
    { id: 'GATE_PO', label: 'Client PO Confirmed', count: countPred((e) => e.metrics.hasEndUserPoGate) },
    { id: 'PO', label: 'Supplier PO Issued', count: countPred((e) => e.metrics.hasPO), filterTarget: 'PO' },
    { id: 'CI', label: 'Commercial Invoice (CI) Issued', count: countPred((e) => e.metrics.hasCI), filterTarget: 'CI' },
    {
      id: 'GATE_PL',
      label: 'Supplier Paid & PL Confirmed',
      count: countPred((e) => e.metrics.hasSupplierShipmentGate),
    },
    { id: 'PL', label: 'Packing List (PL) Issued', count: countPred((e) => e.metrics.hasPL), filterTarget: 'PL' },
    { id: 'DN', label: 'Delivery Note (DN) Issued', count: countPred((e) => e.metrics.hasDN), filterTarget: 'DN' },
  ];

  const stageFunnel: ExecutiveStageMetric[] = rawStages.map((s) => ({
    ...s,
    pct: enriched.length === 0 ? 0 : Math.round((s.count / denom) * 100),
  }));

  // Build last 8 weekly buckets
  const weeklyPerformance: ExecutivePerformanceBucket[] = [];
  for (let i = 7; i >= 0; i--) {
    const dt = new Date(now.getTime() - i * 7 * 86400000);
    const key = getIsoWeekKey(dt);
    weeklyPerformance.push({
      key,
      label: key.slice(5),
      jobsCreated: 0,
      docsIssued: 0,
      value: 0,
      profit: 0,
    });
  }
  const weeklyMap = new Map(weeklyPerformance.map((b) => [b.key, b]));

  // Build last 6 monthly buckets
  const monthlyPerformance: ExecutivePerformanceBucket[] = [];
  for (let i = 5; i >= 0; i--) {
    const dt = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const key = getMonthKey(dt);
    const label = dt.toLocaleString('en-US', { month: 'short', year: '2-digit' });
    monthlyPerformance.push({
      key,
      label,
      jobsCreated: 0,
      docsIssued: 0,
      value: 0,
      profit: 0,
    });
  }
  const monthlyMap = new Map(monthlyPerformance.map((b) => [b.key, b]));

  for (const item of enriched) {
    const dt = new Date(item.job.createdAt);
    if (!Number.isNaN(dt.getTime())) {
      const wb = weeklyMap.get(getIsoWeekKey(dt));
      if (wb) {
        wb.jobsCreated += 1;
        wb.value = round2(wb.value + item.calc.grandTotal);
        wb.profit = round2(wb.profit + item.calc.grossProfit);
      }
      const mb = monthlyMap.get(getMonthKey(dt));
      if (mb) {
        mb.jobsCreated += 1;
        mb.value = round2(mb.value + item.calc.grandTotal);
        mb.profit = round2(mb.profit + item.calc.grossProfit);
      }
    }
  }

  for (const reg of activeRegistry) {
    const dt = new Date(reg.issuedAt);
    if (!Number.isNaN(dt.getTime())) {
      const wb = weeklyMap.get(getIsoWeekKey(dt));
      if (wb) wb.docsIssued += 1;
      const mb = monthlyMap.get(getMonthKey(dt));
      if (mb) mb.docsIssued += 1;
    }
  }

  const executiveJobs: ExecutiveJobRow[] = enriched
    .map((e) => ({
      jobId: e.job.id,
      jobName: e.job.jobName,
      jobSequence: e.job.jobSequence,
      inquiryType: e.job.inquiryType,
      client: e.client,
      supplier: e.job.controls.supplierName || e.job.supplierQuote?.supplierName || '—',
      currency: e.job.controls.currency || 'USD',
      pipelineValue: round2(e.calc.grandTotal),
      estimatedProfit: round2(e.calc.grossProfit),
      marginPct: e.marginPct,
      completionPct: e.metrics.percent,
      currentStageLabel: e.metrics.currentStageLabel,
      nextTab: e.metrics.nextTab,
      nextActionLabel: e.metrics.nextActionLabel,
      isComplete: e.metrics.isComplete,
      preparedBy: e.job.preparedBy || 'Procurement Desk',
      updatedAt: e.job.updatedAt,
      issuedPrefixes: e.issuedPrefixes,
    }))
    .sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime());

  const recentActivity: ExecutiveActivityEvent[] = params.recentAudit.slice(0, 30).map((a) => ({
    id: a.id,
    at: a.at,
    actorName: a.actorName || 'Team Member',
    eventType: a.eventType,
    jobId: a.jobId,
    jobName: a.jobName,
    docNumber: a.docNumber,
    prefix: a.prefix,
    summary: summarizeAuditEvent(a),
  }));

  return {
    organizationId: params.organizationId,
    organizationName: params.organizationName,
    generatedAt: now.toISOString(),
    activeListeners: params.activeListeners,
    seatSummary: {
      used: params.seatUsed,
      limit: params.seatLimit,
      planTier: params.planTier,
      daysUntilExpiry,
    },
    kpis: {
      totalJobs: enriched.length,
      activeJobs: activeItems.length,
      completedJobs: completedItems.length,
      rfqCount: rfqItems.length,
      tenderCount: tenderItems.length,
      rfqValue,
      tenderValue,
      totalPipelineValue,
      totalEstimatedProfit,
      avgMarginPct,
      avgCompletionPct,
      docsIssuedCount: activeRegistry.length,
    },
    stageFunnel,
    weeklyPerformance,
    monthlyPerformance,
    jobs: executiveJobs,
    recentActivity,
  };
}
