import React, { useMemo, useState } from 'react';
import {
  Activity,
  ArrowUpRight,
  Calculator,
  Copy,
  FileText,
  Lock,
  Radio,
  RefreshCw,
  Search,
  Unlock,
} from 'lucide-react';
import type {
  ExecutiveSnapshot,
  OrgRole,
  StreamConnectionState,
} from '../types/enterprise';
import type { WorkspaceTab } from '../types/procurement';
import { formatMoney } from '../utils/calculations';
import { getApiBaseUrl } from '../lib/apiClient';

interface ExecutiveLiveViewProps {
  snapshot: ExecutiveSnapshot | null;
  loading: boolean;
  error: string | null;
  streamState: StreamConnectionState;
  lastPulseAt: string | null;
  role: OrgRole;
  executiveSafeMode: boolean;
  onToggleSafeMode: () => void;
  onRefreshSnapshot: () => Promise<void>;
  onOpenJobInFullApp: (jobId: string, tab: WorkspaceTab) => void;
  onOpenOrgSeats: () => void;
  onNotify: (msg: string) => void;
}

function formatRelativeTime(iso: string): string {
  const diffSec = Math.max(0, Math.floor((Date.now() - new Date(iso).getTime()) / 1000));
  if (diffSec < 10) return 'Just now';
  if (diffSec < 60) return `${diffSec}s ago`;
  const min = Math.floor(diffSec / 60);
  if (min < 60) return `${min}m ago`;
  const hrs = Math.floor(min / 60);
  if (hrs < 24) return `${hrs}h ago`;
  return new Date(iso).toLocaleDateString();
}

export const ExecutiveLiveView: React.FC<ExecutiveLiveViewProps> = ({
  snapshot,
  loading,
  error,
  streamState,
  lastPulseAt,
  role,
  executiveSafeMode,
  onToggleSafeMode,
  onRefreshSnapshot,
  onOpenJobInFullApp,
  onOpenOrgSeats,
  onNotify,
}) => {
  const [search, setSearch] = useState('');
  const [inquiryFilter, setInquiryFilter] = useState<'ALL' | 'RFQ' | 'TENDER'>('ALL');
  const [stageFilter, setStageFilter] = useState<string>('ALL');
  const [perfPeriod, setPerfPeriod] = useState<'weekly' | 'monthly'>('weekly');
  const [showEndpointDrawer, setShowEndpointDrawer] = useState(false);
  const [refreshing, setRefreshing] = useState(false);

  const filteredJobs = useMemo(() => {
    if (!snapshot) return [];
    const q = search.trim().toLowerCase();
    return snapshot.jobs.filter((j) => {
      if (inquiryFilter !== 'ALL' && j.inquiryType !== inquiryFilter) return false;
      if (stageFilter !== 'ALL') {
        if (stageFilter === 'COMPLETE' && !j.isComplete) return false;
        if (stageFilter === 'ACTIVE' && j.isComplete) return false;
        if (
          ['QT', 'PI', 'PO', 'CI', 'PL', 'DN'].includes(stageFilter) &&
          !j.issuedPrefixes.includes(stageFilter as 'QT')
        ) {
          return false;
        }
      }
      if (!q) return true;
      return (
        j.jobName.toLowerCase().includes(q) ||
        j.client.toLowerCase().includes(q) ||
        j.supplier.toLowerCase().includes(q) ||
        j.preparedBy.toLowerCase().includes(q)
      );
    });
  }, [snapshot, search, inquiryFilter, stageFilter]);

  const handleManualRefresh = async () => {
    setRefreshing(true);
    try {
      await onRefreshSnapshot();
      onNotify('Executive snapshot refreshed from PostgreSQL.');
    } finally {
      setRefreshing(false);
    }
  };

  const apiOrigin =
    getApiBaseUrl() || (typeof window !== 'undefined' ? window.location.origin : '');
  const orgId = snapshot?.organizationId || ':orgId';
  const sseEndpointExample = `${apiOrigin}/api/orgs/${orgId}/executive/stream?ticket=<one-time-ticket>`;
  const snapshotEndpointExample = `${apiOrigin}/api/orgs/${orgId}/executive/snapshot`;

  const handleCopyEndpoint = async () => {
    const spec = [
      `# Live Executive Dashboard Streaming Endpoints`,
      `1. Request one-time stream ticket (Bearer Firebase ID token):`,
      `   POST ${apiOrigin}/api/orgs/${orgId}/stream-ticket`,
      `2. Subscribe to real-time SSE updates:`,
      `   GET  ${sseEndpointExample}`,
      `3. Pull on-demand JSON snapshot:`,
      `   GET  ${snapshotEndpointExample}`,
      `4. WebSocket real-time channel:`,
      `   WSS  ${apiOrigin.replace(/^http/, 'ws')}/ws`,
    ].join('\n');
    try {
      await navigator.clipboard.writeText(spec);
      onNotify('Copied Executive Streaming Endpoint specification to clipboard.');
    } catch {
      onNotify('Could not copy automatically. Select the endpoint text manually.');
    }
  };

  if (loading && !snapshot) {
    return (
      <div className="space-y-6">
        <div className="h-16 rounded-xl border border-slate-200 bg-white animate-pulse" />
        <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-4">
          {[0, 1, 2, 3].map((i) => (
            <div key={i} className="h-28 rounded-xl border border-slate-200 bg-white animate-pulse" />
          ))}
        </div>
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          <div className="lg:col-span-7 h-80 rounded-xl border border-slate-200 bg-white animate-pulse" />
          <div className="lg:col-span-5 h-80 rounded-xl border border-slate-200 bg-white animate-pulse" />
        </div>
      </div>
    );
  }

  if (error && !snapshot) {
    return (
      <div className="rounded-xl border border-red-200 bg-white p-8 text-center space-y-3">
        <div className="text-base font-semibold text-slate-900">
          Executive stream could not be loaded
        </div>
        <p className="text-sm text-slate-600">{error}</p>
        <button
          type="button"
          onClick={handleManualRefresh}
          className="inline-flex items-center gap-2 rounded-lg bg-[#1B365D] text-white px-4 py-2 text-xs font-semibold cursor-pointer"
        >
          <RefreshCw className="w-3.5 h-3.5" />
          <span>Retry Connection</span>
        </button>
      </div>
    );
  }

  if (!snapshot) return null;

  const perfBuckets =
    perfPeriod === 'weekly' ? snapshot.weeklyPerformance : snapshot.monthlyPerformance;
  const maxPerfValue = Math.max(1, ...perfBuckets.map((b) => b.value));

  const streamLabel =
    streamState === 'live'
      ? 'Live Stream Connected'
      : streamState === 'reconnecting'
      ? 'Reconnecting Stream'
      : streamState === 'connecting'
      ? 'Connecting Stream'
      : 'Offline Snapshot';

  const streamDotClass =
    streamState === 'live'
      ? 'bg-emerald-500'
      : streamState === 'reconnecting' || streamState === 'connecting'
      ? 'bg-amber-500'
      : 'bg-slate-400';

  return (
    <div className="space-y-6">
      {/* Executive Header Bar */}
      <div className="bg-white border border-slate-200 rounded-xl px-5 py-4 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
        <div className="space-y-1 min-w-0">
          <div className="flex flex-wrap items-center gap-3">
            <h2 className="font-display text-xl font-bold text-slate-900">
              Executive Live Command — {snapshot.organizationName}
            </h2>
            <span className="inline-flex items-center gap-1.5 text-xs font-medium text-slate-700">
              <span className={`w-2 h-2 rounded-full ${streamDotClass}`} aria-hidden="true" />
              <span>{streamLabel}</span>
            </span>
          </div>
          <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500 font-mono tabular-nums">
            <span>Updated {formatRelativeTime(lastPulseAt || snapshot.generatedAt)}</span>
            <span aria-hidden="true">·</span>
            <span>{snapshot.activeListeners} live viewer{snapshot.activeListeners === 1 ? '' : 's'}</span>
            <span aria-hidden="true">·</span>
            <button
              type="button"
              onClick={onOpenOrgSeats}
              className="hover:text-slate-900 underline decoration-slate-300 underline-offset-2 cursor-pointer"
            >
              {snapshot.seatSummary.used}/{snapshot.seatSummary.limit} seats active ({snapshot.seatSummary.planTier})
            </button>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2.5 shrink-0">
          {(role === 'CEO_EXECUTIVE' || role === 'ORG_ADMIN') && (
            <button
              type="button"
              onClick={onToggleSafeMode}
              className={`inline-flex items-center gap-1.5 px-3 py-2 rounded-lg border text-xs font-semibold transition-colors cursor-pointer whitespace-nowrap ${
                executiveSafeMode
                  ? 'border-amber-300 bg-amber-50 text-amber-900 hover:bg-amber-100'
                  : 'border-emerald-300 bg-emerald-50 text-emerald-900 hover:bg-emerald-100'
              }`}
              title="Toggle read-only browsing protection across the full application"
            >
              {executiveSafeMode ? (
                <>
                  <Lock className="w-3.5 h-3.5" />
                  <span>Safe Read-Only Mode: ON</span>
                </>
              ) : (
                <>
                  <Unlock className="w-3.5 h-3.5" />
                  <span>Full Edit Access: ON</span>
                </>
              )}
            </button>
          )}

          <button
            type="button"
            onClick={() => setShowEndpointDrawer((v) => !v)}
            className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-medium text-slate-700 cursor-pointer whitespace-nowrap"
          >
            <Radio className="w-3.5 h-3.5 text-slate-500" />
            <span>Stream Endpoint</span>
          </button>

          <button
            type="button"
            disabled={refreshing}
            onClick={handleManualRefresh}
            className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg bg-[#1B365D] hover:bg-[#152a48] text-white text-xs font-semibold cursor-pointer whitespace-nowrap disabled:opacity-60"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${refreshing ? 'animate-spin' : ''}`} />
            <span>Refresh</span>
          </button>
        </div>
      </div>

      {/* Collapsible Live Stream Endpoint Specification */}
      {showEndpointDrawer && (
        <div className="bg-slate-900 text-slate-100 rounded-xl p-5 space-y-3 border border-slate-800">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <h3 className="text-sm font-semibold text-white">
                CEO Live Dashboard Streaming Endpoints (SSE &amp; WebSocket)
              </h3>
              <p className="text-xs text-slate-400 mt-0.5">
                Every job edit, document issuance, or workflow gate change pushes an immediate delta
                over these authenticated organization-scoped channels.
              </p>
            </div>
            <button
              type="button"
              onClick={handleCopyEndpoint}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-xs font-medium text-white cursor-pointer whitespace-nowrap"
            >
              <Copy className="w-3.5 h-3.5" />
              <span>Copy Spec</span>
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs font-mono">
            <div className="rounded-lg bg-slate-950/70 border border-slate-800 p-3 space-y-1">
              <div className="text-slate-400 font-sans font-medium">Real-Time SSE Stream</div>
              <div className="text-emerald-300 break-all">GET {sseEndpointExample}</div>
            </div>
            <div className="rounded-lg bg-slate-950/70 border border-slate-800 p-3 space-y-1">
              <div className="text-slate-400 font-sans font-medium">On-Demand JSON Snapshot</div>
              <div className="text-sky-300 break-all">GET {snapshotEndpointExample}</div>
            </div>
          </div>
        </div>
      )}

      {/* 4-Column KPI Strip */}
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-4">
        <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-1.5">
          <div className="text-xs font-medium text-slate-500">Total Pipeline Value</div>
          <div className="font-mono text-2xl font-bold text-slate-900 tabular-nums">
            {formatMoney(snapshot.kpis.totalPipelineValue, 'USD')}
          </div>
          <div className="text-xs text-slate-500 font-mono tabular-nums flex items-center gap-2">
            <span>RFQ {formatMoney(snapshot.kpis.rfqValue, 'USD')}</span>
            <span aria-hidden="true">·</span>
            <span>Tender {formatMoney(snapshot.kpis.tenderValue, 'USD')}</span>
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-1.5">
          <div className="text-xs font-medium text-slate-500">Estimated Gross Profit</div>
          <div className="font-mono text-2xl font-bold text-emerald-700 tabular-nums">
            {formatMoney(snapshot.kpis.totalEstimatedProfit, 'USD')}
          </div>
          <div className="text-xs text-slate-500 font-mono tabular-nums">
            Average margin {snapshot.kpis.avgMarginPct.toFixed(1)}% across active deals
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-1.5">
          <div className="text-xs font-medium text-slate-500">Active Inquiries</div>
          <div className="font-mono text-2xl font-bold text-slate-900 tabular-nums">
            {snapshot.kpis.activeJobs}
            <span className="text-sm font-normal text-slate-400 ml-1.5">
              / {snapshot.kpis.totalJobs} total
            </span>
          </div>
          <div className="text-xs text-slate-500 font-mono tabular-nums flex items-center gap-2">
            <span>{snapshot.kpis.rfqCount} RFQs</span>
            <span aria-hidden="true">·</span>
            <span>{snapshot.kpis.tenderCount} Tenders</span>
            <span aria-hidden="true">·</span>
            <span>{snapshot.kpis.completedJobs} Done</span>
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-1.5">
          <div className="text-xs font-medium text-slate-500">Avg Workflow Progress</div>
          <div className="font-mono text-2xl font-bold text-slate-900 tabular-nums">
            {snapshot.kpis.avgCompletionPct}%
          </div>
          <div className="text-xs text-slate-500 font-mono tabular-nums">
            {snapshot.kpis.docsIssuedCount} issued Master Registry document
            {snapshot.kpis.docsIssuedCount === 1 ? '' : 's'}
          </div>
        </div>
      </div>

      {/* 2-Column Split: Funnel & Performance (Left) + Real-Time Activity Stream (Right) */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
        <div className="lg:col-span-7 space-y-6">
          {/* 9-Stage Completion Funnel */}
          <section className="bg-white border border-slate-200 rounded-xl p-5 space-y-4">
            <div className="flex items-center justify-between gap-4">
              <div>
                <h3 className="font-display text-base font-bold text-slate-900">
                  9-Stage Procurement Funnel
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  Click a trade document stage to filter the live job matrix below.
                </p>
              </div>
              {stageFilter !== 'ALL' && (
                <button
                  type="button"
                  onClick={() => setStageFilter('ALL')}
                  className="text-xs font-semibold text-[#1B365D] hover:underline cursor-pointer"
                >
                  Clear filter ({stageFilter})
                </button>
              )}
            </div>

            <div className="space-y-2.5">
              {snapshot.stageFunnel.map((stage) => {
                const active = stage.filterTarget && stageFilter === stage.filterTarget;
                return (
                  <button
                    key={stage.id}
                    type="button"
                    disabled={!stage.filterTarget}
                    onClick={() =>
                      stage.filterTarget &&
                      setStageFilter(active ? 'ALL' : stage.filterTarget)
                    }
                    className={`w-full text-left rounded-lg px-2.5 py-1.5 transition-colors ${
                      stage.filterTarget ? 'cursor-pointer hover:bg-slate-50' : 'cursor-default'
                    } ${active ? 'bg-slate-100' : ''}`}
                  >
                    <div className="flex items-center justify-between text-xs mb-1">
                      <span className="font-medium text-slate-800">{stage.label}</span>
                      <span className="font-mono tabular-nums text-slate-600">
                        {stage.count} ({stage.pct}%)
                      </span>
                    </div>
                    <div className="h-2 w-full rounded-full bg-slate-100 overflow-hidden">
                      <div
                        className="h-full rounded-full bg-[#1B365D] transition-all duration-300"
                        style={{ width: `${Math.min(100, stage.pct)}%` }}
                      />
                    </div>
                  </button>
                );
              })}
            </div>
          </section>

          {/* Weekly / Monthly Pipeline Volume */}
          <section className="bg-white border border-slate-200 rounded-xl p-5 space-y-4">
            <div className="flex items-center justify-between gap-4">
              <div>
                <h3 className="font-display text-base font-bold text-slate-900">
                  Pipeline Volume &amp; Document Issuance
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  Aggregated live from organization jobs and Master Registry timestamps.
                </p>
              </div>
              <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg">
                <button
                  type="button"
                  onClick={() => setPerfPeriod('weekly')}
                  className={`px-3 py-1 text-xs font-medium rounded-md cursor-pointer transition-colors ${
                    perfPeriod === 'weekly'
                      ? 'bg-white text-slate-900 shadow-sm'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Weekly
                </button>
                <button
                  type="button"
                  onClick={() => setPerfPeriod('monthly')}
                  className={`px-3 py-1 text-xs font-medium rounded-md cursor-pointer transition-colors ${
                    perfPeriod === 'monthly'
                      ? 'bg-white text-slate-900 shadow-sm'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Monthly
                </button>
              </div>
            </div>

            <div className="grid grid-cols-6 sm:grid-cols-8 gap-2 items-end h-36 pt-2">
              {perfBuckets.map((b) => {
                const heightPct = b.value > 0 ? Math.max(10, Math.round((b.value / maxPerfValue) * 100)) : 4;
                return (
                  <div key={b.key} className="flex flex-col items-center justify-end h-full gap-1.5">
                    <div className="text-[10px] font-mono tabular-nums text-slate-500">
                      {b.docsIssued > 0 ? `${b.docsIssued}d` : ''}
                    </div>
                    <div className="w-full flex-1 flex items-end">
                      <div
                        title={`${b.label}: ${formatMoney(b.value, 'USD')} pipeline · ${b.docsIssued} docs issued`}
                        className="w-full rounded-t-md bg-[#1B365D]/85 hover:bg-[#1B365D] transition-all"
                        style={{ height: `${heightPct}%` }}
                      />
                    </div>
                    <div className="text-[11px] font-mono text-slate-500 whitespace-nowrap">
                      {b.label}
                    </div>
                  </div>
                );
              })}
            </div>
          </section>
        </div>

        {/* Right: Live Organization Activity Stream */}
        <section className="lg:col-span-5 bg-white border border-slate-200 rounded-xl p-5 space-y-4">
          <div className="flex items-center justify-between gap-2">
            <div>
              <h3 className="font-display text-base font-bold text-slate-900 flex items-center gap-2">
                <Activity className="w-4 h-4 text-slate-500" />
                <span>Live Organization Activity</span>
              </h3>
              <p className="text-xs text-slate-500 mt-0.5">
                Real-time audit feed across all team seats.
              </p>
            </div>
          </div>

          {snapshot.recentActivity.length === 0 ? (
            <div className="py-12 text-center text-xs text-slate-500 border border-dashed border-slate-200 rounded-lg">
              Activity will stream here as procurement officers create jobs, update pricing, and
              issue trade documents.
            </div>
          ) : (
            <div className="divide-y divide-slate-100 max-h-[540px] overflow-y-auto pr-1">
              {snapshot.recentActivity.map((ev) => (
                <div key={ev.id} className="py-3 flex items-start justify-between gap-3">
                  <div className="min-w-0 space-y-0.5">
                    <div className="text-xs font-medium text-slate-900 leading-snug">
                      {ev.summary}
                    </div>
                    <div className="text-[11px] text-slate-500 flex flex-wrap items-center gap-1.5">
                      <span>{ev.actorName}</span>
                      <span aria-hidden="true">·</span>
                      <span className="font-mono tabular-nums">{formatRelativeTime(ev.at)}</span>
                      {ev.docNumber && (
                        <>
                          <span aria-hidden="true">·</span>
                          <span className="font-mono text-slate-700">{ev.docNumber}</span>
                        </>
                      )}
                    </div>
                  </div>

                  {ev.jobId && (
                    <button
                      type="button"
                      onClick={() => onOpenJobInFullApp(ev.jobId, 'QT')}
                      className="shrink-0 inline-flex items-center gap-1 text-[11px] font-semibold text-[#1B365D] hover:underline cursor-pointer"
                    >
                      <span>Open</span>
                      <ArrowUpRight className="w-3 h-3" />
                    </button>
                  )}
                </div>
              ))}
            </div>
          )}
        </section>
      </div>

      {/* Full-Width Live Job Matrix with Full-App Drill-Down */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
          <div>
            <h3 className="font-display text-base font-bold text-slate-900">
              Live Organization Deals — Click Any Row to Enter Full App
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Showing {filteredJobs.length} of {snapshot.jobs.length} jobs across your organization.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            <div className="relative">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Filter client, job, owner..."
                className="pl-8 pr-3 py-1.5 rounded-lg border border-slate-200 text-xs text-slate-900 focus:border-[#1B365D] focus:outline-none w-52"
              />
            </div>

            <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg">
              {(['ALL', 'RFQ', 'TENDER'] as const).map((kind) => (
                <button
                  key={kind}
                  type="button"
                  onClick={() => setInquiryFilter(kind)}
                  className={`px-2.5 py-1 text-xs font-medium rounded-md cursor-pointer transition-colors ${
                    inquiryFilter === kind
                      ? 'bg-white text-slate-900 shadow-sm'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  {kind === 'ALL' ? 'All' : kind}
                </button>
              ))}
            </div>
          </div>
        </div>

        {filteredJobs.length === 0 ? (
          <div className="p-10 text-center text-xs text-slate-500">
            No procurement jobs match your active filters.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 text-[11px] font-semibold text-slate-500">
                  <th className="py-2.5 px-4">Job &amp; Client</th>
                  <th className="py-2.5 px-3">Owner</th>
                  <th className="py-2.5 px-3">Stage</th>
                  <th className="py-2.5 px-3 text-right">Pipeline Value</th>
                  <th className="py-2.5 px-3 text-right">Est. Profit</th>
                  <th className="py-2.5 px-3 text-right">Margin</th>
                  <th className="py-2.5 px-4 text-right">Full-App Drill-Down</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-xs">
                {filteredJobs.map((row) => (
                  <tr key={row.jobId} className="hover:bg-slate-50/80 transition-colors">
                    <td className="py-3 px-4">
                      <div className="font-semibold text-slate-900">{row.jobName}</div>
                      <div className="text-[11px] text-slate-500 mt-0.5 flex items-center gap-1.5">
                        <span>{row.client}</span>
                        <span aria-hidden="true">·</span>
                        <span>{row.inquiryType}</span>
                        {row.jobSequence ? (
                          <>
                            <span aria-hidden="true">·</span>
                            <span className="font-mono tabular-nums">
                              #{String(row.jobSequence).padStart(3, '0')}
                            </span>
                          </>
                        ) : null}
                      </div>
                    </td>
                    <td className="py-3 px-3 text-slate-600 whitespace-nowrap">{row.preparedBy}</td>
                    <td className="py-3 px-3">
                      <div className="text-slate-800 font-medium">{row.currentStageLabel}</div>
                      <div className="text-[11px] font-mono tabular-nums text-slate-500">
                        {row.completionPct}% complete
                        {row.issuedPrefixes.length > 0
                          ? ` · ${row.issuedPrefixes.join(', ')}`
                          : ''}
                      </div>
                    </td>
                    <td className="py-3 px-3 text-right font-mono tabular-nums font-semibold text-slate-900 whitespace-nowrap">
                      {formatMoney(row.pipelineValue, row.currency)}
                    </td>
                    <td className="py-3 px-3 text-right font-mono tabular-nums font-semibold text-emerald-700 whitespace-nowrap">
                      {formatMoney(row.estimatedProfit, row.currency)}
                    </td>
                    <td className="py-3 px-3 text-right font-mono tabular-nums text-slate-700 whitespace-nowrap">
                      {row.marginPct.toFixed(1)}%
                    </td>
                    <td className="py-3 px-4 text-right whitespace-nowrap">
                      <div className="inline-flex items-center gap-1.5">
                        <button
                          type="button"
                          onClick={() => onOpenJobInFullApp(row.jobId, 'PRICE_CALC')}
                          className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-md border border-slate-200 bg-white hover:bg-slate-100 text-[11px] font-medium text-slate-700 cursor-pointer"
                          title="Inspect cost & margin calculation"
                        >
                          <Calculator className="w-3 h-3" />
                          <span>Price Calc</span>
                        </button>
                        <button
                          type="button"
                          onClick={() => onOpenJobInFullApp(row.jobId, 'QT')}
                          className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-md bg-[#1B365D] hover:bg-[#152a48] text-white text-[11px] font-semibold cursor-pointer"
                          title="Open full job workspace"
                        >
                          <FileText className="w-3 h-3" />
                          <span>Open Job</span>
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>
    </div>
  );
};
