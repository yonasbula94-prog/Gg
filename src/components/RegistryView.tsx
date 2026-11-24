import React, { useMemo, useRef, useState } from 'react';
import {
  ArrowDown,
  ArrowUp,
  ArrowUpDown,
  Ban,
  FileSpreadsheet,
  RotateCcw,
  Search,
  Upload,
} from 'lucide-react';
import { DOC_META } from '../constants/themes';
import {
  CompanyProfile,
  ConfirmDialogRequest,
  DocPrefix,
  RegistryEntry,
  ThemePalette,
} from '../types/procurement';
import { formatMoney } from '../utils/calculations';
import { exportMasterRegistryXlsx, importWorkbookFromXlsx } from '../utils/excelEngine';

type RegistrySortKey =
  | 'sequence'
  | 'docNumber'
  | 'prefix'
  | 'customer'
  | 'jobName'
  | 'rfqRef'
  | 'amount'
  | 'issuedAt';

interface RegistryViewProps {
  entries: RegistryEntry[];
  nextSeq: number;
  theme: ThemePalette;
  company: CompanyProfile;
  onVoidEntry: (docNumber: string, reason: string) => void;
  onSyncRegistryEntries: (entries: RegistryEntry[]) => void;
  onResetRegistry: () => void;
  onExportAuditLog: () => void;
  onRequestConfirm: (req: ConfirmDialogRequest) => void;
  onNotify: (msg: string) => void;
}

export const RegistryView: React.FC<RegistryViewProps> = ({
  entries,
  nextSeq,
  theme,
  company,
  onVoidEntry,
  onSyncRegistryEntries,
  onResetRegistry,
  onExportAuditLog,
  onRequestConfirm,
  onNotify,
}) => {
  const [search, setSearch] = useState('');
  const [prefixFilter, setPrefixFilter] = useState<'ALL' | DocPrefix>('ALL');
  const [statusFilter, setStatusFilter] = useState<'ALL' | 'ISSUED' | 'VOID'>('ALL');
  const [sortKey, setSortKey] = useState<RegistrySortKey>('sequence');
  const [sortDir, setSortDir] = useState<'asc' | 'desc'>('desc');
  const [voidingDoc, setVoidingDoc] = useState<string | null>(null);
  const [voidReason, setVoidReason] = useState('');
  const [exporting, setExporting] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const handleToggleSort = (key: RegistrySortKey) => {
    if (sortKey === key) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortKey(key);
      setSortDir(key === 'sequence' || key === 'issuedAt' || key === 'amount' ? 'desc' : 'asc');
    }
  };

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    const list = entries.filter((e) => {
      if (prefixFilter !== 'ALL' && e.prefix !== prefixFilter) return false;
      if (statusFilter !== 'ALL' && e.status !== statusFilter) return false;
      if (!q) return true;
      return (
        e.docNumber.toLowerCase().includes(q) ||
        e.customer.toLowerCase().includes(q) ||
        e.jobName.toLowerCase().includes(q) ||
        (e.rfqRef || '').toLowerCase().includes(q) ||
        DOC_META[e.prefix].shortName.toLowerCase().includes(q)
      );
    });

    const dir = sortDir === 'asc' ? 1 : -1;
    return [...list].sort((a, b) => {
      switch (sortKey) {
        case 'sequence':
          if (a.sequence !== b.sequence) return (a.sequence - b.sequence) * dir;
          return a.docNumber.localeCompare(b.docNumber) * dir;
        case 'docNumber':
          return a.docNumber.localeCompare(b.docNumber) * dir;
        case 'prefix':
          return DOC_META[a.prefix].shortName.localeCompare(DOC_META[b.prefix].shortName) * dir;
        case 'customer':
          return a.customer.localeCompare(b.customer) * dir;
        case 'jobName':
          return a.jobName.localeCompare(b.jobName) * dir;
        case 'rfqRef':
          return (a.rfqRef || '').localeCompare(b.rfqRef || '') * dir;
        case 'amount':
          return (a.amount - b.amount) * dir;
        case 'issuedAt':
          return a.issuedAt.localeCompare(b.issuedAt) * dir;
        default:
          return 0;
      }
    });
  }, [entries, search, prefixFilter, statusFilter, sortKey, sortDir]);

  const issuedCount = entries.filter((e) => e.status === 'ISSUED').length;
  const voidCount = entries.filter((e) => e.status === 'VOID').length;
  const uniqueJobSequences = new Set(entries.map((e) => e.sequence)).size;

  const handleExportRegistry = async () => {
    setExporting(true);
    try {
      const filename = await exportMasterRegistryXlsx(entries, theme, company);
      onNotify(`Exported ${filename} (${entries.length} records)`);
    } catch {
      onNotify('Failed to export Master_Registry.xlsx. Please try again.');
    } finally {
      setExporting(false);
    }
  };

  const handleImportRegistry = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const result = await importWorkbookFromXlsx(file);
      if (result.kind === 'registry') {
        onSyncRegistryEntries(result.entries);
        onNotify(`Synced ${result.importedCount} entries from ${file.name}`);
      } else {
        onNotify('Selected file is a Job workbook, not Master_Registry.xlsx. Use Import & Export to open Job files.');
      }
    } catch {
      onNotify('Could not read Master_Registry.xlsx. Check file format.');
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const confirmVoid = () => {
    if (!voidingDoc) return;
    onVoidEntry(voidingDoc, voidReason.trim() || 'Voided by user');
    setVoidingDoc(null);
    setVoidReason('');
  };

  const renderSortIcon = (key: RegistrySortKey) => {
    if (sortKey !== key) return <ArrowUpDown className="w-3 h-3 opacity-55 shrink-0" />;
    return sortDir === 'asc' ? (
      <ArrowUp className="w-3 h-3 shrink-0" />
    ) : (
      <ArrowDown className="w-3 h-3 shrink-0" />
    );
  };

  return (
    <div className="space-y-6">
      {/* Summary Metrics Row */}
      <div className="grid grid-cols-1 sm:grid-cols-4 gap-4">
        <div className="rounded-lg border border-slate-200 bg-white p-4">
          <div className="text-xs text-slate-500">Next New Job Sequence</div>
          <div className="mt-1.5 text-2xl font-bold font-mono-tabular text-slate-900">
            #{String(nextSeq).padStart(3, '0')}
          </div>
          <div className="mt-1 text-xs text-slate-500 font-mono-tabular">
            PREFIX-GG-{String(nextSeq).padStart(3, '0')}-{new Date().getFullYear()}
          </div>
        </div>
        <div className="rounded-lg border border-slate-200 bg-white p-4">
          <div className="text-xs text-slate-500">Issued Documents</div>
          <div className="mt-1.5 text-2xl font-bold font-mono-tabular text-slate-900">
            {issuedCount}
          </div>
          <div className="mt-1 text-xs text-slate-500">
            {uniqueJobSequences === 0
              ? 'Starting fresh from #001'
              : `Across ${uniqueJobSequences} job sequence${uniqueJobSequences === 1 ? '' : 's'}`}
          </div>
        </div>
        <div className="rounded-lg border border-slate-200 bg-white p-4">
          <div className="text-xs text-slate-500">Per-Job Sequence Rule</div>
          <div className="mt-1.5 text-base font-bold font-mono-tabular text-emerald-700">
            1 Job = 1 Sequence
          </div>
          <div className="mt-1 text-xs text-slate-500">
            Prefix changes per doc (QT · PI · PO · CI · PL · DN)
          </div>
        </div>
        <div className="rounded-lg border border-slate-200 bg-white p-4">
          <div className="text-xs text-slate-500">Voided Records</div>
          <div className="mt-1.5 text-2xl font-bold font-mono-tabular text-red-700">
            {voidCount}
          </div>
          <div className="mt-1 text-xs text-slate-500">Audit trail preserved</div>
        </div>
      </div>

      {/* Filter, Sort & Export Controls */}
      <div className="rounded-lg border border-slate-200 bg-white p-4 flex flex-wrap items-center justify-between gap-3">
        <div className="flex flex-wrap items-center gap-2.5 flex-1 min-w-[260px]">
          <div className="relative min-w-[210px] flex-1 max-w-xs">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Search doc no., client, job, RFQ..."
              className="w-full min-h-[38px] rounded-lg border border-slate-200 pl-9 pr-3 py-1.5 text-xs text-slate-900 focus:border-slate-900 focus:outline-none"
            />
          </div>

          <div className="inline-flex items-center p-1 rounded-lg bg-slate-100">
            {(['ALL', 'QT', 'PI', 'PO', 'CI', 'PL', 'DN'] as const).map((p) => (
              <button
                key={p}
                type="button"
                onClick={() => setPrefixFilter(p)}
                className={`px-2.5 py-1.5 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap ${
                  prefixFilter === p
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                {p === 'ALL' ? 'All' : p}
              </button>
            ))}
          </div>

          <div className="inline-flex items-center p-1 rounded-lg bg-slate-100">
            {(['ALL', 'ISSUED', 'VOID'] as const).map((st) => (
              <button
                key={st}
                type="button"
                onClick={() => setStatusFilter(st)}
                className={`px-2.5 py-1.5 rounded-md text-xs font-medium transition-colors cursor-pointer whitespace-nowrap ${
                  statusFilter === st
                    ? 'bg-white text-slate-900 shadow-sm'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                {st === 'ALL' ? 'All Status' : st === 'ISSUED' ? 'Issued' : 'Void'}
              </button>
            ))}
          </div>

          {/* Explicit Sort Selector Dropdown (Syncs with Clickable Column Headers) */}
          <div className="flex items-center gap-1.5">
            <label htmlFor="registry-sort-select" className="text-xs text-slate-500 whitespace-nowrap">
              Sort:
            </label>
            <select
              id="registry-sort-select"
              value={sortKey}
              onChange={(e) => setSortKey(e.target.value as RegistrySortKey)}
              className="min-h-[38px] rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-800 focus:border-slate-900 focus:outline-none"
            >
              <option value="sequence">Job Sequence</option>
              <option value="docNumber">Document No.</option>
              <option value="prefix">Doc Type</option>
              <option value="customer">Client / Supplier</option>
              <option value="jobName">Job Reference</option>
              <option value="rfqRef">RFQ No. / Tender Ref</option>
              <option value="issuedAt">Issued Date</option>
              <option value="amount">Amount</option>
            </select>
            <button
              type="button"
              onClick={() => setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'))}
              title={sortDir === 'asc' ? 'Ascending' : 'Descending'}
              aria-label="Toggle sort direction"
              className="inline-flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[38px] whitespace-nowrap"
            >
              {sortDir === 'asc' ? (
                <>
                  <ArrowUp className="w-3.5 h-3.5" />
                  <span>Asc</span>
                </>
              ) : (
                <>
                  <ArrowDown className="w-3.5 h-3.5" />
                  <span>Desc</span>
                </>
              )}
            </button>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx"
            onChange={handleImportRegistry}
            className="hidden"
          />
          <button
            type="button"
            disabled={entries.length === 0}
            onClick={() =>
              onRequestConfirm({
                title: 'Reset Master Registry to 0?',
                message:
                  'This clears all registry entries and resets the next job sequence number back to #001. Issued numbers on existing jobs will also be unlocked so you can start from the beginning.',
                confirmLabel: 'Reset to 0',
                destructive: true,
                onConfirm: onResetRegistry,
              })
            }
            className="inline-flex items-center gap-1.5 rounded-lg border border-red-200 bg-white px-3 py-2 text-xs font-medium text-red-700 hover:bg-red-50 cursor-pointer min-h-[40px] whitespace-nowrap disabled:opacity-40 disabled:cursor-not-allowed"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>Reset to 0</span>
          </button>
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[40px] whitespace-nowrap"
          >
            <Upload className="w-3.5 h-3.5" />
            <span>Sync Master_Registry.xlsx</span>
          </button>
          <button
            type="button"
            disabled={exporting}
            onClick={handleExportRegistry}
            className="inline-flex items-center gap-1.5 rounded-lg px-3.5 py-2 text-xs font-semibold text-white hover:opacity-90 cursor-pointer min-h-[40px] whitespace-nowrap disabled:opacity-50"
            style={{ backgroundColor: theme.primary, color: theme.primaryText }}
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>{exporting ? 'Exporting...' : 'Export Master_Registry.xlsx'}</span>
          </button>
        </div>
      </div>

      <div className="flex justify-end">
        <button
          type="button"
          onClick={onExportAuditLog}
          className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[40px]"
        >
          <span>Export Change Log (.json)</span>
        </button>
      </div>

      {/* Registry Ledger Table with Clickable Sortable Headers */}
      <div className="rounded-lg border border-slate-200 bg-white overflow-hidden">
        <div className="overflow-x-auto sheet-scroll">
          <table className="w-full border-collapse text-xs">
            <thead>
              <tr style={{ backgroundColor: theme.primary, color: theme.primaryText }}>
                <th className="py-2.5 px-2 text-center font-semibold w-14">
                  <button
                    type="button"
                    onClick={() => handleToggleSort('sequence')}
                    className="inline-flex items-center gap-1 cursor-pointer hover:opacity-85 whitespace-nowrap"
                  >
                    <span>Seq</span>
                    {renderSortIcon('sequence')}
                  </button>
                </th>
                <th className="py-2.5 px-2.5 text-left font-semibold">
                  <button
                    type="button"
                    onClick={() => handleToggleSort('docNumber')}
                    className="inline-flex items-center gap-1 cursor-pointer hover:opacity-85 whitespace-nowrap"
                  >
                    <span>Document No.</span>
                    {renderSortIcon('docNumber')}
                  </button>
                </th>
                <th className="py-2.5 px-2.5 text-left font-semibold">
                  <button
                    type="button"
                    onClick={() => handleToggleSort('prefix')}
                    className="inline-flex items-center gap-1 cursor-pointer hover:opacity-85 whitespace-nowrap"
                  >
                    <span>Doc Type</span>
                    {renderSortIcon('prefix')}
                  </button>
                </th>
                <th className="py-2.5 px-2.5 text-left font-semibold">
                  <button
                    type="button"
                    onClick={() => handleToggleSort('customer')}
                    className="inline-flex items-center gap-1 cursor-pointer hover:opacity-85 whitespace-nowrap"
                  >
                    <span>Client / Supplier</span>
                    {renderSortIcon('customer')}
                  </button>
                </th>
                <th className="py-2.5 px-2.5 text-left font-semibold">
                  <button
                    type="button"
                    onClick={() => handleToggleSort('jobName')}
                    className="inline-flex items-center gap-1 cursor-pointer hover:opacity-85 whitespace-nowrap"
                  >
                    <span>Job Reference</span>
                    {renderSortIcon('jobName')}
                  </button>
                </th>
                <th className="py-2.5 px-2.5 text-left font-semibold">
                  <button
                    type="button"
                    onClick={() => handleToggleSort('rfqRef')}
                    className="inline-flex items-center gap-1 cursor-pointer hover:opacity-85 whitespace-nowrap"
                  >
                    <span>RFQ No. / Tender Ref</span>
                    {renderSortIcon('rfqRef')}
                  </button>
                </th>
                <th className="py-2.5 px-2.5 text-right font-semibold">
                  <button
                    type="button"
                    onClick={() => handleToggleSort('amount')}
                    className="inline-flex items-center gap-1 cursor-pointer hover:opacity-85 whitespace-nowrap"
                  >
                    <span>Amount</span>
                    {renderSortIcon('amount')}
                  </button>
                </th>
                <th className="py-2.5 px-2 text-left font-semibold">
                  <button
                    type="button"
                    onClick={() => handleToggleSort('issuedAt')}
                    className="inline-flex items-center gap-1 cursor-pointer hover:opacity-85 whitespace-nowrap"
                  >
                    <span>Issued Date</span>
                    {renderSortIcon('issuedAt')}
                  </button>
                </th>
                <th className="py-2.5 px-2 text-left font-semibold">Status</th>
                <th className="py-2.5 px-2.5 text-right font-semibold">Action</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {filtered.length === 0 ? (
                <tr>
                  <td colSpan={10} className="py-10 text-center text-slate-500">
                    {entries.length === 0
                      ? 'Registry is at 0. Issue your first document to claim #001.'
                      : 'No registry records match your current filter.'}
                  </td>
                </tr>
              ) : (
                filtered.map((entry, idx) => {
                  const isSupplier = entry.prefix === 'PO';
                  return (
                    <tr
                      key={entry.docNumber}
                      style={idx % 2 === 1 ? { backgroundColor: theme.zebraRow } : undefined}
                      className="hover:bg-slate-50 transition-colors"
                    >
                      <td className="py-2.5 px-2 text-center font-mono-tabular text-slate-500">
                        #{String(entry.sequence).padStart(3, '0')}
                      </td>
                      <td
                        className={`py-2.5 px-2.5 font-mono-tabular font-semibold whitespace-nowrap ${
                          entry.status === 'VOID' ? 'line-through text-red-700' : 'text-slate-900'
                        }`}
                      >
                        {entry.docNumber}
                      </td>
                      <td className="py-2.5 px-2.5 text-slate-700 whitespace-nowrap">
                        {DOC_META[entry.prefix].shortName}
                      </td>
                      <td className="py-2.5 px-2.5 max-w-[180px]">
                        <div className="font-medium text-slate-900 truncate" title={entry.customer}>
                          {entry.customer}
                        </div>
                        <div className="text-[11px] text-slate-500">
                          {isSupplier ? 'Supplier' : 'Client'}
                        </div>
                      </td>
                      <td className="py-2.5 px-2.5 text-slate-700 max-w-[150px] truncate" title={entry.jobName}>
                        {entry.jobName}
                      </td>
                      <td className="py-2.5 px-2.5 font-mono-tabular text-slate-700 whitespace-nowrap">
                        {entry.rfqRef?.trim() ? entry.rfqRef : '—'}
                      </td>
                      <td className="py-2.5 px-2.5 text-right font-mono-tabular text-slate-900 whitespace-nowrap">
                        {entry.amount > 0 ? formatMoney(entry.amount, entry.currency) : '—'}
                      </td>
                      <td className="py-2.5 px-2 font-mono-tabular text-slate-600 whitespace-nowrap">
                        {entry.issuedAt.slice(0, 10)}
                      </td>
                      <td className="py-2.5 px-2 whitespace-nowrap">
                        <span
                          className={`font-semibold ${
                            entry.status === 'VOID' ? 'text-red-700' : 'text-emerald-700'
                          }`}
                        >
                          {entry.status === 'VOID' ? 'VOID' : 'ISSUED'}
                        </span>
                        {entry.voidReason && (
                          <span
                            className="ml-1 text-[11px] text-slate-500 inline-block max-w-[90px] truncate align-bottom"
                            title={entry.voidReason}
                          >
                            · {entry.voidReason}
                          </span>
                        )}
                      </td>
                      <td className="py-2.5 px-2.5 text-right whitespace-nowrap">
                        {entry.status === 'ISSUED' ? (
                          <button
                            type="button"
                            onClick={() => setVoidingDoc(entry.docNumber)}
                            className="inline-flex items-center gap-1 text-xs font-medium text-red-700 hover:text-red-900 cursor-pointer whitespace-nowrap"
                          >
                            <Ban className="w-3.5 h-3.5" />
                            <span>Mark VOID</span>
                          </button>
                        ) : (
                          <span className="text-slate-400">—</span>
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

      {/* Void Confirmation Modal */}
      {voidingDoc !== null && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-6 shadow-xl">
            <h3 className="text-base font-semibold text-slate-900">
              Mark {voidingDoc} as VOID?
            </h3>
            <p className="mt-2 text-xs text-slate-600">
              Only this document record ({voidingDoc}) will be marked VOID in the registry. Other documents sharing this job&apos;s sequence remain intact.
            </p>
            <div className="mt-4">
              <label className="block text-xs font-medium text-slate-700 mb-1.5">
                Reason (Optional)
              </label>
              <input
                type="text"
                value={voidReason}
                onChange={(e) => setVoidReason(e.target.value)}
                placeholder="e.g. Client revised quantity / specification"
                className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm focus:border-slate-900 focus:outline-none"
              />
            </div>
            <div className="mt-5 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => {
                  setVoidingDoc(null);
                  setVoidReason('');
                }}
                className="rounded-lg border border-slate-200 px-4 py-2 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={confirmVoid}
                className="rounded-lg bg-red-600 px-4 py-2 text-xs font-semibold text-white hover:bg-red-700 cursor-pointer"
              >
                Confirm VOID
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
