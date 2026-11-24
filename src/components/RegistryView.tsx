import React, { useMemo, useRef, useState } from 'react';
import { Ban, FileSpreadsheet, RotateCcw, Search, Upload } from 'lucide-react';
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

interface RegistryViewProps {
  entries: RegistryEntry[];
  nextSeq: number;
  theme: ThemePalette;
  company: CompanyProfile;
  onVoidEntry: (sequence: number, reason: string) => void;
  onSyncRegistryEntries: (entries: RegistryEntry[]) => void;
  onResetRegistry: () => void;
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
  onRequestConfirm,
  onNotify,
}) => {
  const [search, setSearch] = useState('');
  const [prefixFilter, setPrefixFilter] = useState<'ALL' | DocPrefix>('ALL');
  const [statusFilter, setStatusFilter] = useState<'ALL' | 'ISSUED' | 'VOID'>('ALL');
  const [voidingSeq, setVoidingSeq] = useState<number | null>(null);
  const [voidReason, setVoidReason] = useState('');
  const [exporting, setExporting] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return [...entries]
      .sort((a, b) => b.sequence - a.sequence)
      .filter((e) => {
        if (prefixFilter !== 'ALL' && e.prefix !== prefixFilter) return false;
        if (statusFilter !== 'ALL' && e.status !== statusFilter) return false;
        if (!q) return true;
        return (
          e.docNumber.toLowerCase().includes(q) ||
          e.customer.toLowerCase().includes(q) ||
          e.jobName.toLowerCase().includes(q)
        );
      });
  }, [entries, search, prefixFilter, statusFilter]);

  const issuedCount = entries.filter((e) => e.status === 'ISSUED').length;
  const voidCount = entries.filter((e) => e.status === 'VOID').length;

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
    if (voidingSeq === null) return;
    onVoidEntry(voidingSeq, voidReason.trim() || 'Voided by user');
    setVoidingSeq(null);
    setVoidReason('');
  };

  return (
    <div className="space-y-6">
      {/* Summary Metrics Row (Clean flat grid, no nested cards) */}
      <div className="grid grid-cols-1 sm:grid-cols-4 gap-4">
        <div className="rounded-lg border border-slate-200 bg-white p-4">
          <div className="text-xs text-slate-500">Next Available Sequence</div>
          <div className="mt-1.5 text-2xl font-bold font-mono-tabular text-slate-900">
            #{String(nextSeq).padStart(3, '0')}
          </div>
          <div className="mt-1 text-xs text-slate-500 font-mono-tabular">
            PREFIX-GG-{String(nextSeq).padStart(3, '0')}-{new Date().getFullYear()}
          </div>
        </div>
        <div className="rounded-lg border border-slate-200 bg-white p-4">
          <div className="text-xs text-slate-500">Total Claimed</div>
          <div className="mt-1.5 text-2xl font-bold font-mono-tabular text-slate-900">{entries.length}</div>
          <div className="mt-1 text-xs text-slate-500">
            {entries.length === 0 ? 'Starting fresh from #001' : 'Shared across all document types'}
          </div>
        </div>
        <div className="rounded-lg border border-slate-200 bg-white p-4">
          <div className="text-xs text-slate-500">Active Issued</div>
          <div className="mt-1.5 text-2xl font-bold font-mono-tabular text-emerald-700">{issuedCount}</div>
          <div className="mt-1 text-xs text-slate-500">Shared across 6 prefixes</div>
        </div>
        <div className="rounded-lg border border-slate-200 bg-white p-4">
          <div className="text-xs text-slate-500">Voided (Never Reused)</div>
          <div className="mt-1.5 text-2xl font-bold font-mono-tabular text-red-700">{voidCount}</div>
          <div className="mt-1 text-xs text-slate-500">Gaps allowed · Zero duplicates</div>
        </div>
      </div>

      {/* Filter & Export Controls */}
      <div className="rounded-lg border border-slate-200 bg-white p-4 flex flex-wrap items-center justify-between gap-3">
        <div className="flex flex-wrap items-center gap-2.5 flex-1 min-w-[260px]">
          <div className="relative flex-1 max-w-xs">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Search doc no., customer, job..."
              className="w-full min-h-[40px] rounded-lg border border-slate-200 pl-9 pr-3 py-2 text-xs text-slate-900 focus:border-slate-900 focus:outline-none"
            />
          </div>

          <div className="inline-flex items-center p-1 rounded-lg bg-slate-100">
            {(['ALL', 'PI', 'QT', 'PO', 'PL', 'DN', 'CI'] as const).map((p) => (
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
                  'This clears all registry entries and resets the next document number back to #001. Issued numbers on existing jobs will also be unlocked so you can start from the beginning.',
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
            style={{ backgroundColor: theme.primary }}
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>{exporting ? 'Exporting...' : 'Export Master_Registry.xlsx'}</span>
          </button>
        </div>
      </div>

      {/* Registry Ledger Table */}
      <div className="rounded-lg border border-slate-200 bg-white overflow-hidden">
        <div className="overflow-x-auto sheet-scroll">
          <table className="w-full border-collapse text-xs">
            <thead>
              <tr style={{ backgroundColor: theme.primary, color: theme.primaryText }}>
                <th className="py-3 px-3 text-center font-semibold w-16">Seq</th>
                <th className="py-3 px-3 text-left font-semibold">Document No.</th>
                <th className="py-3 px-3 text-left font-semibold">Type</th>
                <th className="py-3 px-3 text-left font-semibold">Customer / Supplier</th>
                <th className="py-3 px-3 text-left font-semibold">Job Reference</th>
                <th className="py-3 px-3 text-right font-semibold">Amount</th>
                <th className="py-3 px-3 text-left font-semibold">Issued Date</th>
                <th className="py-3 px-3 text-left font-semibold">Status</th>
                <th className="py-3 px-3 text-right font-semibold">Action</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200">
              {filtered.length === 0 ? (
                <tr>
                  <td colSpan={9} className="py-10 text-center text-slate-500">
                    {entries.length === 0
                      ? 'Registry is at 0. Issue your first document to claim #001.'
                      : 'No registry records match your current filter.'}
                  </td>
                </tr>
              ) : (
                filtered.map((entry, idx) => (
                  <tr
                    key={entry.sequence}
                    style={idx % 2 === 1 ? { backgroundColor: theme.zebraRow } : undefined}
                    className="hover:bg-slate-50 transition-colors"
                  >
                    <td className="py-3 px-3 text-center font-mono-tabular text-slate-500">
                      {entry.sequence}
                    </td>
                    <td
                      className={`py-3 px-3 font-mono-tabular font-semibold ${
                        entry.status === 'VOID' ? 'line-through text-red-700' : 'text-slate-900'
                      }`}
                    >
                      {entry.docNumber}
                    </td>
                    <td className="py-3 px-3 text-slate-700">{DOC_META[entry.prefix].shortName}</td>
                    <td className="py-3 px-3 font-medium text-slate-900">{entry.customer}</td>
                    <td className="py-3 px-3 text-slate-600">{entry.jobName}</td>
                    <td className="py-3 px-3 text-right font-mono-tabular text-slate-900">
                      {entry.amount > 0 ? formatMoney(entry.amount, entry.currency) : '—'}
                    </td>
                    <td className="py-3 px-3 font-mono-tabular text-slate-600">
                      {entry.issuedAt.slice(0, 10)}
                    </td>
                    <td className="py-3 px-3">
                      <span
                        className={`font-semibold ${
                          entry.status === 'VOID' ? 'text-red-700' : 'text-emerald-700'
                        }`}
                      >
                        {entry.status === 'VOID' ? 'VOID' : 'ISSUED'}
                      </span>
                      {entry.voidReason && (
                        <span className="ml-1.5 text-slate-500">· {entry.voidReason}</span>
                      )}
                    </td>
                    <td className="py-3 px-3 text-right">
                      {entry.status === 'ISSUED' ? (
                        <button
                          type="button"
                          onClick={() => setVoidingSeq(entry.sequence)}
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
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Void Confirmation Modal */}
      {voidingSeq !== null && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-6 shadow-xl">
            <h3 className="text-base font-semibold text-slate-900">
              Mark Sequence #{voidingSeq} as VOID?
            </h3>
            <p className="mt-2 text-xs text-slate-600">
              Per Grand Group FZC registry rules, voided numbers are permanently retired and never reused. Sequence gaps are acceptable; duplicate numbers are never allowed.
            </p>
            <div className="mt-4">
              <label className="block text-xs font-medium text-slate-700 mb-1.5">
                Reason (Optional)
              </label>
              <input
                type="text"
                value={voidReason}
                onChange={(e) => setVoidReason(e.target.value)}
                placeholder="e.g. File not saved / customer revised qty"
                className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm focus:border-slate-900 focus:outline-none"
              />
            </div>
            <div className="mt-5 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => {
                  setVoidingSeq(null);
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
