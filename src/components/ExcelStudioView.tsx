import React, { useRef, useState } from 'react';
import {
  Check,
  ClipboardPaste,
  Download,
  FileSpreadsheet,
  FolderOpen,
  HardDriveDownload,
  Image as ImageIcon,
  Monitor,
  Plus,
  RotateCcw,
  Tablet,
  Trash2,
  Upload,
} from 'lucide-react';
import { EXCEL_THEMES } from '../constants/themes';
import {
  CompanyProfile,
  ConfirmDialogRequest,
  ProcurementJob,
  RegistryEntry,
  ThemeVariantId,
} from '../types/procurement';
import {
  exportJobTemplateXlsx,
  exportMasterRegistryXlsx,
  importWorkbookFromXlsx,
} from '../utils/excelEngine';
import { downloadOfflinePackage } from '../utils/offlineDownload';
import { shareOrDownloadBlob } from '../utils/platform';
import { ParsedSupplierQuote, parseSupplierPdfFile } from '../utils/supplierQuoteParser';

interface ExcelStudioViewProps {
  activeJob: ProcurementJob;
  allJobs: ProcurementJob[];
  registry: RegistryEntry[];
  themeId: ThemeVariantId;
  company: CompanyProfile;
  onSelectTheme: (id: ThemeVariantId) => void;
  onUpdateCompany: (patch: Partial<CompanyProfile>) => void;
  onResetCompany: () => void;
  onImportJob: (job: ProcurementJob) => void;
  onOpenSupplierQuote: (quote: ParsedSupplierQuote) => void;
  onSyncRegistry: (entries: RegistryEntry[]) => void;
  onRestoreBackup: (
    jobs: ProcurementJob[],
    registry: RegistryEntry[],
    company?: CompanyProfile
  ) => void;
  onRequestConfirm: (req: ConfirmDialogRequest) => void;
  onNotify: (msg: string) => void;
}

export const ExcelStudioView: React.FC<ExcelStudioViewProps> = ({
  activeJob,
  allJobs,
  registry,
  themeId,
  company,
  onSelectTheme,
  onUpdateCompany,
  onResetCompany,
  onImportJob,
  onOpenSupplierQuote,
  onSyncRegistry,
  onRestoreBackup,
  onRequestConfirm,
  onNotify,
}) => {
  const [busyAction, setBusyAction] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [newEndUser, setNewEndUser] = useState('');
  const xlsxInputRef = useRef<HTMLInputElement | null>(null);
  const jsonInputRef = useRef<HTMLInputElement | null>(null);
  const logoInputRef = useRef<HTMLInputElement | null>(null);

  const handleDownloadTemplate = async (includeData: boolean) => {
    setBusyAction(includeData ? 'job' : 'template');
    try {
      const filename = await exportJobTemplateXlsx(activeJob, themeId, includeData, company);
      onNotify(`Generated live-formula workbook: ${filename}`);
    } catch {
      onNotify('Failed to build Excel workbook.');
    } finally {
      setBusyAction(null);
    }
  };

  const handleDownloadRegistry = async () => {
    setBusyAction('registry');
    try {
      const filename = await exportMasterRegistryXlsx(registry, themeId, company);
      onNotify(`Generated ${filename}`);
    } catch {
      onNotify('Failed to export Master Registry.');
    } finally {
      setBusyAction(null);
    }
  };

  const processUploadedFile = async (file: File) => {
    setBusyAction('import');
    try {
      const isPdf =
        file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
      if (isPdf) {
        const quote = await parseSupplierPdfFile(file);
        onOpenSupplierQuote(quote);
        return;
      }

      const res = await importWorkbookFromXlsx(file, company.endUsers);
      if (res.kind === 'registry') {
        onSyncRegistry(res.entries);
        onNotify(`Synced ${res.importedCount} Master Registry records from ${file.name}`);
      } else if (res.kind === 'supplier_quote') {
        onOpenSupplierQuote(res.quote);
      } else {
        onImportJob(res.job);
        onNotify(`Opened "${res.job.jobName}" inside app (${res.itemCount} line items loaded)`);
      }
    } catch {
      onNotify('Could not read this file. Upload a supplier .pdf/.xlsx quote or Grand Group .xlsx workbook.');
    } finally {
      setBusyAction(null);
      if (xlsxInputRef.current) xlsxInputRef.current.value = '';
    }
  };

  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith('image/')) {
      onNotify('Please choose a PNG or JPG logo image.');
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === 'string') {
        onUpdateCompany({ logoDataUrl: reader.result });
        onNotify('Updated company logo across app, A4 print, and Excel exports.');
      }
    };
    reader.readAsDataURL(file);
    if (logoInputRef.current) logoInputRef.current.value = '';
  };

  const handleAddEndUser = () => {
    const trimmed = newEndUser.trim();
    if (!trimmed) return;
    if (company.endUsers.includes(trimmed)) {
      onNotify(`"${trimmed}" is already in the End-User list.`);
      return;
    }
    // Keep OTHER at the end of the list
    const withoutOther = company.endUsers.filter((u) => u !== 'OTHER');
    onUpdateCompany({ endUsers: [...withoutOther, trimmed, 'OTHER'] });
    setNewEndUser('');
    onNotify(`Added "${trimmed}" to End-User dropdown.`);
  };

  const handleRemoveEndUser = (name: string) => {
    if (name === 'OTHER') return;
    onUpdateCompany({ endUsers: company.endUsers.filter((u) => u !== name) });
  };

  const handleExportJsonBackup = async () => {
    setBusyAction('backup');
    try {
      const payload = {
        exportedAt: new Date().toISOString(),
        version: 3,
        company,
        jobs: allJobs,
        registry,
      };
      const filename = `GG_Procurement_Offline_Backup_${new Date().toISOString().slice(0, 10)}.json`;
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
      await shareOrDownloadBlob(blob, filename, 'application/json');
      onNotify('Exported offline JSON backup (includes jobs, registry, logo & company settings).');
    } catch {
      onNotify('Failed to export JSON backup.');
    } finally {
      setBusyAction(null);
    }
  };

  const handleImportJsonBackup = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      if (Array.isArray(parsed.jobs) && Array.isArray(parsed.registry)) {
        const incomingJobs = parsed.jobs as ProcurementJob[];
        const incomingRegistry = parsed.registry as RegistryEntry[];
        const incomingCompany = parsed.company as CompanyProfile | undefined;
        onRequestConfirm({
          title: 'Restore Workspace From Backup?',
          message: `This will replace your current ${allJobs.length} job${allJobs.length === 1 ? '' : 's'} and ${registry.length} registry record${registry.length === 1 ? '' : 's'} with ${incomingJobs.length} job${incomingJobs.length === 1 ? '' : 's'} and ${incomingRegistry.length} registry record${incomingRegistry.length === 1 ? '' : 's'} from "${file.name}".`,
          confirmLabel: 'Restore Backup',
          destructive: true,
          onConfirm: () => {
            onRestoreBackup(incomingJobs, incomingRegistry, incomingCompany);
            onNotify(
              `Restored ${incomingJobs.length} jobs and ${incomingRegistry.length} registry entries.`
            );
          },
        });
      } else {
        onNotify('Invalid backup JSON structure.');
      }
    } catch {
      onNotify('Failed to read backup file.');
    } finally {
      if (jsonInputRef.current) jsonInputRef.current.value = '';
    }
  };

  return (
    <div className="space-y-6">
      {/* 3-Variant Excel Color Theme Switcher & Visual Comparison */}
      <div className="rounded-lg border border-slate-200 bg-white p-6">
        <div className="flex flex-wrap items-center justify-between gap-2 pb-4 border-b border-slate-100">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              Excel Document Style Variants (v1 · v2 · v3)
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Select a palette below—it updates every A4 document preview and every exported `.xlsx` workbook immediately.
            </p>
          </div>
          <span className="text-xs font-mono-tabular text-slate-600">
            Active: {EXCEL_THEMES[themeId].name}
          </span>
        </div>

        <div className="mt-5 grid grid-cols-1 lg:grid-cols-3 gap-5">
          {(['v1', 'v2', 'v3'] as ThemeVariantId[]).map((id) => {
            const t = EXCEL_THEMES[id];
            const isSelected = themeId === id;

            return (
              <button
                key={id}
                type="button"
                onClick={() => onSelectTheme(id)}
                className={`text-left rounded-lg border p-4 transition-colors cursor-pointer ${
                  isSelected
                    ? 'border-slate-900 ring-1 ring-slate-900 bg-slate-50/40'
                    : 'border-slate-200 hover:border-slate-300 bg-white'
                }`}
              >
                <div className="flex items-center justify-between gap-2">
                  <span className="text-sm font-semibold text-slate-900">{t.name}</span>
                  {isSelected && (
                    <span className="inline-flex items-center gap-1 text-xs font-semibold text-slate-900">
                      <Check className="w-3.5 h-3.5" /> Active
                    </span>
                  )}
                </div>
                <p className="text-xs text-slate-500 mt-1">{t.subtitle}</p>

                {/* Live Mini Excel Sheet Specimen */}
                <div className="mt-4 rounded border overflow-hidden text-[11px] font-mono-tabular" style={{ borderColor: t.border }}>
                  <div
                    className="px-2.5 py-1.5 flex items-center justify-between font-sans font-semibold"
                    style={{ backgroundColor: t.primary, color: t.primaryText }}
                  >
                    <span className="flex items-center gap-1.5 min-w-0">
                      {company.logoDataUrl && (
                        <span className="w-4 h-4 rounded-sm bg-white p-0.5 flex items-center justify-center shrink-0">
                          <img src={company.logoDataUrl} alt="" className="max-w-full max-h-full object-contain" />
                        </span>
                      )}
                      <span className="truncate">{company.name}</span>
                    </span>
                    <span className="shrink-0">PROFORMA</span>
                  </div>
                  <div
                    className="px-2.5 py-1 text-[10px] border-b-2 flex justify-between"
                    style={{ backgroundColor: t.secondaryBg, borderColor: t.accent, color: t.mutedText }}
                  >
                    <span>AIAA · CIF Massawa</span>
                    <span>PI-GG-043-2026</span>
                  </div>
                  <div className="bg-white divide-y" style={{ borderColor: t.border }}>
                    <div className="px-2.5 py-1 flex justify-between text-slate-800">
                      <span className="font-sans truncate pr-2">LED Runway Edge Light</span>
                      <span>7,866.00</span>
                    </div>
                    <div
                      className="px-2.5 py-1 flex justify-between text-slate-800"
                      style={{ backgroundColor: t.zebraRow }}
                    >
                      <span className="font-sans truncate pr-2">Isolation Transformer 100W</span>
                      <span>1,987.20</span>
                    </div>
                    <div
                      className="px-2.5 py-1.5 flex justify-between font-bold"
                      style={{ backgroundColor: t.totalBg, color: t.primary }}
                    >
                      <span className="font-sans">GRAND TOTAL</span>
                      <span>10,345.86</span>
                    </div>
                  </div>
                </div>
              </button>
            );
          })}
        </div>
      </div>

      {/* Open / Import .xlsx Directly Inside App + Workbook Downloads */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Dropzone: Upload Supplier Quote (.pdf / .xlsx) or Grand Group Workbook */}
        <div className="rounded-lg border border-slate-200 bg-white p-6 flex flex-col justify-between">
          <div>
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h3 className="text-sm font-semibold text-slate-900">
                Upload Supplier Quote (.pdf / .xlsx) or Grand Group Workbook
              </h3>
              <button
                type="button"
                onClick={() =>
                  onOpenSupplierQuote({
                    sourceType: 'text',
                    filename: 'Pasted Supplier Quote',
                    supplierName: '',
                    supplierRef: '',
                    currency: activeJob.controls.currency || 'USD',
                    incoterm: '',
                    paymentTerms: '',
                    deliveryLeadTime: '',
                    detectedCodeLabel: 'Part No.',
                    items: [],
                    rawExtractedText: '',
                  })
                }
                className="inline-flex items-center gap-1.5 rounded-md border border-slate-200 px-2.5 py-1 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer"
              >
                <ClipboardPaste className="w-3.5 h-3.5" />
                <span>Paste Quote Text</span>
              </button>
            </div>
            <p className="text-xs text-slate-500 mt-1">
              Works offline with <strong>any supplier Excel (.xlsx) or PDF quote</strong> (auto-captures descriptions, Part/CAS numbers, qty, prices &amp; terms into Price Calc), plus Grand Group templates &amp; Master_Registry.xlsx.
            </p>

            <input
              ref={xlsxInputRef}
              type="file"
              accept=".xlsx,.pdf,application/pdf"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) processUploadedFile(f);
              }}
              className="hidden"
            />

            <div
              onDragOver={(e) => {
                e.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={() => setDragOver(false)}
              onDrop={(e) => {
                e.preventDefault();
                setDragOver(false);
                const f = e.dataTransfer.files?.[0];
                if (f) processUploadedFile(f);
              }}
              onClick={() => xlsxInputRef.current?.click()}
              className={`mt-4 rounded-lg border-2 border-dashed p-7 text-center transition-colors cursor-pointer ${
                dragOver
                  ? 'border-slate-900 bg-slate-50'
                  : 'border-slate-300 hover:border-slate-400 bg-slate-50/50'
              }`}
            >
              <FolderOpen className="w-7 h-7 text-slate-500 mx-auto" />
              <div className="mt-2.5 text-sm font-medium text-slate-900">
                {busyAction === 'import'
                  ? 'Reading file & extracting prices...'
                  : 'Click or drop supplier .pdf / .xlsx or Grand Group .xlsx'}
              </div>
              <div className="mt-1 text-xs text-slate-500">
                Supplier quotes open in a review window before loading into Price Calc
              </div>
            </div>
          </div>

          <div className="mt-5 pt-4 border-t border-slate-100 space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div>
                <div className="text-xs font-semibold text-slate-800">
                  Full App Backup &amp; Device Transfer (.json)
                </div>
                <div className="text-[11px] text-slate-500">
                  Saves ALL your jobs, Registry numbers, logo &amp; settings into one file to copy between Laptop &amp; iPad
                </div>
              </div>
              <div className="flex items-center gap-2">
                <input
                  ref={jsonInputRef}
                  type="file"
                  accept=".json"
                  onChange={handleImportJsonBackup}
                  className="hidden"
                />
                <button
                  type="button"
                  onClick={handleExportJsonBackup}
                  className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[36px] whitespace-nowrap"
                >
                  <HardDriveDownload className="w-3.5 h-3.5" />
                  <span>Backup JSON</span>
                </button>
                <button
                  type="button"
                  onClick={() => jsonInputRef.current?.click()}
                  className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[36px] whitespace-nowrap"
                >
                  <Upload className="w-3.5 h-3.5" />
                  <span>Restore JSON</span>
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* Formula-Driven Excel Deliverables */}
        <div className="rounded-lg border border-slate-200 bg-white p-6 space-y-4">
          <div>
            <h3 className="text-sm font-semibold text-slate-900">
              Download Styled Workbooks &amp; Standalone Offline App
            </h3>
            <p className="text-xs text-slate-500 mt-1">
              Every exported workbook embeds your Grand Group logo, active style ({EXCEL_THEMES[themeId].shortLabel}), hidden gridlines, A4 print fit, and your live End-User list.
            </p>
          </div>

          <div className="divide-y divide-slate-100">
            <div className="py-3.5 flex items-center justify-between gap-4">
              <div>
                <div className="text-sm font-medium text-slate-900">GG_Job_Template.xlsx (Blank Master)</div>
                <div className="text-xs text-slate-500">
                  9 linked tabs with live Excel formulas, embedded logo, and updated CONTROLS list
                </div>
              </div>
              <button
                type="button"
                disabled={busyAction !== null}
                onClick={() => handleDownloadTemplate(false)}
                className="inline-flex items-center gap-1.5 rounded-lg px-3.5 py-2 text-xs font-semibold text-white hover:opacity-90 cursor-pointer min-h-[40px] whitespace-nowrap shrink-0 disabled:opacity-50"
                style={{ backgroundColor: EXCEL_THEMES[themeId].primary }}
              >
                <Download className="w-3.5 h-3.5" />
                <span>{busyAction === 'template' ? 'Building...' : 'Download Template'}</span>
              </button>
            </div>

            <div className="py-3.5 flex items-center justify-between gap-4">
              <div>
                <div className="text-sm font-medium text-slate-900">Current Job Full Workbook (.xlsx)</div>
                <div className="text-xs text-slate-500">
                  Exports &ldquo;{activeJob.jobName}&rdquo; with live formulas + internal PRICE CALC
                </div>
              </div>
              <button
                type="button"
                disabled={busyAction !== null}
                onClick={() => handleDownloadTemplate(true)}
                className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-xs font-medium text-slate-800 hover:bg-slate-50 cursor-pointer min-h-[40px] whitespace-nowrap shrink-0 disabled:opacity-50"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-700" />
                <span>{busyAction === 'job' ? 'Building...' : 'Export Job .xlsx'}</span>
              </button>
            </div>

            <div className="py-3.5 flex items-center justify-between gap-4">
              <div>
                <div className="text-sm font-medium text-slate-900">Master_Registry.xlsx</div>
                <div className="text-xs text-slate-500">
                  Contains REGISTRY ({registry.length} records, seeded from PI-GG-042-2026) and PREFIXES sheets
                </div>
              </div>
              <button
                type="button"
                disabled={busyAction !== null}
                onClick={handleDownloadRegistry}
                className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-xs font-medium text-slate-800 hover:bg-slate-50 cursor-pointer min-h-[40px] whitespace-nowrap shrink-0 disabled:opacity-50"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 text-slate-700" />
                <span>{busyAction === 'registry' ? 'Building...' : 'Download Registry'}</span>
              </button>
            </div>

            {typeof window !== 'undefined' && window.location.protocol !== 'file:' && (
              <>
                <div className="py-3.5 flex items-center justify-between gap-4">
                  <div>
                    <div className="text-sm font-medium text-slate-900">
                      Windows PC Offline Package (.zip)
                    </div>
                    <div className="text-xs text-slate-500">
                      Unzip to launch the standalone offline .html app directly on Windows (zero internet needed) — includes separate Tutorial/ guide &amp; demo video attachments.
                    </div>
                  </div>
                  <button
                    type="button"
                    disabled={busyAction !== null}
                    onClick={async () => {
                      setBusyAction('offline-win');
                      try {
                        const fn = await downloadOfflinePackage('windows');
                        onNotify(`Downloaded ${fn} — unzip to run the standalone Windows app (Tutorial/ folder attached inside).`);
                      } catch {
                        onNotify('Could not download Windows offline package.');
                      } finally {
                        setBusyAction(null);
                      }
                    }}
                    className="inline-flex items-center gap-1.5 rounded-lg bg-slate-900 px-3.5 py-2 text-xs font-semibold text-white hover:bg-slate-800 cursor-pointer min-h-[40px] whitespace-nowrap shrink-0 disabled:opacity-50"
                  >
                    <Monitor className="w-3.5 h-3.5" />
                    <span>{busyAction === 'offline-win' ? 'Preparing...' : 'Windows (.zip)'}</span>
                  </button>
                </div>

                <div className="py-3.5 space-y-2">
                  <div className="flex items-center justify-between gap-4">
                    <div>
                      <div className="text-sm font-medium text-slate-900">
                        iPad / iOS Offline App Package (.zip)
                      </div>
                      <div className="text-xs text-slate-500">
                        Includes touch app, offline cache (sw.js), Grand Group Home Screen icon, and separate Tutorial/ attachments. Unzip → drag the app folder onto app.netlify.com/drop → open in iPad Safari → Add to Home Screen.
                      </div>
                    </div>
                    <button
                      type="button"
                      disabled={busyAction !== null}
                      onClick={async () => {
                        setBusyAction('offline-ipad');
                        try {
                          const fn = await downloadOfflinePackage('ipad');
                          onNotify(`Downloaded ${fn} — contains the iPad PWA app folder plus separate Tutorial/ attachments.`);
                        } catch {
                          onNotify('Could not download iPad package.');
                        } finally {
                          setBusyAction(null);
                        }
                      }}
                      className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3.5 py-2 text-xs font-semibold text-slate-900 hover:bg-slate-50 cursor-pointer min-h-[40px] whitespace-nowrap shrink-0 disabled:opacity-50"
                    >
                      <Tablet className="w-3.5 h-3.5" />
                      <span>{busyAction === 'offline-ipad' ? 'Preparing...' : 'iPad App (.zip)'}</span>
                    </button>
                  </div>
                </div>
              </>
            )}
          </div>
        </div>
      </div>

      {/* Live Editable Company Profile, Logo & End-User Dropdown (Zero Code Needed!) */}
      <div className="rounded-lg border border-slate-200 bg-white p-6 space-y-5">
        <div className="flex flex-wrap items-center justify-between gap-3 pb-4 border-b border-slate-100">
          <div>
            <h3 className="text-base font-semibold text-slate-900">
              Company Branding, Logo &amp; End-User Dropdown Settings
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Edit company details, bank info, logo, or client dropdown list anytime directly inside your offline app—updates all A4 prints and `.xlsx` exports automatically.
            </p>
          </div>
          <button
            type="button"
            onClick={() =>
              onRequestConfirm({
                title: 'Reset Company Branding & Defaults?',
                message:
                  'This will restore the default Grand Group FZC logo, address, bank details, and standard End-User dropdown list.',
                confirmLabel: 'Reset Defaults',
                destructive: true,
                onConfirm: onResetCompany,
              })
            }
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 hover:text-slate-900 cursor-pointer min-h-[36px] whitespace-nowrap"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            <span>Reset Defaults</span>
          </button>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          {/* Logo Preview & Upload */}
          <div className="lg:col-span-3 flex flex-col items-center justify-center rounded-lg border border-slate-200 bg-slate-50/50 p-5 text-center">
            <div className="w-24 h-24 rounded-lg border border-slate-200 bg-white p-2 flex items-center justify-center">
              {company.logoDataUrl ? (
                <img
                  src={company.logoDataUrl}
                  alt={company.name}
                  className="max-w-full max-h-full object-contain"
                />
              ) : (
                <ImageIcon className="w-8 h-8 text-slate-300" />
              )}
            </div>
            <input
              ref={logoInputRef}
              type="file"
              accept="image/png,image/jpeg"
              onChange={handleLogoUpload}
              className="hidden"
            />
            <button
              type="button"
              onClick={() => logoInputRef.current?.click()}
              className="mt-3 inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[36px]"
            >
              <Upload className="w-3.5 h-3.5" />
              <span>Replace Logo</span>
            </button>
            <span className="mt-1.5 text-[11px] text-slate-400">PNG or JPG · Embedded in .xlsx</span>
          </div>

          {/* Company Text Fields */}
          <div className="lg:col-span-9 grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">Company Name</label>
              <input
                type="text"
                value={company.name}
                onChange={(e) => onUpdateCompany({ name: e.target.value })}
                className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">Tagline</label>
              <input
                type="text"
                value={company.tagline}
                onChange={(e) => onUpdateCompany({ tagline: e.target.value })}
                className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">Address Line 1</label>
              <input
                type="text"
                value={company.addressLine1}
                onChange={(e) => onUpdateCompany({ addressLine1: e.target.value })}
                className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">Address Line 2</label>
              <input
                type="text"
                value={company.addressLine2}
                onChange={(e) => onUpdateCompany({ addressLine2: e.target.value })}
                className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">Email</label>
              <input
                type="text"
                value={company.email}
                onChange={(e) => onUpdateCompany({ email: e.target.value })}
                className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">TRN / Tax Registration</label>
              <input
                type="text"
                value={company.trn}
                onChange={(e) => onUpdateCompany({ trn: e.target.value })}
                className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono-tabular text-slate-900 focus:border-slate-900 focus:outline-none"
              />
            </div>
            <div className="sm:col-span-2">
              <label className="block text-xs font-medium text-slate-600 mb-1">Bank &amp; Wire Details (Shown on PI &amp; CI)</label>
              <input
                type="text"
                value={company.bankDetails}
                onChange={(e) => onUpdateCompany({ bankDetails: e.target.value })}
                className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono-tabular text-slate-900 focus:border-slate-900 focus:outline-none"
              />
            </div>
          </div>
        </div>

        {/* Editable End-User Dropdown List */}
        <div className="pt-4 border-t border-slate-100 space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <span className="text-xs font-semibold text-slate-800">
              End-User Dropdown List ({company.endUsers.length} entries · CONTROLS!A2:A{company.endUsers.length + 1})
            </span>
            <div className="flex items-center gap-2">
              <input
                type="text"
                value={newEndUser}
                onChange={(e) => setNewEndUser(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleAddEndUser();
                }}
                placeholder="Add new End User..."
                className="min-h-[36px] rounded-lg border border-slate-200 px-3 py-1.5 text-xs text-slate-900 focus:border-slate-900 focus:outline-none w-52"
              />
              <button
                type="button"
                onClick={handleAddEndUser}
                className="inline-flex items-center gap-1 rounded-lg bg-slate-900 px-3 py-1.5 text-xs font-medium text-white hover:bg-slate-800 cursor-pointer min-h-[36px] whitespace-nowrap"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>Add</span>
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2">
            {company.endUsers.map((user) => (
              <div
                key={user}
                className="flex items-center justify-between gap-2 rounded-md border border-slate-200 bg-slate-50/60 px-3 py-2 text-xs text-slate-800"
              >
                <span className="truncate">{user}</span>
                {user !== 'OTHER' && (
                  <button
                    type="button"
                    onClick={() => handleRemoveEndUser(user)}
                    title={`Remove ${user}`}
                    className="text-slate-400 hover:text-red-600 cursor-pointer shrink-0"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                )}
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};
