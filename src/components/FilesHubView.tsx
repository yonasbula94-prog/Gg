import React, { useRef, useState } from 'react';
import {
  ClipboardPaste,
  Download,
  FileSpreadsheet,
  FolderOpen,
  HardDriveDownload,
  Monitor,
  Tablet,
  Upload,
} from 'lucide-react';
import { EXCEL_THEMES } from '../constants/themes';
import {
  CompanyProfile,
  ConfirmDialogRequest,
  ProcurementJob,
  RegistryEntry,
  ThemeCustomization,
  ThemePalette,
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
import { HelpTip } from './HelpTip';

interface FilesHubViewProps {
  activeJob: ProcurementJob;
  allJobs: ProcurementJob[];
  registry: RegistryEntry[];
  themeId: ThemeVariantId;
  theme: ThemePalette;
  themeCustomization: ThemeCustomization;
  company: CompanyProfile;
  onImportJob: (job: ProcurementJob) => void;
  onOpenSupplierQuote: (quote: ParsedSupplierQuote) => void;
  onSyncRegistry: (entries: RegistryEntry[]) => void;
  onRestoreBackup: (
    jobs: ProcurementJob[],
    registry: RegistryEntry[],
    company?: CompanyProfile,
    themeId?: ThemeVariantId,
    themeCustomization?: ThemeCustomization
  ) => void;
  onRequestConfirm: (req: ConfirmDialogRequest) => void;
  onNotify: (msg: string) => void;
}

export const FilesHubView: React.FC<FilesHubViewProps> = ({
  activeJob,
  allJobs,
  registry,
  themeId,
  theme,
  themeCustomization,
  company,
  onImportJob,
  onOpenSupplierQuote,
  onSyncRegistry,
  onRestoreBackup,
  onRequestConfirm,
  onNotify,
}) => {
  const [busyAction, setBusyAction] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const jsonInputRef = useRef<HTMLInputElement | null>(null);

  const handleDownloadTemplate = async (includeData: boolean) => {
    setBusyAction(includeData ? 'job' : 'template');
    try {
      const filename = await exportJobTemplateXlsx(activeJob, theme, includeData, company);
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
      const filename = await exportMasterRegistryXlsx(registry, theme, company);
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
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleExportJsonBackup = async () => {
    setBusyAction('backup');
    try {
      const payload = {
        exportedAt: new Date().toISOString(),
        version: 4,
        themeId,
        themeCustomization,
        company,
        jobs: allJobs,
        registry,
      };
      const filename = `GG_Procurement_Offline_Backup_${new Date().toISOString().slice(0, 10)}.json`;
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
      await shareOrDownloadBlob(blob, filename, 'application/json');
      onNotify('Exported offline JSON backup (jobs, registry, clients, branding, colors & font).');
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
        const incomingThemeId =
          parsed.themeId === 'v1' || parsed.themeId === 'v2' || parsed.themeId === 'v3'
            ? (parsed.themeId as ThemeVariantId)
            : undefined;
        const incomingThemeCustom = parsed.themeCustomization as ThemeCustomization | undefined;
        onRequestConfirm({
          title: 'Restore Workspace From Backup?',
          message: `This will replace your current ${allJobs.length} job${allJobs.length === 1 ? '' : 's'} and ${registry.length} registry record${registry.length === 1 ? '' : 's'} with ${incomingJobs.length} job${incomingJobs.length === 1 ? '' : 's'} and ${incomingRegistry.length} registry record${incomingRegistry.length === 1 ? '' : 's'} from "${file.name}".`,
          confirmLabel: 'Restore Backup',
          destructive: true,
          onConfirm: () => {
            onRestoreBackup(
              incomingJobs,
              incomingRegistry,
              incomingCompany,
              incomingThemeId,
              incomingThemeCustom
            );
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
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
      {/* Card 1: Import Quotes, Workbooks & Backup */}
      <div className="rounded-lg border border-slate-200 bg-white p-6 flex flex-col justify-between space-y-5">
        <div>
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-1.5">
              <h2 className="text-base font-semibold text-slate-900">Import Quotes &amp; Files</h2>
              <HelpTip text="Upload any supplier PDF or Excel (.xlsx) quote to capture descriptions, Part/CAS numbers, quantities, and unit prices via the Column Mapper Assistant—or open a saved Grand Group Job Workbook or Master_Registry.xlsx." />
            </div>
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
              className="inline-flex items-center gap-1.5 rounded-md border border-slate-200 px-2.5 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer"
            >
              <ClipboardPaste className="w-3.5 h-3.5" />
              <span>Paste Quote Text</span>
            </button>
          </div>

          <input
            ref={fileInputRef}
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
            onClick={() => fileInputRef.current?.click()}
            className={`mt-4 rounded-lg border-2 border-dashed p-8 text-center transition-colors cursor-pointer ${
              dragOver
                ? 'border-slate-900 bg-slate-50'
                : 'border-slate-300 hover:border-slate-400 bg-slate-50/50'
            }`}
          >
            <FolderOpen className="w-7 h-7 text-slate-500 mx-auto" />
            <div className="mt-2.5 text-sm font-semibold text-slate-900">
              {busyAction === 'import'
                ? 'Reading file...'
                : 'Upload Supplier Quote (.pdf / .xlsx) or Workbook'}
            </div>
            <div className="mt-1 text-xs text-slate-500">
              Click to browse or drop file here
            </div>
          </div>
        </div>

        {/* Backup & Restore (.json) */}
        <div className="pt-4 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-1.5">
            <span className="text-xs font-semibold text-slate-800">Workspace Backup (.json)</span>
            <HelpTip text="Saves ALL jobs, Master Registry numbers, End-User directory with addresses, company logo/branding, and custom colors/font into one file to back up or transfer between Windows PC and iPad." />
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

      {/* Card 2: Export Workbooks & Offline App */}
      <div className="rounded-lg border border-slate-200 bg-white p-6 space-y-4">
        <div className="flex items-center gap-1.5">
          <h2 className="text-base font-semibold text-slate-900">
            Export Workbooks &amp; Offline App
          </h2>
          <HelpTip text={`All exported Excel files embed your company logo, active style (${EXCEL_THEMES[themeId].shortLabel} · ${theme.fontName}), hidden gridlines, A4 print fit, and your live End-User list.`} />
        </div>

        <div className="divide-y divide-slate-100">
          <div className="py-3.5 flex items-center justify-between gap-4">
            <div className="flex items-center gap-1.5">
              <span className="text-sm font-medium text-slate-900">Blank Job Template</span>
              <HelpTip text="Downloads GG_Job_Template.xlsx with 9 formula-linked sheets, your company logo, active style/font, and current End-User dropdown list." />
            </div>
            <button
              type="button"
              disabled={busyAction !== null}
              onClick={() => handleDownloadTemplate(false)}
              className="inline-flex items-center gap-1.5 rounded-lg px-3.5 py-2 text-xs font-semibold hover:opacity-90 cursor-pointer min-h-[40px] whitespace-nowrap shrink-0 disabled:opacity-50"
              style={{ backgroundColor: theme.primary, color: theme.primaryText }}
            >
              <Download className="w-3.5 h-3.5" />
              <span>{busyAction === 'template' ? 'Building...' : 'Download .xlsx'}</span>
            </button>
          </div>

          <div className="py-3.5 flex items-center justify-between gap-4">
            <div className="flex items-center gap-1.5">
              <span className="text-sm font-medium text-slate-900">Current Job Workbook</span>
              <HelpTip text={`Exports "${activeJob.jobName}" as a full 9-tab Excel workbook with live formulas, custom columns, column order, and internal PRICE CALC.`} />
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
            <div className="flex items-center gap-1.5">
              <span className="text-sm font-medium text-slate-900">Master Registry</span>
              <HelpTip text={`Exports Master_Registry.xlsx containing all ${registry.length} issued/voided sequence records and PREFIXES lookup sheet.`} />
            </div>
            <button
              type="button"
              disabled={busyAction !== null}
              onClick={handleDownloadRegistry}
              className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3.5 py-2 text-xs font-medium text-slate-800 hover:bg-slate-50 cursor-pointer min-h-[40px] whitespace-nowrap shrink-0 disabled:opacity-50"
            >
              <FileSpreadsheet className="w-3.5 h-3.5 text-slate-700" />
              <span>{busyAction === 'registry' ? 'Building...' : 'Export Registry .xlsx'}</span>
            </button>
          </div>

          {typeof window !== 'undefined' && window.location.protocol !== 'file:' && (
            <>
              <div className="py-3.5 flex items-center justify-between gap-4">
                <div className="flex items-center gap-1.5">
                  <span className="text-sm font-medium text-slate-900">Windows Offline App</span>
                  <HelpTip text="Downloads a .zip containing the standalone offline .html app for Windows PC (works with zero internet) plus separate Tutorial/ guide and demo video attachments." />
                </div>
                <button
                  type="button"
                  disabled={busyAction !== null}
                  onClick={async () => {
                    setBusyAction('offline-win');
                    try {
                      const fn = await downloadOfflinePackage('windows');
                      onNotify(`Downloaded ${fn} — unzip to run offline on Windows (Tutorial/ folder included).`);
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

              <div className="py-3.5 flex items-center justify-between gap-4">
                <div className="flex items-center gap-1.5">
                  <span className="text-sm font-medium text-slate-900">iPad / iOS Offline App</span>
                  <HelpTip text="Downloads the iPad PWA app package (touch UI, offline cache, Home Screen icon) plus separate Tutorial/ attachments. Unzip, drop app folder on app.netlify.com/drop, open in iPad Safari, tap Add to Home Screen." />
                </div>
                <button
                  type="button"
                  disabled={busyAction !== null}
                  onClick={async () => {
                    setBusyAction('offline-ipad');
                    try {
                      const fn = await downloadOfflinePackage('ipad');
                      onNotify(`Downloaded ${fn} — contains the iPad app folder plus Tutorial/ attachments.`);
                    } catch {
                      onNotify('Could not download iPad package.');
                    } finally {
                      setBusyAction(null);
                    }
                  }}
                  className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3.5 py-2 text-xs font-semibold text-slate-900 hover:bg-slate-50 cursor-pointer min-h-[40px] whitespace-nowrap shrink-0 disabled:opacity-50"
                >
                  <Tablet className="w-3.5 h-3.5" />
                  <span>{busyAction === 'offline-ipad' ? 'Preparing...' : 'iPad (.zip)'}</span>
                </button>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
};
