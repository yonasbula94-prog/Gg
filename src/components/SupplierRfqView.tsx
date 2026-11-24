import React, { useMemo, useState } from 'react';
import ExcelJS from 'exceljs';
import { FileSpreadsheet, Printer, Upload } from 'lucide-react';
import { CURRENCY_OPTIONS, INCOTERM_OPTIONS } from '../constants/themes';
import { exportSupplierRfqXlsx } from '../utils/excelEngine';
import { CompanyProfile, ProcurementJob, SupplierRfqData, ThemePalette } from '../types/procurement';
import { ParsedQuoteItem, parseGenericSupplierXlsx, parseSupplierPdfFile } from '../utils/supplierQuoteParser';

interface SupplierRfqViewProps {
  job: ProcurementJob;
  company: CompanyProfile;
  theme: ThemePalette;
  onUpdateRfq: (patch: Partial<SupplierRfqData>) => void;
  /** Fills empty item rows, in order. Existing rows are never overwritten. */
  onFillEmptyRows: (items: { requestedDescription: string; uom: string; qty: number | null }[]) => void;
  onNotify: (msg: string) => void;
}

const isEmptyRow = (it: { requestedDescription: string; offeredDescription: string }) =>
  !it.requestedDescription && !it.offeredDescription;

/**
 * Supplier RFQ: a request for quotation sent to suppliers, built from the end-user RFQ.
 * The end-user's name is hidden by default. Prints on the company letterhead (Save as PDF from the print dialog).
 */
export const SupplierRfqView: React.FC<SupplierRfqViewProps> = ({
  job,
  company,
  theme,
  onUpdateRfq,
  onFillEmptyRows,
  onNotify,
}) => {
  const rfq: SupplierRfqData = job.supplierRfq ?? {
    rfqNumber: '',
    supplierName: '',
    supplierAddress: '',
    supplierContact: '',
    supplierEmail: '',
    responseDueDate: '',
    showEndUser: false,
    sourceName: '',
    currency: '',
    incoterm: '',
    destination: '',
  };
  const [uploaded, setUploaded] = useState<ParsedQuoteItem[] | null>(null);
  const [busy, setBusy] = useState(false);

  const items = useMemo(() => job.items.filter((it) => !isEmptyRow(it)), [job.items]);
  const emptyRowCount = job.items.length - items.length;
  const today = new Date().toISOString().slice(0, 10);
  const rfqNumber =
    rfq.rfqNumber || (job.jobSequence ? `SRFQ-${String(job.jobSequence).padStart(3, '0')}` : 'SRFQ-—');
  const endUser = job.controls.endUserSelect === 'OTHER' ? job.controls.customEndUser : job.controls.endUserSelect;
  const currency = rfq.currency || job.controls.currency || 'USD';
  const incoterm = rfq.incoterm || job.controls.incoterm || '';
  const destination = rfq.destination || job.controls.destination || '';

  const handleFile = async (file: File) => {
    setBusy(true);
    try {
      const quote = file.name.toLowerCase().endsWith('.pdf')
        ? await parseSupplierPdfFile(file)
        : parseGenericSupplierXlsx(
            await (async () => {
              const wb = new ExcelJS.Workbook();
              await wb.xlsx.load(await file.arrayBuffer());
              return wb;
            })(),
            file.name
          );
      const parsed = quote.items.filter((i) => i.description && i.description.trim());
      if (parsed.length === 0) {
        setUploaded(null);
        onNotify('No item lines were found in that file. Try the .xlsx version of the end-user RFQ.');
        return;
      }
      setUploaded(parsed);
      onUpdateRfq({ sourceName: file.name });
    } catch {
      setUploaded(null);
      onNotify('Could not read that file. Upload the end-user RFQ as .pdf or .xlsx.');
    } finally {
      setBusy(false);
    }
  };

  const fillFromUpload = () => {
    if (!uploaded) return;
    const fillable = Math.min(uploaded.length, emptyRowCount);
    onFillEmptyRows(
      uploaded.slice(0, fillable).map((i) => ({
        requestedDescription: i.description.trim(),
        uom: i.uom || 'PCS',
        qty: i.qty,
      }))
    );
    onNotify(
      uploaded.length > emptyRowCount
        ? `Filled ${fillable} empty rows. ${uploaded.length - fillable} lines did not fit (rows 17–36 are full).`
        : `Filled ${fillable} rows from ${rfq.sourceName || 'the uploaded RFQ'}.`
    );
    setUploaded(null);
  };

  const field = (label: string, value: string, onChange: (v: string) => void, placeholder = '') => (
    <label className="block space-y-1">
      <span className="text-xs font-medium text-slate-600">{label}</span>
      <input
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className="w-full min-h-[38px] rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
      />
    </label>
  );

  return (
    <div className="space-y-5">
      {/* Controls (not printed) */}
      <div className="no-print rounded-lg border border-slate-200 bg-white p-4 sm:p-5 space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <h2 className="text-sm font-semibold text-slate-900">Supplier RFQ</h2>
            <p className="text-xs text-slate-500">
              Built from the end-user RFQ. Send it to suppliers on your letterhead. The end-user name is hidden unless you turn it on.
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <label className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-3.5 py-2 text-xs font-semibold text-slate-900 hover:bg-slate-50 cursor-pointer min-h-[40px]">
              <Upload className="w-3.5 h-3.5" />
              <span>{busy ? 'Reading file…' : 'Upload end-user RFQ (.pdf / .xlsx)'}</span>
              <input
                type="file"
                accept=".pdf,.xlsx,application/pdf"
                className="hidden"
                disabled={busy}
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) handleFile(f);
                  e.target.value = '';
                }}
              />
            </label>
            <button
              type="button"
              onClick={() => window.print()}
              className="inline-flex items-center gap-2 rounded-lg px-4 py-2 text-xs font-semibold text-white hover:opacity-90 cursor-pointer min-h-[40px]"
              style={{ backgroundColor: theme.primary }}
            >
              <Printer className="w-3.5 h-3.5" />
              <span>Print / Save as PDF</span>
            </button>
            <button
              type="button"
              onClick={async () => {
                try {
                  const file = await exportSupplierRfqXlsx(job, theme, company);
                  onNotify(`Downloaded ${file}`);
                } catch {
                  onNotify('Could not build the Supplier RFQ workbook.');
                }
              }}
              className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-4 py-2 text-xs font-semibold text-slate-900 hover:bg-slate-50 cursor-pointer min-h-[40px]"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>Export Excel</span>
            </button>
          </div>
        </div>

        {uploaded && (
          <div className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-amber-200 bg-amber-50 px-4 py-3 text-xs text-amber-900">
            <span>
              Found {uploaded.length} item lines in {rfq.sourceName}. {emptyRowCount} empty rows are available. Existing rows are not changed.
            </span>
            <div className="flex gap-2">
              <button type="button" onClick={() => setUploaded(null)} className="rounded border border-amber-300 bg-white px-3 py-1.5 font-medium cursor-pointer">
                Cancel
              </button>
              <button type="button" onClick={fillFromUpload} disabled={emptyRowCount === 0} className="rounded bg-amber-900 px-3 py-1.5 font-semibold text-white cursor-pointer disabled:opacity-50">
                Fill empty rows
              </button>
            </div>
          </div>
        )}

        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
          {field('Supplier name', rfq.supplierName, (v) => onUpdateRfq({ supplierName: v }))}
          {field('Supplier contact person', rfq.supplierContact, (v) => onUpdateRfq({ supplierContact: v }))}
          {field('Supplier address', rfq.supplierAddress, (v) => onUpdateRfq({ supplierAddress: v }))}
          {field('Supplier email', rfq.supplierEmail, (v) => onUpdateRfq({ supplierEmail: v }))}
          <label className="block space-y-1">
            <span className="text-xs font-medium text-slate-600">Currency</span>
            <select
              value={rfq.currency || job.controls.currency || 'USD'}
              onChange={(e) => onUpdateRfq({ currency: e.target.value })}
              className="w-full min-h-[38px] rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
            >
              {CURRENCY_OPTIONS.map((c) => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </label>
          <label className="block space-y-1">
            <span className="text-xs font-medium text-slate-600">Incoterm</span>
            <select
              value={rfq.incoterm || job.controls.incoterm || ''}
              onChange={(e) => onUpdateRfq({ incoterm: e.target.value })}
              className="w-full min-h-[38px] rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
            >
              {INCOTERM_OPTIONS.map((inc) => (
                <option key={inc} value={inc}>{inc}</option>
              ))}
            </select>
          </label>
          {field('Delivery destination', rfq.destination || job.controls.destination, (v) => onUpdateRfq({ destination: v }))}
          {field('RFQ number', rfq.rfqNumber, (v) => onUpdateRfq({ rfqNumber: v }), rfqNumber)}
          <label className="block space-y-1">
            <span className="text-xs font-medium text-slate-600">Quotes due by</span>
            <input
              type="date"
              value={rfq.responseDueDate}
              onChange={(e) => onUpdateRfq({ responseDueDate: e.target.value })}
              className="w-full min-h-[38px] rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-900 focus:border-slate-900 focus:outline-none"
            />
          </label>
        </div>
        <label className="inline-flex items-center gap-2 text-xs text-slate-700">
          <input
            type="checkbox"
            checked={rfq.showEndUser}
            onChange={(e) => onUpdateRfq({ showEndUser: e.target.checked })}
          />
          Show the end-user name on this RFQ
        </label>
      </div>

      {/* Printable letter (letterhead, details, items, terms) */}
      <div id="rfq-print" className="rounded-lg border border-slate-200 bg-white p-6 sm:p-10 text-slate-900 space-y-6">
        <header className="flex items-start justify-between gap-6 border-b-2 pb-4" style={{ borderColor: theme.primary }}>
          <div className="flex items-center gap-4 min-w-0">
            {company.logoDataUrl && <img src={company.logoDataUrl} alt="" className="w-20 h-20 object-contain shrink-0" />}
            <div className="min-w-0">
              <div className="text-xl font-bold tracking-wide" style={{ color: theme.primary }}>
                {company.name}
              </div>
              <div className="text-xs text-slate-600">{company.tagline}</div>
            </div>
          </div>
          <div className="text-right text-[11px] leading-relaxed text-slate-600 shrink-0">
            <div>{company.addressLine1}</div>
            <div>{company.addressLine2}</div>
            <div>{company.phone} · {company.email}</div>
            <div>{company.trn}</div>
          </div>
        </header>

        <div className="flex flex-wrap items-start justify-between gap-6">
          <div className="min-w-0 space-y-1">
            <div className="text-[10px] font-semibold uppercase tracking-wider text-slate-500">To</div>
            <div className="text-sm font-semibold">{rfq.supplierName || '—'}</div>
            {rfq.supplierContact && <div className="text-sm">Attn: {rfq.supplierContact}</div>}
            {rfq.supplierAddress && <div className="text-sm whitespace-pre-line">{rfq.supplierAddress}</div>}
            {rfq.supplierEmail && <div className="text-sm">{rfq.supplierEmail}</div>}
          </div>
          <div className="text-xs space-y-1 min-w-[220px]">
            <div className="text-base font-bold tracking-wide" style={{ color: theme.primary }}>REQUEST FOR QUOTATION</div>
            <div><span className="text-slate-500">RFQ No.:</span> <span className="font-semibold">{rfqNumber}</span></div>
            <div><span className="text-slate-500">Date:</span> {today}</div>
            <div><span className="text-slate-500">Quotes due by:</span> {rfq.responseDueDate || '—'}</div>
            <div><span className="text-slate-500">Your reference:</span> {job.controls.rfqRef || '—'}</div>
            {rfq.showEndUser && endUser && <div><span className="text-slate-500">End user:</span> {endUser}</div>}
          </div>
        </div>

        <p className="text-sm">
          Please quote for the items below. Quote unit price and total in <strong>{currency}</strong>, delivered{' '}
          <strong>{incoterm} {destination}</strong>, with lead time, validity and country of origin.
          Include certificates of analysis or safety data sheets where they apply.
        </p>

        <table className="w-full border-collapse text-xs">
          <thead>
            <tr style={{ backgroundColor: theme.primary, color: theme.primaryText }}>
              <th className="py-2 px-2 text-center w-10">No.</th>
              <th className="py-2 px-2 text-left">Description</th>
              <th className="py-2 px-2 text-center w-16">UoM</th>
              <th className="py-2 px-2 text-right w-16">Qty</th>
              <th className="py-2 px-2 text-right w-28">Unit price ({currency})</th>
              <th className="py-2 px-2 text-right w-28">Total ({currency})</th>
            </tr>
          </thead>
          <tbody>
            {items.map((it, i) => (
              <tr key={it.rowNumber} className="border-b border-slate-200">
                <td className="py-2 px-2 text-center">{i + 1}</td>
                <td className="py-2 px-2">{it.requestedDescription || it.offeredDescription}</td>
                <td className="py-2 px-2 text-center">{it.uom}</td>
                <td className="py-2 px-2 text-right font-mono-tabular">{it.qty ?? ''}</td>
                <td className="py-2 px-2 h-9 border-l border-slate-200" />
                <td className="py-2 px-2 h-9 border-l border-slate-200" />
              </tr>
            ))}
            {items.length === 0 && (
              <tr>
                <td colSpan={6} className="py-8 text-center text-slate-400">
                  No items yet. Upload the end-user RFQ above or enter items in the Internal Quote.
                </td>
              </tr>
            )}
          </tbody>
        </table>

        <div className="flex flex-wrap justify-between gap-6 pt-6 text-xs">
          <div>
            <div className="h-10 border-b border-slate-400 w-56" />
            <div className="pt-1 text-slate-600">{job.controls.authorizedSignatory || 'Authorized signatory'}</div>
            <div className="text-slate-500">{company.name}</div>
          </div>
        </div>
      </div>
    </div>
  );
};
