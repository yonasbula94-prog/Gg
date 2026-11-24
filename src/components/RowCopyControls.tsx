import React, { useState } from 'react';
import { Copy, ClipboardPaste } from 'lucide-react';
import { JobItemRow } from '../types/procurement';

export type RowFields = Partial<JobItemRow>;

/** Last copied row. Kept per page, so a row copied on Quotation only pastes on Quotation. */
let rowClip: { page: string; sourceRow: number; fields: RowFields } | null = null;

interface RowCopyControlsProps {
  /** Page key, e.g. 'PRICE_CALC', 'INTERNAL_QUOTE', 'DOC-QT'. Paste works only on the same page. */
  page: string;
  /** Selected row numbers, in table order */
  selectedRows: number[];
  disabled?: boolean;
  /** Fields of one row to copy (returns null when the row is not found) */
  readRow: (rowNumber: number) => RowFields | null;
  /** Applies the copied fields to each target row */
  onPaste: (rows: { rowNumber: number; fields: RowFields }[]) => void;
}

/**
 * Copy one ticked row, then paste it into other ticked rows on the same page.
 * Use it when most of a row repeats and only a small part changes per row (for example the offered description).
 */
export const RowCopyControls: React.FC<RowCopyControlsProps> = ({ page, selectedRows, disabled, readRow, onPaste }) => {
  const [note, setNote] = useState('');
  const [, setVersion] = useState(0);
  const clipForPage = rowClip && rowClip.page === page ? rowClip : null;

  const copy = () => {
    if (selectedRows.length !== 1) {
      setNote('Tick exactly one row to copy it.');
      return;
    }
    const row = selectedRows[0];
    const fields = readRow(row);
    if (!fields) {
      setNote('That row could not be read.');
      return;
    }
    rowClip = { page, sourceRow: row, fields };
    setVersion((v) => v + 1);
    setNote(`Copied row ${row}. Tick other rows on this page and paste.`);
  };

  const paste = () => {
    if (!clipForPage) {
      setNote('Copy a row first.');
      return;
    }
    const targets = selectedRows.filter((rn) => rn !== clipForPage.sourceRow);
    if (targets.length === 0) {
      setNote('Tick the rows to paste into (the copied row is skipped).');
      return;
    }
    onPaste(targets.map((rowNumber) => ({ rowNumber, fields: clipForPage.fields })));
    setNote(`Pasted row ${clipForPage.sourceRow} into ${targets.length} row${targets.length === 1 ? '' : 's'}.`);
  };

  return (
    <span className="flex flex-wrap items-center gap-2">
      <button
        type="button"
        onClick={copy}
        disabled={disabled}
        className="inline-flex items-center gap-1.5 rounded border border-slate-200 bg-white px-2.5 py-1.5 font-medium text-slate-800 hover:bg-slate-100 disabled:opacity-50 cursor-pointer"
      >
        <Copy className="w-3.5 h-3.5" />
        <span>Copy row</span>
      </button>
      <button
        type="button"
        onClick={paste}
        disabled={disabled || !clipForPage}
        title={clipForPage ? `Row ${clipForPage.sourceRow} copied` : 'Copy a row first'}
        className="inline-flex items-center gap-1.5 rounded border border-slate-200 bg-white px-2.5 py-1.5 font-medium text-slate-800 hover:bg-slate-100 disabled:opacity-50 cursor-pointer"
      >
        <ClipboardPaste className="w-3.5 h-3.5" />
        <span>Paste row into selected{clipForPage ? ` (row ${clipForPage.sourceRow})` : ''}</span>
      </button>
      {note && <span className="text-slate-600">{note}</span>}
    </span>
  );
};

/** Pick only the listed fields from a row, so a paste never touches other columns. */
export function pickFields(item: JobItemRow | undefined, keys: (keyof JobItemRow)[]): RowFields | null {
  if (!item) return null;
  const out: RowFields = {};
  keys.forEach((k) => {
    (out as Record<string, unknown>)[k] = JSON.parse(JSON.stringify(item[k] ?? null));
  });
  return out;
}
