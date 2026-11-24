import React, { useState } from 'react';
import { ClipboardPaste } from 'lucide-react';

export interface PasteField {
  key: string;
  label: string;
  /** Converts the pasted text to the stored value. Returns undefined to skip an invalid entry. */
  parse: (raw: string) => unknown;
}

interface PasteIntoRowsProps {
  fields: PasteField[];
  /** Selected row numbers, in table order */
  rowIds: number[];
  disabled?: boolean;
  onApply: (fieldKey: string, values: { rowId: number; value: unknown }[]) => void;
}

/**
 * Copy values from Excel (or any text) and apply them to the selected rows, in order.
 * One pasted value applies to every selected row; several values go one per row.
 */
export const PasteIntoRows: React.FC<PasteIntoRowsProps> = ({ fields, rowIds, disabled, onApply }) => {
  const [fieldKey, setFieldKey] = useState(fields[0]?.key ?? '');
  const [manual, setManual] = useState('');
  const [needsManual, setNeedsManual] = useState(false);
  const [message, setMessage] = useState('');
  const [typed, setTyped] = useState('');

  const field = fields.find((f) => f.key === fieldKey) ?? fields[0];

  const apply = (text: string) => {
    if (!field) return;
    const lines = text.replace(/\r/g, '').split('\n').map((l) => l.split('\t')[0].trim()).filter((l, i, arr) => !(l === '' && i === arr.length - 1));
    const sorted = [...rowIds].sort((a, b) => a - b);
    if (lines.length === 0 || sorted.length === 0) {
      setMessage('Nothing to paste. Copy values from Excel first.');
      return;
    }
    const pairs: { rowId: number; value: unknown }[] = [];
    sorted.forEach((rowId, i) => {
      const raw = lines.length === 1 ? lines[0] : lines[i];
      // Empty cells in the copied text leave that row unchanged
      if (raw === undefined || raw === '') return;
      const value = field.parse(raw);
      if (value !== undefined) pairs.push({ rowId, value });
    });
    if (pairs.length === 0) {
      setMessage('The pasted values could not be read for this field.');
      return;
    }
    onApply(field.key, pairs);
    setMessage(`Applied ${pairs.length} value${pairs.length === 1 ? '' : 's'} to ${sorted.length} selected row${sorted.length === 1 ? '' : 's'}.`);
    setManual('');
    setNeedsManual(false);
  };

  const fromClipboard = async () => {
    try {
      const text = await navigator.clipboard.readText();
      apply(text);
    } catch {
      // Clipboard access can be blocked (permissions or an embedded browser): use the paste box instead
      setNeedsManual(true);
      setMessage('Clipboard access is blocked here. Paste into the box below and apply.');
    }
  };

  if (!field) return null;
  return (
    <div className="flex flex-wrap items-center gap-2">
      <select
        value={field.key}
        onChange={(e) => setFieldKey(e.target.value)}
        disabled={disabled}
        aria-label="Field to paste into"
        className="min-h-[32px] rounded border border-slate-200 bg-white px-2 disabled:opacity-50"
      >
        {fields.map((f) => (
          <option key={f.key} value={f.key}>
            {f.label}
          </option>
        ))}
      </select>
      <button
        type="button"
        disabled={disabled || rowIds.length === 0}
        onClick={fromClipboard}
        className="inline-flex items-center gap-1.5 rounded border border-slate-200 bg-white px-2.5 py-1.5 font-medium text-slate-800 hover:bg-slate-100 disabled:opacity-50 cursor-pointer"
      >
        <ClipboardPaste className="w-3.5 h-3.5" />
        <span>Paste into selected</span>
      </button>
      {needsManual && (
        <span className="flex items-center gap-2">
          <input
            value={manual}
            onChange={(e) => setManual(e.target.value)}
            placeholder="Paste here (Ctrl+V)"
            className="min-h-[32px] w-48 rounded border border-slate-200 bg-white px-2"
          />
          <button
            type="button"
            disabled={!manual.trim()}
            onClick={() => apply(manual)}
            className="rounded border border-slate-200 bg-white px-2.5 py-1.5 font-medium text-slate-800 hover:bg-slate-100 disabled:opacity-50 cursor-pointer"
          >
            Apply
          </button>
        </span>
      )}
      <span className="flex items-center gap-2">
        <input
          value={typed}
          onChange={(e) => setTyped(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && typed.trim()) {
              apply(typed);
              setTyped('');
            }
          }}
          disabled={disabled}
          placeholder={`Type a ${field?.label ?? 'value'} for all selected`}
          aria-label="Value to apply to selected rows"
          className="min-h-[32px] w-56 rounded border border-slate-200 bg-white px-2 disabled:opacity-50"
        />
        <button
          type="button"
          disabled={disabled || rowIds.length === 0 || !typed.trim()}
          onClick={() => {
            apply(typed);
            setTyped('');
          }}
          className="rounded border border-slate-200 bg-white px-2.5 py-1.5 font-medium text-slate-800 hover:bg-slate-100 disabled:opacity-50 cursor-pointer"
        >
          Apply to selected
        </button>
      </span>
      {message && <span className="text-slate-600">{message}</span>}
    </div>
  );
};
