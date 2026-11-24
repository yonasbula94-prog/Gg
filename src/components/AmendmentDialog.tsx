import React, { useState } from 'react';
import { Lock, Save } from 'lucide-react';

interface AmendmentDialogProps {
  mode: 'unlock' | 'save';
  docNumber: string;
  currentRevision: number;
  otherIssuedDocs: string[];
  verifyPassword: (candidate: string) => boolean;
  onConfirmUnlock: () => void;
  onConfirmSave: (reason: string) => void;
  onCancel: () => void;
}

export const AmendmentDialog: React.FC<AmendmentDialogProps> = ({
  mode,
  docNumber,
  currentRevision,
  otherIssuedDocs,
  verifyPassword,
  onConfirmUnlock,
  onConfirmSave,
  onCancel,
}) => {
  const [password, setPassword] = useState('');
  const [reason, setReason] = useState('');
  const [error, setError] = useState('');

  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    if (mode === 'unlock') {
      if (!verifyPassword(password)) {
        setError('Incorrect admin password.');
        return;
      }
      onConfirmUnlock();
      return;
    }
    if (reason.trim().length < 5) {
      setError('Enter a reason for the change (at least 5 characters).');
      return;
    }
    onConfirmSave(reason.trim());
  };

  return (
    <div className="no-print fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
      <form
        onSubmit={submit}
        className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-6 shadow-xl space-y-4"
      >
        <div className="flex items-center gap-2">
          {mode === 'unlock' ? (
            <Lock className="w-4 h-4 text-slate-700" />
          ) : (
            <Save className="w-4 h-4 text-slate-700" />
          )}
          <h3 className="text-base font-semibold text-slate-900">
            {mode === 'unlock'
              ? `Unlock ${docNumber} to amend`
              : `Save ${docNumber} as Rev ${currentRevision + 1}`}
          </h3>
        </div>

        {mode === 'unlock' ? (
          <>
            <p className="text-xs text-slate-600">
              Issued documents are locked to protect the figures already sent. Enter the admin password to
              open this job for amendment. Every change is logged with a reason.
            </p>
            <input
              type="password"
              autoFocus
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Admin password"
              className="w-full min-h-[40px] rounded-lg border border-slate-200 px-3 py-2 text-sm focus:border-slate-900 focus:outline-none"
            />
          </>
        ) : (
          <>
            <p className="text-xs text-slate-600">
              The previous figures are kept in this document's history. Reason is required and goes into the
              change log.
            </p>
            <textarea
              autoFocus
              rows={3}
              value={reason}
              onChange={(e) => setReason(e.target.value)}
              placeholder="e.g. Supplier revised unit cost; customer reduced quantity"
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm focus:border-slate-900 focus:outline-none"
            />
            {otherIssuedDocs.length > 0 && (
              <p className="rounded-lg border border-amber-200 bg-amber-50 px-3 py-2 text-xs text-amber-900">
                Other issued documents in this job ({otherIssuedDocs.join(', ')}) keep their current revision and
                totals. Amend them separately if they need the same change.
              </p>
            )}
          </>
        )}

        {error && <p className="text-xs text-red-600">{error}</p>}

        <div className="flex justify-end gap-2 pt-1">
          <button
            type="button"
            onClick={onCancel}
            className="rounded-lg border border-slate-200 px-4 py-2 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer"
          >
            Cancel
          </button>
          <button
            type="submit"
            className="rounded-lg bg-slate-900 px-4 py-2 text-xs font-semibold text-white hover:bg-slate-800 cursor-pointer"
          >
            {mode === 'unlock' ? 'Unlock' : 'Save Revision'}
          </button>
        </div>
      </form>
    </div>
  );
};
