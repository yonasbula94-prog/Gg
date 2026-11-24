import React, { useEffect } from 'react';
import { AlertTriangle } from 'lucide-react';
import { EXCEL_THEMES } from '../constants/themes';
import { ConfirmDialogRequest, ThemeVariantId } from '../types/procurement';

interface ConfirmDialogProps {
  request: ConfirmDialogRequest | null;
  themeId: ThemeVariantId;
  onClose: () => void;
}

export const ConfirmDialog: React.FC<ConfirmDialogProps> = ({
  request,
  themeId,
  onClose,
}) => {
  useEffect(() => {
    if (!request) return;
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKey);
    return () => window.removeEventListener('keydown', handleKey);
  }, [request, onClose]);

  if (!request) return null;

  const theme = EXCEL_THEMES[themeId];
  const isDestructive = request.destructive !== false;

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="confirm-dialog-title"
      className="no-print fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-[1px] p-4"
      onClick={onClose}
    >
      <div
        className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-6 shadow-xl space-y-4"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-start gap-3">
          <div
            className={`w-9 h-9 rounded-lg flex items-center justify-center shrink-0 ${
              isDestructive ? 'bg-red-50 text-red-600' : 'bg-slate-100 text-slate-700'
            }`}
          >
            <AlertTriangle className="w-5 h-5" />
          </div>
          <div className="min-w-0">
            <h3
              id="confirm-dialog-title"
              className="text-base font-semibold text-slate-900"
            >
              {request.title}
            </h3>
            <p className="mt-1.5 text-xs text-slate-600 leading-relaxed">
              {request.message}
            </p>
          </div>
        </div>

        <div className="pt-2 flex justify-end gap-2">
          <button
            type="button"
            onClick={onClose}
            className="rounded-lg border border-slate-200 bg-white px-4 py-2 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer min-h-[38px]"
          >
            {request.cancelLabel || 'Cancel'}
          </button>
          <button
            type="button"
            onClick={() => {
              const fn = request.onConfirm;
              onClose();
              fn();
            }}
            className={`rounded-lg px-4 py-2 text-xs font-semibold text-white cursor-pointer min-h-[38px] transition-opacity hover:opacity-90 ${
              isDestructive ? 'bg-red-600 hover:bg-red-700' : ''
            }`}
            style={!isDestructive ? { backgroundColor: theme.primary } : undefined}
          >
            {request.confirmLabel || 'Confirm'}
          </button>
        </div>
      </div>
    </div>
  );
};
