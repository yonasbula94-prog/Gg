import React from 'react';
import { AlertTriangle, HardDriveDownload, RefreshCw, RotateCcw } from 'lucide-react';
import { clearAllAppStorage, exportEmergencyRawStorageJson } from '../store/storage';

interface AppErrorBoundaryProps {
  children: React.ReactNode;
}

interface AppErrorBoundaryState {
  hasError: boolean;
  errorMessage: string;
  confirmReset: boolean;
}

export class AppErrorBoundary extends React.Component<
  AppErrorBoundaryProps,
  AppErrorBoundaryState
> {
  constructor(props: AppErrorBoundaryProps) {
    super(props);
    this.state = {
      hasError: false,
      errorMessage: '',
      confirmReset: false,
    };
  }

  static getDerivedStateFromError(error: Error): Partial<AppErrorBoundaryState> {
    return {
      hasError: true,
      errorMessage: error?.message || 'An unexpected error interrupted the workspace.',
    };
  }

  componentDidCatch(error: Error): void {
    // Kept concise for diagnostics if inspected via dev tools
    console.error('[Grand Group Procurement] Workspace recovery triggered:', error);
  }

  private handleReload = () => {
    window.location.reload();
  };

  private handleExportEmergencyBackup = () => {
    try {
      const json = exportEmergencyRawStorageJson();
      const blob = new Blob([json], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `GG_Procurement_Emergency_Backup_${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch {
      // Ignore if downloads are blocked
    }
  };

  private handleResetStorage = () => {
    clearAllAppStorage();
    window.location.reload();
  };

  render() {
    if (!this.state.hasError) {
      return this.props.children;
    }

    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-6">
        <div className="w-full max-w-lg rounded-xl border border-slate-200 bg-white p-6 shadow-sm space-y-5">
          <div className="flex items-start gap-3">
            <div className="w-10 h-10 rounded-lg bg-red-50 text-red-600 flex items-center justify-center shrink-0">
              <AlertTriangle className="w-5 h-5" />
            </div>
            <div>
              <h1 className="text-base font-bold text-slate-900">
                Workspace Recovery Mode
              </h1>
              <p className="mt-1 text-xs text-slate-600 leading-relaxed">
                Grand Group Procurement Workspace paused to protect your saved jobs and registry data. You can reload safely or export an emergency JSON backup below.
              </p>
            </div>
          </div>

          {this.state.errorMessage && (
            <div className="rounded-lg border border-slate-200 bg-slate-50 px-3.5 py-2.5 text-xs font-mono text-slate-700 break-words">
              {this.state.errorMessage}
            </div>
          )}

          <div className="flex flex-wrap items-center gap-2.5">
            <button
              type="button"
              onClick={this.handleReload}
              className="inline-flex items-center gap-1.5 rounded-lg bg-slate-900 px-4 py-2.5 text-xs font-semibold text-white hover:bg-slate-800 cursor-pointer min-h-[40px]"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>Reload Workspace</span>
            </button>
            <button
              type="button"
              onClick={this.handleExportEmergencyBackup}
              className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-4 py-2.5 text-xs font-medium text-slate-800 hover:bg-slate-50 cursor-pointer min-h-[40px]"
            >
              <HardDriveDownload className="w-3.5 h-3.5" />
              <span>Export Emergency JSON</span>
            </button>
          </div>

          <div className="pt-4 border-t border-slate-100">
            {!this.state.confirmReset ? (
              <button
                type="button"
                onClick={() => this.setState({ confirmReset: true })}
                className="inline-flex items-center gap-1.5 text-xs font-medium text-red-600 hover:text-red-800 cursor-pointer"
              >
                <RotateCcw className="w-3.5 h-3.5" />
                <span>Reset corrupted local storage and start fresh</span>
              </button>
            ) : (
              <div className="rounded-lg border border-red-200 bg-red-50 p-3.5 space-y-2.5">
                <p className="text-xs text-red-900 font-medium">
                  This clears locally stored jobs and resets the workspace to a blank job. Export an Emergency JSON backup first if you want a copy.
                </p>
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={this.handleResetStorage}
                    className="rounded-lg bg-red-600 px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-red-700 cursor-pointer"
                  >
                    Yes, Reset &amp; Restart
                  </button>
                  <button
                    type="button"
                    onClick={() => this.setState({ confirmReset: false })}
                    className="rounded-lg border border-slate-300 bg-white px-3.5 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50 cursor-pointer"
                  >
                    Cancel
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }
}
