import React, { useState } from 'react';
import { Download, Monitor, Share, Tablet, X } from 'lucide-react';
import { useOnlineStatus, usePWAInstall } from '../hooks/usePWAInstall';
import { downloadOfflinePackage } from '../utils/offlineDownload';

export const PWAInstallButton: React.FC = () => {
  const { isInstallable, isInstalled, isIOS, install } = usePWAInstall();
  const [showIOSGuide, setShowIOSGuide] = useState(false);
  const [downloading, setDownloading] = useState<'windows' | 'ipad' | null>(null);
  const [downloadErr, setDownloadErr] = useState<string | null>(null);

  const isLocalFile =
    typeof window !== 'undefined' && window.location.protocol === 'file:';

  if (isLocalFile) {
    return (
      <div className="text-[11px] font-medium text-emerald-700 text-center py-1">
        Running Offline Standalone File
      </div>
    );
  }

  const handleDownload = async (target: 'windows' | 'ipad') => {
    setDownloading(target);
    setDownloadErr(null);
    try {
      await downloadOfflinePackage(target);
      if (target === 'ipad') {
        setShowIOSGuide(true);
      }
    } catch {
      setDownloadErr('Download blocked—try in Excel Studio.');
    } finally {
      setDownloading(null);
    }
  };

  return (
    <div className="space-y-1.5">
      <button
        type="button"
        disabled={downloading !== null}
        onClick={() => handleDownload('windows')}
        className="w-full flex items-center justify-center gap-2 rounded-lg bg-slate-900 px-3 py-2 text-xs font-semibold text-white hover:bg-slate-800 transition-colors cursor-pointer min-h-[38px] whitespace-nowrap disabled:opacity-60"
      >
        <Monitor className="w-3.5 h-3.5 shrink-0" />
        <span>{downloading === 'windows' ? 'Preparing...' : 'Windows Offline (.zip)'}</span>
      </button>

      <button
        type="button"
        disabled={downloading !== null}
        onClick={() => (isIOS && !isInstalled ? setShowIOSGuide(true) : handleDownload('ipad'))}
        className="w-full flex items-center justify-center gap-2 rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs font-medium text-slate-800 hover:bg-slate-50 transition-colors cursor-pointer min-h-[38px] whitespace-nowrap disabled:opacity-60"
      >
        <Tablet className="w-3.5 h-3.5 shrink-0" />
        <span>{downloading === 'ipad' ? 'Preparing...' : 'iPad / iOS Version (.zip)'}</span>
      </button>

      {isInstallable && !isInstalled && (
        <button
          type="button"
          onClick={install}
          className="w-full flex items-center justify-center gap-2 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 cursor-pointer min-h-[34px] whitespace-nowrap"
        >
          <Download className="w-3.5 h-3.5 shrink-0" />
          <span>Install Browser App</span>
        </button>
      )}

      {downloadErr && (
        <div className="text-[11px] text-red-600 text-center">{downloadErr}</div>
      )}

      {showIOSGuide && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4">
          <div className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-6 shadow-xl space-y-4">
            <div className="flex items-center justify-between">
              <h3 className="text-base font-semibold text-slate-900">
                How to Run Offline on iPad / iOS
              </h3>
              <button
                type="button"
                onClick={() => setShowIOSGuide(false)}
                className="p-1.5 rounded-md text-slate-400 hover:text-slate-700 hover:bg-slate-100 cursor-pointer"
                aria-label="Close iPad guide"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="rounded-lg border border-amber-200 bg-amber-50 p-3 text-xs text-amber-900">
              <strong>Why iPad needs a 1-time 30-second setup:</strong> Apple&apos;s built-in <strong>Files</strong> app blocks code inside local <code>.html</code> files (blank screen). Hosting the unzipped app folder once gives you a private link that installs onto your iPad Home Screen and then works 100% offline in Airplane Mode.
            </div>

            <ol className="space-y-3 text-xs text-slate-700 list-decimal pl-4">
              <li>
                <strong className="text-slate-900">Download &amp; Unzip:</strong> Click below to download <code>Grand_Group_iPad_App.zip</code>, then unzip it. You will see the clean app folder <code>Grand_Group_iPad_App</code> plus a separate <code>Tutorial/</code> attachment folder.
              </li>
              <li>
                <strong className="text-slate-900">Get Your Private Link (Free, 10 sec):</strong> Open <code>app.netlify.com/drop</code> on a laptop/PC and drag the unzipped <strong><code>Grand_Group_iPad_App</code> folder</strong> into the box. Copy the <code>https://...netlify.app</code> link it gives you.
              </li>
              <li>
                <strong className="text-slate-900">Install on iPad:</strong> Open that link in <strong>Safari</strong> on your iPad → tap{' '}
                <span className="inline-flex items-center gap-0.5 font-semibold text-slate-900"><Share className="w-3 h-3" /> Share</span> → <strong>Add to Home Screen</strong>. Done — it opens full-screen with your Grand Group logo even with Wi-Fi turned off.
              </li>
            </ol>

            <div className="flex items-center gap-2 pt-1">
              <button
                type="button"
                onClick={() => handleDownload('ipad')}
                className="flex-1 rounded-lg border border-slate-200 py-2.5 text-xs font-semibold text-slate-800 hover:bg-slate-50 cursor-pointer"
              >
                Download iPad App (.zip)
              </button>
              <button
                type="button"
                onClick={() => setShowIOSGuide(false)}
                className="flex-1 rounded-lg bg-slate-900 py-2.5 text-xs font-semibold text-white hover:bg-slate-800 cursor-pointer"
              >
                Done
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export const OfflineIndicator: React.FC = () => {
  const isOnline = useOnlineStatus();
  if (isOnline) return null;

  return (
    <div className="no-print fixed bottom-20 lg:bottom-4 right-4 z-50 flex items-center gap-2 rounded-lg border border-amber-300 bg-amber-50 px-3.5 py-2 text-xs font-medium text-amber-900 shadow-sm">
      <span>Offline Mode — All jobs, calculations, and Excel exports run locally.</span>
    </div>
  );
};
