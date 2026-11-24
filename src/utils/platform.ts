import { PlatformMode } from '../types/procurement';

declare global {
  interface Window {
    __GG_DEFAULT_PLATFORM__?: PlatformMode;
  }
}

export function isIOSDevice(): boolean {
  if (typeof navigator === 'undefined') return false;
  const ua = navigator.userAgent.toLowerCase();
  const isIPadOS =
    navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1;
  return /iphone|ipad|ipod/.test(ua) || isIPadOS;
}

export function detectDefaultPlatform(): PlatformMode {
  if (typeof window === 'undefined') return 'windows';

  // 1. URL query override (?mode=ipad or ?mode=windows)
  try {
    const params = new URLSearchParams(window.location.search);
    const paramMode = params.get('mode');
    if (paramMode === 'ipad' || paramMode === 'windows') return paramMode;
  } catch {
    // Ignore on file://
  }

  // 2. Saved user preference
  try {
    const saved = localStorage.getItem('gg_fzc_platform_mode_v1');
    if (saved === 'ipad' || saved === 'windows') return saved;
  } catch {
    // Ignore
  }

  // 3. Build-time baked default (Windows vs iPad offline package)
  if (
    window.__GG_DEFAULT_PLATFORM__ === 'ipad' ||
    window.__GG_DEFAULT_PLATFORM__ === 'windows'
  ) {
    return window.__GG_DEFAULT_PLATFORM__;
  }

  // 4. Hardware auto-detection
  return isIOSDevice() ? 'ipad' : 'windows';
}

export function savePlatformMode(mode: PlatformMode): void {
  try {
    localStorage.setItem('gg_fzc_platform_mode_v1', mode);
  } catch {
    // Ignore
  }
}

/**
 * On iPadOS / iOS Safari, triggers the native Share Sheet ("Open in Excel", "Save to Files", AirDrop)
 * when Web Share Level 2 supports files; falls back cleanly to standard Blob download on Windows/Mac.
 */
export async function shareOrDownloadBlob(
  blob: Blob,
  filename: string,
  mimeType: string
): Promise<void> {
  if (isIOSDevice() && typeof navigator !== 'undefined' && typeof navigator.share === 'function') {
    try {
      const file = new File([blob], filename, { type: mimeType });
      if (typeof navigator.canShare === 'function' && navigator.canShare({ files: [file] })) {
        await navigator.share({
          files: [file],
          title: filename,
        });
        return;
      }
    } catch (err) {
      // If user explicitly cancelled the iOS share sheet, do not force fallback download
      if (err instanceof DOMException && err.name === 'AbortError') {
        return;
      }
    }
  }

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}
