import { shareOrDownloadBlob } from './platform';

export async function downloadOfflinePackage(target: 'windows' | 'ipad'): Promise<string> {
  const endpoint = target === 'ipad' ? '/offline-ipad.dat' : '/offline-windows.dat';
  const res = await fetch(endpoint, { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }

  const buf = await res.arrayBuffer();
  const bytes = new Uint8Array(buf);
  // Verify ZIP magic header ("PK\x03\x04")
  if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    throw new Error(`${target === 'ipad' ? 'iPad' : 'Windows'} offline zip package unavailable`);
  }

  const filename =
    target === 'ipad'
      ? 'Grand_Group_iPad_App.zip'
      : 'Grand_Group_FZC_Windows_Offline.zip';
  const blob = new Blob([buf], { type: 'application/zip' });
  await shareOrDownloadBlob(blob, filename, 'application/zip');
  return filename;
}
