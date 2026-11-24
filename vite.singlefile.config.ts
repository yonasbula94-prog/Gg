import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';
import path from 'path';
import { defineConfig } from 'vite';
import { viteSingleFile } from 'vite-plugin-singlefile';

interface ZipEntry {
  name: string;
  data: Buffer;
}

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) !== 0 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(buf: Buffer): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function createStoredZip(entries: ZipEntry[]): Buffer {
  const localParts: Buffer[] = [];
  const centralParts: Buffer[] = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBuf = Buffer.from(entry.name, 'utf8');
    const dataBuf = entry.data;
    const crc = crc32(dataBuf);

    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0, 6);
    localHeader.writeUInt16LE(0, 8); // store (no compression)
    localHeader.writeUInt16LE(0, 10);
    localHeader.writeUInt16LE(0, 12);
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(dataBuf.length, 18);
    localHeader.writeUInt32LE(dataBuf.length, 22);
    localHeader.writeUInt16LE(nameBuf.length, 26);
    localHeader.writeUInt16LE(0, 28);

    localParts.push(localHeader, nameBuf, dataBuf);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0, 8);
    centralHeader.writeUInt16LE(0, 10);
    centralHeader.writeUInt16LE(0, 12);
    centralHeader.writeUInt16LE(0, 14);
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(dataBuf.length, 20);
    centralHeader.writeUInt32LE(dataBuf.length, 24);
    centralHeader.writeUInt16LE(nameBuf.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);

    centralParts.push(centralHeader, nameBuf);
    offset += localHeader.length + nameBuf.length + dataBuf.length;
  }

  const centralDir = Buffer.concat(centralParts);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralDir.length, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);

  return Buffer.concat([...localParts, centralDir, eocd]);
}

const IPAD_SW_JS = `const CACHE_NAME = 'gg-procure-ipad-v1';
const ASSETS = ['./', './index.html', './manifest.json', './apple-touch-icon.png'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request)
        .then((res) => {
          if (res && res.status === 200 && res.type === 'basic') {
            const copy = res.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          }
          return res;
        })
        .catch(() => caches.match('./index.html'));
    })
  );
});
`;

const IPAD_MANIFEST_JSON = JSON.stringify(
  {
    name: 'Grand Group FZC Procurement',
    short_name: 'GG Procure',
    description: 'Offline iPad Procurement Document System for Grand Group FZC',
    start_url: './index.html',
    scope: './',
    display: 'standalone',
    background_color: '#f8fafc',
    theme_color: '#1B365D',
    icons: [
      {
        src: './apple-touch-icon.png',
        sizes: '306x320',
        type: 'image/png',
        purpose: 'any',
      },
    ],
  },
  null,
  2
);

export default defineConfig({
  plugins: [
    react(),
    tailwindcss(),
    viteSingleFile(),
    {
      name: 'output-dual-platform-offline-bundles',
      apply: 'build',
      closeBundle(error?: Error) {
        if (error) return;
        const builtHtmlPath = path.resolve(import.meta.dirname, 'dist-offline/index.html');
        const oldSingle = path.resolve(import.meta.dirname, 'public/offline-app.dat');
        if (fs.existsSync(oldSingle)) fs.unlinkSync(oldSingle);

        if (fs.existsSync(builtHtmlPath)) {
          let html = fs.readFileSync(builtHtmlPath, 'utf8');
          const scriptStart = html.indexOf('<script');
          const scriptTagEnd = scriptStart !== -1 ? html.indexOf('>', scriptStart) : -1;
          const scriptClose = scriptTagEnd !== -1 ? html.indexOf('</script>', scriptTagEnd) : -1;

          if (scriptStart !== -1 && scriptTagEnd !== -1 && scriptClose !== -1) {
            const scriptCode = html.slice(scriptTagEnd + 1, scriptClose);
            html = html.slice(0, scriptStart) + html.slice(scriptClose + '</script>'.length);
            html = html.replace(/<link[^>]+href="\/(apple-touch-icon\.png|icon\.svg)"[^>]*>/gi, '');

            const bodyEnd = html.lastIndexOf('</body>');
            const headEnd = html.indexOf('</head>');

            if (bodyEnd !== -1 && headEnd !== -1) {
              const publicDir = path.resolve(import.meta.dirname, 'public');
              fs.mkdirSync(publicDir, { recursive: true });

              // Load separate Tutorial/ attachments from tutorial-attachments/ (outside public/)
              const tutorialDir = path.resolve(import.meta.dirname, 'tutorial-attachments');
              const tutorialFiles = [
                'Grand_Group_Procurement_NotebookLM_Tutorial_Guide.pdf',
                'Grand_Group_Procurement_Demo_Walkthrough.webm',
                'Grand_Group_Procurement_NotebookLM_Guide.md',
              ];
              const readTutorialEntries = (prefix: string): ZipEntry[] => {
                const entries: ZipEntry[] = [];
                for (const fileName of tutorialFiles) {
                  const fullPath = path.join(tutorialDir, fileName);
                  if (fs.existsSync(fullPath)) {
                    entries.push({
                      name: `${prefix}${fileName}`,
                      data: fs.readFileSync(fullPath),
                    });
                  }
                }
                return entries;
              };

              // 1) Windows offline .zip (clean double-clickable .html + separate Tutorial/ folder)
              const windowsHtml =
                html.slice(0, bodyEnd) +
                '<script>window.__GG_DEFAULT_PLATFORM__="windows";</script>' +
                '<script>' +
                scriptCode +
                '</script>' +
                html.slice(bodyEnd);

              const windowsZipBuffer = createStoredZip([
                {
                  name: 'Grand_Group_FZC_Windows_Offline/Grand_Group_FZC_Windows_Offline.html',
                  data: Buffer.from(windowsHtml, 'utf8'),
                },
                ...readTutorialEntries('Grand_Group_FZC_Windows_Offline/Tutorial/'),
              ]);

              fs.writeFileSync(path.join(publicDir, 'offline-windows.dat'), windowsZipBuffer);

              // 2) iPad / iOS complete offline PWA folder + separate Tutorial/ folder packaged in a .zip
              const ipadHeadInject =
                '<link rel="manifest" href="./manifest.json" />' +
                '<link rel="apple-touch-icon" href="./apple-touch-icon.png" />' +
                '<link rel="icon" type="image/png" href="./apple-touch-icon.png" />';

              const ipadWithHead =
                html.slice(0, headEnd) + ipadHeadInject + html.slice(headEnd);
              const ipadBodyEnd = ipadWithHead.lastIndexOf('</body>');

              const ipadSwBoot =
                'window.__GG_DEFAULT_PLATFORM__="ipad";' +
                'if("serviceWorker" in navigator && (location.protocol==="https:" || location.hostname==="localhost")){' +
                'window.addEventListener("load",function(){navigator.serviceWorker.register("./sw.js").catch(function(){});});' +
                '}';

              const ipadHtml =
                ipadWithHead.slice(0, ipadBodyEnd) +
                `<script>${ipadSwBoot}</script>` +
                '<script>' +
                scriptCode +
                '</script>' +
                ipadWithHead.slice(ipadBodyEnd);

              const logoPath = path.resolve(import.meta.dirname, 'src/assets/grand-group-logo.png');
              const iconBuf = fs.readFileSync(logoPath);

              const ipadZipBuffer = createStoredZip([
                { name: 'Grand_Group_iPad_App/index.html', data: Buffer.from(ipadHtml, 'utf8') },
                { name: 'Grand_Group_iPad_App/sw.js', data: Buffer.from(IPAD_SW_JS, 'utf8') },
                {
                  name: 'Grand_Group_iPad_App/manifest.json',
                  data: Buffer.from(IPAD_MANIFEST_JSON, 'utf8'),
                },
                { name: 'Grand_Group_iPad_App/apple-touch-icon.png', data: iconBuf },
                ...readTutorialEntries('Tutorial/'),
              ]);

              fs.writeFileSync(path.join(publicDir, 'offline-ipad.dat'), ipadZipBuffer);
            }
          }
          fs.rmSync(path.resolve(import.meta.dirname, 'dist-offline'), {
            recursive: true,
            force: true,
          });
        }
      },
    },
  ],
  resolve: {
    alias: {
      '@': path.resolve(import.meta.dirname, '.'),
      'virtual:pwa-register': path.resolve(import.meta.dirname, 'src/utils/pwaStub.ts'),
    },
  },
  publicDir: false,
  build: {
    outDir: 'dist-offline',
    emptyOutDir: true,
    assetsInlineLimit: 100000000,
    cssCodeSplit: false,
    rollupOptions: {
      output: {
        format: 'iife',
      },
    },
  },
});
