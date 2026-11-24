// Captures the Help screenshots from the running app and draws numbered red markers on the real controls.
// Usage: 1) bun -e "..." writes scripts/help-shots.json from src/constants/helpTopics.ts (see package.json "help:shots")
//        2) node scripts/capture-help.cjs [http://localhost:4173]
const { chromium } = require('playwright');
const fs = require('fs');
const os = require('os');
const path = require('path');

const BASE = process.argv[2] || 'http://localhost:4173/?mode=windows';
const SHOTS = JSON.parse(fs.readFileSync(path.join(__dirname, 'help-shots.json'), 'utf8'));
const OUT = path.join(__dirname, '..', 'src', 'assets', 'help');
const SIDEBAR = {
  TRACKER: 'Dashboard', INTERNAL_QUOTE: 'Internal Quote', PRICE_CALC: 'Price Calc (Internal)', QT: 'Quotation (Master)',
  PI: 'Proforma Invoice', PO: 'Supplier PO', CI: 'Commercial Invoice', PL: 'Packing List', DN: 'Delivery Note',
  REGISTRY: 'Master Registry', FILES: 'Import & Export', STUDIO: 'Clients & Settings',
};
// A small text-based quote for the import dialog (same layout the importer reads)
const QUOTE_PDF = (() => {
  const content = [
    'BT /F1 10 Tf 50 700 Td (PN-4402 Hydraulic Seal Kit PCS 10 42.50 425.00) Tj ET',
    'BT /F1 10 Tf 50 680 Td (PN-8810 Pressure Gauge 0-16 Bar PCS 5 68.00 340.00) Tj ET',
    'BT /F1 10 Tf 50 660 Td (Payment Terms: 50% advance) Tj ET',
  ].join('\n');
  return Buffer.from(`%PDF-1.4\n1 0 obj << /Type /Catalog >> endobj\n4 0 obj << /Length ${content.length} >>\nstream\n${content}\nendstream\nendobj\ntrailer << >>\n%%EOF`, 'latin1');
})();

async function seed(page, issued) {
  await page.goto(BASE, { waitUntil: 'networkidle' });
  await page.evaluate((isIssued) => {
    const key = Object.keys(localStorage).find((k) => k.includes('procurement_jobs'));
    const jobs = JSON.parse(localStorage.getItem(key));
    const j = jobs[0];
    j.jobName = 'Hydraulic Seal Kits — MOA';
    j.controls.endUserSelect = 'MOA (Ministry of Agriculture)';
    j.controls.rfqRef = 'RFQ-2026-014';
    [0, 1, 2].forEach((i) => Object.assign(j.items[i], {
      requestedDescription: ['Hydraulic seal kit', 'Pressure gauge 0-16 bar', 'Pump bearing'][i],
      offeredDescription: ['Seal kit HX-40', 'Gauge 0-16 bar, brass', 'Bearing 6205-2RS'][i],
      uom: ['SET', 'PCS', 'PCS'][i], qty: [4, 6, 10][i], supplierUnitCost: [42.5, 68, 3.2][i],
      boxes: [2, 1, 4][i], boxLengthCm: 30, boxWidthCm: 20, boxHeightCm: 15, kgPerBox: 8,
    }));
    if (isIssued) {
      j.issuedDocs = { QT: { prefix: 'QT', docNumber: 'QT-GG-043-2026', sequence: 43, revision: 1, issuedAt: new Date().toISOString(), issuedBy: 'Procurement Desk', frozenSubtotal: 0, frozenVat: 0, frozenFreight: 0, frozenGrandTotal: 0 } };
    }
    localStorage.setItem(key, JSON.stringify(jobs));
  }, issued);
  await page.reload({ waitUntil: 'networkidle' });
}

async function locate(page, find) {
  const strategies = [
    () => page.getByRole('button', { name: find, exact: true }),
    () => page.getByRole('button', { name: find }),
    () => page.getByRole('tab', { name: find }),
    () => page.getByRole('checkbox', { name: find }),
    () => page.getByLabel(find),
    () => page.getByPlaceholder(find),
    () => page.getByText(find),
  ];
  for (const make of strategies) {
    const loc = make();
    const n = await loc.count();
    for (let i = 0; i < Math.min(n, 6); i++) {
      if (await loc.nth(i).isVisible().catch(() => false)) return loc.nth(i);
    }
  }
  return null;
}

async function rectOf(handle) {
  return handle.evaluate((el) => {
    let t = el;
    // A text label usually sits above its field: outline the label and field together
    if (!['BUTTON', 'INPUT', 'SELECT', 'TEXTAREA', 'TH', 'TD', 'TR'].includes(el.tagName) && el.getAttribute('role') !== 'button') {
      let p = el.parentElement;
      for (let i = 0; i < 2 && p; i++) {
        if (p.querySelector('input,select,textarea') && !p.querySelector('th,td')) { t = p; break; }
        p = p.parentElement;
      }
    }
    const r = t.getBoundingClientRect();
    return { x: r.left, y: r.top, w: r.width, h: r.height };
  });
}

async function drawMarkers(page, items) {
  await page.evaluate((list) => {
    document.querySelectorAll('.__help-marker').forEach((n) => n.remove());
    for (const m of list) {
      const box = document.createElement('div');
      box.className = '__help-marker';
      Object.assign(box.style, {
        position: 'fixed', left: `${m.x - 3}px`, top: `${m.y - 3}px`, width: `${m.w + 6}px`, height: `${m.h + 6}px`,
        border: '3px solid #DC2626', borderRadius: '6px', boxSizing: 'border-box', pointerEvents: 'none',
        zIndex: '99999', background: 'rgba(220,38,38,0.06)',
      });
      const badge = document.createElement('div');
      badge.className = '__help-marker';
      badge.textContent = String(m.n);
      Object.assign(badge.style, {
        position: 'fixed', left: `${Math.max(2, m.x - 16)}px`, top: `${Math.max(2, m.y - 16)}px`, width: '28px', height: '28px',
        borderRadius: '50%', background: '#DC2626', color: '#fff', fontWeight: '800', fontSize: '15px', lineHeight: '28px',
        textAlign: 'center', fontFamily: 'Arial, sans-serif', boxShadow: '0 1px 4px rgba(0,0,0,0.4)', zIndex: '100000', pointerEvents: 'none',
      });
      const tag = document.createElement('div');
      tag.className = '__help-marker';
      tag.textContent = m.label;
      Object.assign(tag.style, {
        position: 'fixed', left: `${Math.max(2, m.x + 16)}px`, top: `${Math.max(2, m.y - 14)}px`, background: '#DC2626', color: '#fff',
        fontSize: '12px', fontWeight: '700', padding: '2px 7px', borderRadius: '4px', whiteSpace: 'nowrap', fontFamily: 'Arial, sans-serif',
        boxShadow: '0 1px 3px rgba(0,0,0,0.35)', zIndex: '100000', pointerEvents: 'none',
      });
      document.documentElement.appendChild(box);
      document.documentElement.appendChild(badge);
      document.documentElement.appendChild(tag);
    }
  }, items);
}

async function openDialog(page, open) {
  if (open === 'importReview') {
    const quote = path.join(os.tmpdir(), 'gg-help-quote.pdf');
    fs.writeFileSync(quote, QUOTE_PDF);
    await page.locator('input[type=file]').first().setInputFiles(quote);
    await page.getByText('Supplier Quote Importer').first().waitFor({ timeout: 8000 });
  } else if (open === 'unlockDialog') {
    await page.getByRole('button', { name: 'Unlock to amend' }).click();
    await page.getByPlaceholder('Admin password').waitFor({ timeout: 8000 });
  } else if (open === 'amendMode') {
    await page.getByRole('button', { name: 'Unlock to amend' }).click();
    await page.getByPlaceholder('Admin password').fill('GGAdmin2026');
    await page.getByRole('button', { name: 'Unlock', exact: true }).click();
    await page.getByRole('button', { name: /Save as Rev/ }).waitFor({ timeout: 8000 });
  } else if (open === 'selectAllJobs') {
    await page.getByLabel('Select all jobs').check();
  }
}

(async () => {
  const browser = await chromium.launch({ args: ['--no-sandbox'] });
  for (const shot of SHOTS) {
    const ctx = await browser.newContext({ viewport: { width: 1180, height: shot.height || 900 }, deviceScaleFactor: 2 });
    const page = await ctx.newPage();
    page.on('pageerror', (e) => console.log('  page error:', e.message));
    await seed(page, Boolean(shot.issued));
    await page.locator('aside button', { hasText: SIDEBAR[shot.page] }).first().click();
    await page.waitForTimeout(350);
    if (shot.zoom) await page.evaluate((z) => { document.body.style.zoom = String(z); }, shot.zoom);
    if (shot.open) await openDialog(page, shot.open);
    await page.waitForTimeout(300);
    // Bring the first off-screen control into view once, then measure every marker at that same scroll position
    await page.evaluate(() => window.scrollTo(0, 0));
    const found = [];
    for (const c of shot.callouts) {
      const loc = await locate(page, c.find);
      if (!loc) { console.log(`  ! ${shot.file}: marker ${c.n} "${c.find}" not found`); continue; }
      found.push({ c, loc, r: await rectOf(await loc.elementHandle()) });
    }
    const offscreen = found.find((f) => f.r.y < 0 || f.r.y + f.r.h > (shot.height || 900));
    if (offscreen) {
      await offscreen.loc.evaluate((el) => el.scrollIntoView({ block: 'center' }));
      await page.waitForTimeout(200);
    }
    const items = [];
    for (const f of found) {
      const r = await rectOf(await f.loc.elementHandle());
      if (r.y < 0 || r.y + r.h > (shot.height || 900) || r.x < 0 || r.x + r.w > 1180) {
        console.log(`  ! ${shot.file}: marker ${f.c.n} "${f.c.find}" still outside the screen ${JSON.stringify(r)}`);
        continue;
      }
      items.push({ n: f.c.n, label: f.c.label, ...r });
    }
    await drawMarkers(page, items);
    await page.waitForTimeout(150);
    await page.screenshot({ path: path.join(OUT, shot.file), type: 'jpeg', quality: 78 });
    console.log(`saved ${shot.file} (${items.length}/${shot.callouts.length} markers)`);
    await ctx.close();
  }
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
