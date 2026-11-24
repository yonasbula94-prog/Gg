import {
  CompanyProfile,
  DashboardPreferences,
  ProcurementJob,
  RegistryEntry,
  ThemePalette,
} from '../types/procurement';
import {
  calculateJob,
  formatMoney,
  getJobCompletionMetrics,
  getResolvedEndUser,
} from './calculations';
import { shareOrDownloadBlob } from './platform';

export type DashboardPdfOrientation = 'landscape' | 'portrait';

// A4 page sizes in PDF points (1 pt = 1/72 in)
const PAGE_SIZES: Record<DashboardPdfOrientation, { w: number; h: number }> = {
  landscape: { w: 842, h: 595 },
  portrait: { w: 595, h: 842 },
};
// Generous print-safe outer margin (42pt) for both orientations
const MARGIN_PT = 42;

function hexToRgb(hex: string): [number, number, number] {
  const clean = (hex || '#1B365D').replace('#', '').trim();
  const full =
    clean.length === 3
      ? clean
          .split('')
          .map((c) => c + c)
          .join('')
      : clean.padEnd(6, '0').slice(0, 6);
  const r = parseInt(full.slice(0, 2), 16);
  const g = parseInt(full.slice(2, 4), 16);
  const b = parseInt(full.slice(4, 6), 16);
  return [
    Number.isFinite(r) ? +(r / 255).toFixed(4) : 0.1,
    Number.isFinite(g) ? +(g / 255).toFixed(4) : 0.2,
    Number.isFinite(b) ? +(b / 255).toFixed(4) : 0.36,
  ];
}

function pdfSafeText(input: string): string {
  return (input || '')
    .replace(/\u2014|\u2013/g, '-')
    .replace(/\u2022|\u00b7/g, '|')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[^\x20-\x7E]/g, ' ')
    .replace(/\\/g, '\\\\')
    .replace(/\(/g, '\\(')
    .replace(/\)/g, '\\)');
}

function truncateText(text: string, maxChars: number): string {
  const clean = (text || '').trim();
  if (clean.length <= maxChars) return clean;
  return `${clean.slice(0, Math.max(1, maxChars - 3))}...`;
}

function approxTextWidth(text: string, fontSize: number, bold = false): number {
  return text.length * fontSize * (bold ? 0.56 : 0.52);
}

class PdfDocBuilder {
  private pages: string[][] = [];
  private current: string[] = [];
  private pw: number;
  private ph: number;
  private mg = MARGIN_PT;

  constructor(orientation: DashboardPdfOrientation = 'landscape') {
    const size = PAGE_SIZES[orientation];
    this.pw = size.w;
    this.ph = size.h;
    this.newPage();
  }

  newPage(): void {
    this.current = [];
    this.pages.push(this.current);
  }

  pageCount(): number {
    return this.pages.length;
  }

  private y(topY: number): number {
    return +(this.ph - topY).toFixed(2);
  }

  rect(
    x: number,
    topY: number,
    w: number,
    h: number,
    fillHex?: string,
    strokeHex?: string,
    lineWidth = 0.75
  ): void {
    const ops: string[] = ['q'];
    if (fillHex) {
      const [r, g, b] = hexToRgb(fillHex);
      ops.push(`${r} ${g} ${b} rg`);
    }
    if (strokeHex) {
      const [r, g, b] = hexToRgb(strokeHex);
      ops.push(`${r} ${g} ${b} RG`);
      ops.push(`${lineWidth} w`);
    }
    const bottomY = +(this.ph - topY - h).toFixed(2);
    ops.push(`${x.toFixed(2)} ${bottomY} ${w.toFixed(2)} ${h.toFixed(2)} re`);
    if (fillHex && strokeHex) ops.push('B');
    else if (fillHex) ops.push('f');
    else ops.push('S');
    ops.push('Q');
    this.current.push(ops.join(' '));
  }

  line(x1: number, topY1: number, x2: number, topY2: number, strokeHex = '#E2E8F0', lineWidth = 0.75): void {
    const [r, g, b] = hexToRgb(strokeHex);
    this.current.push(
      `q ${r} ${g} ${b} RG ${lineWidth} w ${x1.toFixed(2)} ${this.y(topY1)} m ${x2.toFixed(2)} ${this.y(topY2)} l S Q`
    );
  }

  text(
    str: string,
    x: number,
    topY: number,
    opts?: {
      size?: number;
      bold?: boolean;
      color?: string;
      align?: 'left' | 'right' | 'center';
      maxWidth?: number;
    }
  ): void {
    const size = opts?.size ?? 9;
    const bold = opts?.bold ?? false;
    const color = opts?.color ?? '#0F172A';
    const align = opts?.align ?? 'left';
    let out = str || '';
    if (opts?.maxWidth) {
      const maxChars = Math.max(4, Math.floor(opts.maxWidth / (size * (bold ? 0.56 : 0.52))));
      out = truncateText(out, maxChars);
    }
    const safe = pdfSafeText(out);
    const estW = approxTextWidth(out, size, bold);
    let drawX = x;
    if (align === 'right') drawX = x - estW;
    else if (align === 'center') drawX = x - estW / 2;

    const [r, g, b] = hexToRgb(color);
    const fontKey = bold ? '/F2' : '/F1';
    const baselineY = +(this.ph - topY - size * 0.82).toFixed(2);
    this.current.push(
      `BT ${fontKey} ${size} Tf ${r} ${g} ${b} rg 1 0 0 1 ${drawX.toFixed(2)} ${baselineY} Tm (${safe}) Tj ET`
    );
  }

  addPageFooters(companyName: string, generatedAt: string): void {
    const total = this.pages.length;
    this.pages.forEach((stream, idx) => {
      const saved = this.current;
      this.current = stream;
      this.line(this.mg, this.ph - 26, this.pw - this.mg, this.ph - 26, '#CBD5E1', 0.6);
      this.text(`${companyName} | Procurement Dashboard Report | Generated ${generatedAt}`, this.mg, this.ph - 21, {
        size: 7.5,
        color: '#64748B',
      });
      this.text(`Page ${idx + 1} of ${total}`, this.pw - this.mg, this.ph - 21, {
        size: 7.5,
        bold: true,
        color: '#475569',
        align: 'right',
      });
      this.current = saved;
    });
  }

  buildBytes(): Uint8Array {
    const encoder = new TextEncoder();
    const objects: string[] = [];

    // 1: Catalog, 2: Pages, 3: Font Helvetica, 4: Font Helvetica-Bold
    const pageObjNums: number[] = [];
    const contentObjNums: number[] = [];

    let nextObj = 5;
    for (let i = 0; i < this.pages.length; i++) {
      pageObjNums.push(nextObj++);
      contentObjNums.push(nextObj++);
    }

    objects.push(`1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n`);
    objects.push(
      `2 0 obj\n<< /Type /Pages /Kids [${pageObjNums.map((n) => `${n} 0 R`).join(' ')}] /Count ${pageObjNums.length} >>\nendobj\n`
    );
    objects.push(`3 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n`);
    objects.push(`4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>\nendobj\n`);

    for (let i = 0; i < this.pages.length; i++) {
      const pageNum = pageObjNums[i];
      const contentNum = contentObjNums[i];
      const streamBody = this.pages[i].join('\n');
      const streamBytes = encoder.encode(streamBody).length;
      objects.push(
        `${pageNum} 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${this.pw} ${this.ph}] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents ${contentNum} 0 R >>\nendobj\n`
      );
      objects.push(
        `${contentNum} 0 obj\n<< /Length ${streamBytes} >>\nstream\n${streamBody}\nendstream\nendobj\n`
      );
    }

    const chunks: Uint8Array[] = [];
    const offsets: number[] = [0];
    let cursor = 0;

    const pushStr = (s: string) => {
      const b = encoder.encode(s);
      chunks.push(b);
      cursor += b.length;
    };

    pushStr('%PDF-1.4\n');
    for (const obj of objects) {
      offsets.push(cursor);
      pushStr(obj);
    }

    const xrefPos = cursor;
    pushStr(`xref\n0 ${objects.length + 1}\n`);
    pushStr('0000000000 65535 f \n');
    for (let i = 1; i <= objects.length; i++) {
      pushStr(`${String(offsets[i]).padStart(10, '0')} 00000 n \n`);
    }
    pushStr(
      `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefPos}\n%%EOF\n`
    );

    const out = new Uint8Array(cursor);
    let pos = 0;
    for (const ch of chunks) {
      out.set(ch, pos);
      pos += ch.length;
    }
    return out;
  }
}

function getIsoWeekKey(date: Date): string {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil(((d.getTime() - yearStart.getTime()) / 86400000 + 1) / 7);
  return `${d.getUTCFullYear()}-W${String(weekNo).padStart(2, '0')}`;
}

function getMonthKey(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

/**
 * Generates and downloads a styled, multi-page A4 PDF report (landscape or portrait)
 * mirroring the user's active Dashboard configuration and theme colors.
 */
export async function exportDashboardReportPdf(
  jobs: ProcurementJob[],
  registry: RegistryEntry[],
  theme: ThemePalette,
  company: CompanyProfile,
  prefs: DashboardPreferences,
  orientation: DashboardPdfOrientation = 'landscape'
): Promise<string> {
  const todayStr = new Date().toISOString().slice(0, 10);
  const { w: PAGE_W, h: PAGE_H } = PAGE_SIZES[orientation];
  const MARGIN = MARGIN_PT;
  const CONTENT_W = PAGE_W - MARGIN * 2;
  const pdf = new PdfDocBuilder(orientation);

  const enrichedAll = jobs.map((job) => {
    const calc = calculateJob(job);
    const metrics = getJobCompletionMetrics(job);
    const client = getResolvedEndUser(job) || 'Unassigned Client';
    const marginPct =
      calc.goodsSubtotalExVat > 0
        ? Math.round((calc.grossProfit / calc.goodsSubtotalExVat) * 1000) / 10
        : 0;
    return { job, calc, metrics, client, marginPct };
  });

  const scoped = prefs.includeCompleted
    ? enrichedAll
    : enrichedAll.filter((e) => !e.metrics.isComplete);

  const rfqJobs = scoped.filter((e) => e.job.inquiryType !== 'TENDER');
  const tenderJobs = scoped.filter((e) => e.job.inquiryType === 'TENDER');
  const totalCount = scoped.length;
  const completedCount = enrichedAll.filter((e) => e.metrics.isComplete).length;
  const avgCompletionPct =
    totalCount > 0
      ? Math.round(scoped.reduce((s, e) => s + e.metrics.percent, 0) / totalCount)
      : 0;
  const rfqValue = rfqJobs.reduce((s, e) => s + e.calc.grandTotal, 0);
  const tenderValue = tenderJobs.reduce((s, e) => s + e.calc.grandTotal, 0);
  const totalPipelineValue = rfqValue + tenderValue;
  const totalEstProfit = scoped.reduce((s, e) => s + e.calc.grossProfit, 0);
  const primaryCurr = scoped[0]?.job.controls.currency || 'USD';

  const drawTopBanner = (subtitle: string): number => {
    pdf.rect(MARGIN, 24, CONTENT_W, 46, theme.primary);
    pdf.text(company.name, MARGIN + 14, 34, {
      size: 13,
      bold: true,
      color: theme.primaryText || '#FFFFFF',
    });
    pdf.text(
      `${company.addressLine1}${company.addressLine2 ? ` | ${company.addressLine2}` : ''}`,
      MARGIN + 14,
      51,
      { size: 7.5, color: '#CBD5E1', maxWidth: 420 }
    );
    pdf.text('PROCUREMENT DASHBOARD REPORT', PAGE_W - MARGIN - 14, 34, {
      size: 11,
      bold: true,
      color: theme.primaryText || '#FFFFFF',
      align: 'right',
    });
    pdf.text(subtitle, PAGE_W - MARGIN - 14, 51, {
      size: 8,
      color: '#E2E8F0',
      align: 'right',
    });
    return 82;
  };

  let curY = drawTopBanner(
    `Date: ${todayStr} | Scope: ${prefs.includeCompleted ? 'All Jobs' : 'Active Only'} | ${
      prefs.performancePeriod === 'weekly' ? 'Weekly (8W)' : 'Monthly (6M)'
    }`
  );

  // 1. KPI Summary Strip (Respects showRfqTenderCounters & showFinancials)
  if (prefs.showRfqTenderCounters) {
    const cards: { label: string; value: string; sub: string; accent?: string }[] = [
      {
        label: 'ACTIVE RFQS',
        value: String(rfqJobs.length),
        sub: prefs.showFinancials
          ? `${totalCount > 0 ? Math.round((rfqJobs.length / totalCount) * 100) : 0}% of jobs | ${formatMoney(rfqValue, primaryCurr)}`
          : `${totalCount > 0 ? Math.round((rfqJobs.length / totalCount) * 100) : 0}% of jobs`,
      },
      {
        label: 'ACTIVE TENDERS',
        value: String(tenderJobs.length),
        sub: prefs.showFinancials
          ? `${totalCount > 0 ? Math.round((tenderJobs.length / totalCount) * 100) : 0}% of jobs | ${formatMoney(tenderValue, primaryCurr)}`
          : `${totalCount > 0 ? Math.round((tenderJobs.length / totalCount) * 100) : 0}% of jobs`,
      },
      {
        label: 'OVERALL COMPLETION',
        value: `${avgCompletionPct}%`,
        sub: `${completedCount} of ${enrichedAll.length} fully delivered`,
        accent: theme.primary,
      },
      ...(prefs.showFinancials
        ? [
            {
              label: 'PIPELINE VALUE & EST. PROFIT',
              value: formatMoney(totalPipelineValue, primaryCurr),
              sub: `Est. Profit: ${formatMoney(totalEstProfit, primaryCurr)}`,
              accent: '#047857',
            },
          ]
        : []),
    ];

    const gap = 10;
    const cardW = (CONTENT_W - gap * (cards.length - 1)) / cards.length;
    const cardH = 56;
    cards.forEach((c, idx) => {
      const cx = MARGIN + idx * (cardW + gap);
      pdf.rect(cx, curY, cardW, cardH, '#FFFFFF', '#CBD5E1', 0.7);
      pdf.text(c.label, cx + 10, curY + 9, { size: 7.5, bold: true, color: '#64748B' });
      pdf.text(c.value, cx + 10, curY + 23, {
        size: 14,
        bold: true,
        color: c.accent || '#0F172A',
        maxWidth: cardW - 20,
      });
      pdf.text(c.sub, cx + 10, curY + 42, {
        size: 7.5,
        color: '#475569',
        maxWidth: cardW - 20,
      });
    });
    curY += cardH + 14;
  }

  // 2. Two-Column Visual Summary: Stage Completion Funnel + Weekly/Monthly Performance
  const showFunnel = prefs.showCompletionFunnel;
  const showPerf = prefs.showPerformanceChart;
  if (showFunnel || showPerf) {
    const colGap = 12;
    const twoCol = showFunnel && showPerf;
    const leftW = twoCol ? (CONTENT_W - colGap) * 0.54 : CONTENT_W;
    const rightW = twoCol ? (CONTENT_W - colGap) * 0.46 : CONTENT_W;
    const panelH = 212;

    if (showFunnel) {
      pdf.rect(MARGIN, curY, leftW, panelH, '#FFFFFF', '#CBD5E1', 0.7);
      pdf.text('STAGE COMPLETION BREAKDOWN (%)', MARGIN + 12, curY + 11, {
        size: 8.5,
        bold: true,
        color: '#0F172A',
      });
      pdf.text(`${totalCount} Job${totalCount === 1 ? '' : 's'}`, MARGIN + leftW - 12, curY + 11, {
        size: 7.5,
        color: '#64748B',
        align: 'right',
      });

      const stages = [
        { label: 'Internal Supplier Quote', count: scoped.filter((e) => e.metrics.hasInternalQuote).length, color: theme.primary },
        { label: 'Quotation (QT) Issued', count: scoped.filter((e) => e.metrics.hasQT).length, color: theme.primary },
        { label: 'Proforma Invoice (PI) Issued', count: scoped.filter((e) => e.metrics.hasPI).length, color: theme.primary },
        { label: 'Client PO Confirmed', count: scoped.filter((e) => e.metrics.hasEndUserPoGate).length, color: '#D97706' },
        { label: 'Supplier PO Issued', count: scoped.filter((e) => e.metrics.hasPO).length, color: theme.primary },
        { label: 'Commercial Invoice (CI) Issued', count: scoped.filter((e) => e.metrics.hasCI).length, color: theme.primary },
        { label: 'Supplier Paid & PL Confirmed', count: scoped.filter((e) => e.metrics.hasSupplierShipmentGate).length, color: '#D97706' },
        { label: 'Packing List (PL) Issued', count: scoped.filter((e) => e.metrics.hasPL).length, color: theme.primary },
        { label: 'Delivery Note (DN) Issued', count: scoped.filter((e) => e.metrics.hasDN).length, color: '#059669' },
      ];

      const barTrackW = leftW - 24;
      stages.forEach((st, idx) => {
        const sy = curY + 29 + idx * 19.5;
        const pct = totalCount > 0 ? Math.round((st.count / totalCount) * 100) : 0;
        pdf.text(st.label, MARGIN + 12, sy, { size: 7.5, color: '#334155', maxWidth: leftW - 95 });
        pdf.text(`${st.count}/${totalCount} (${pct}%)`, MARGIN + leftW - 12, sy, {
          size: 7.5,
          bold: true,
          color: '#0F172A',
          align: 'right',
        });
        pdf.rect(MARGIN + 12, sy + 10.5, barTrackW, 4.5, '#E2E8F0');
        if (pct > 0) {
          pdf.rect(MARGIN + 12, sy + 10.5, Math.max(3, (barTrackW * pct) / 100), 4.5, st.color);
        }
      });
    }

    if (showPerf) {
      const rx = twoCol ? MARGIN + leftW + colGap : MARGIN;
      pdf.rect(rx, curY, rightW, panelH, '#FFFFFF', '#CBD5E1', 0.7);
      const perfLabel =
        prefs.performancePeriod === 'weekly'
          ? 'WEEKLY PERFORMANCE (LAST 8 WEEKS)'
          : 'MONTHLY PERFORMANCE (LAST 6 MONTHS)';
      pdf.text(perfLabel, rx + 12, curY + 11, { size: 8.5, bold: true, color: '#0F172A' });

      const now = new Date();
      const buckets: { key: string; label: string; jobsStarted: number; docsIssued: number }[] = [];
      if (prefs.performancePeriod === 'weekly') {
        for (let i = 7; i >= 0; i--) {
          const d = new Date(now);
          d.setDate(now.getDate() - i * 7);
          const key = getIsoWeekKey(d);
          buckets.push({ key, label: key.split('-')[1] || key, jobsStarted: 0, docsIssued: 0 });
        }
        const map = new Map(buckets.map((b) => [b.key, b]));
        jobs.forEach((j) => {
          const b = map.get(getIsoWeekKey(new Date(j.createdAt || j.updatedAt)));
          if (b) b.jobsStarted += 1;
        });
        registry.forEach((e) => {
          if (e.status === 'VOID') return;
          const b = map.get(getIsoWeekKey(new Date(e.issuedAt)));
          if (b) b.docsIssued += 1;
        });
      } else {
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        for (let i = 5; i >= 0; i--) {
          const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
          const key = getMonthKey(d);
          buckets.push({ key, label: months[d.getMonth()], jobsStarted: 0, docsIssued: 0 });
        }
        const map = new Map(buckets.map((b) => [b.key, b]));
        jobs.forEach((j) => {
          const b = map.get(getMonthKey(new Date(j.createdAt || j.updatedAt)));
          if (b) b.jobsStarted += 1;
        });
        registry.forEach((e) => {
          if (e.status === 'VOID') return;
          const b = map.get(getMonthKey(new Date(e.issuedAt)));
          if (b) b.docsIssued += 1;
        });
      }

      const totalDocsInRange = buckets.reduce((s, b) => s + b.docsIssued, 0);
      pdf.text(`Period Total: ${totalDocsInRange} docs`, rx + rightW - 12, curY + 11, {
        size: 7.5,
        bold: true,
        color: '#475569',
        align: 'right',
      });

      const chartX = rx + 26;
      const chartY = curY + 34;
      const chartW = rightW - 40;
      const chartH = 132;
      const maxVal = Math.max(1, ...buckets.map((b) => Math.max(b.jobsStarted, b.docsIssued)));

      // Horizontal grid lines
      for (let t = 0; t <= 4; t++) {
        const gy = chartY + chartH - (chartH * t) / 4;
        const val = Math.round((maxVal * t) / 4);
        pdf.line(chartX, gy, chartX + chartW, gy, '#E2E8F0', 0.5);
        pdf.text(String(val), chartX - 5, gy - 4, { size: 7, color: '#64748B', align: 'right' });
      }

      const slotW = chartW / Math.max(1, buckets.length);
      const barW = Math.min(14, slotW * 0.34);

      buckets.forEach((b, idx) => {
        const centerX = chartX + idx * slotW + slotW / 2;
        const jH = b.jobsStarted > 0 ? Math.max(3, (b.jobsStarted / maxVal) * chartH) : 1.5;
        const dH = b.docsIssued > 0 ? Math.max(3, (b.docsIssued / maxVal) * chartH) : 1.5;

        pdf.rect(centerX - barW - 1.5, chartY + chartH - jH, barW, jH, '#94A3B8');
        pdf.rect(centerX + 1.5, chartY + chartH - dH, barW, dH, theme.primary);

        if (b.docsIssued > 0) {
          pdf.text(String(b.docsIssued), centerX + 1.5 + barW / 2, chartY + chartH - dH - 10, {
            size: 7,
            bold: true,
            color: theme.primary,
            align: 'center',
          });
        }

        pdf.text(b.label, centerX, chartY + chartH + 6, {
          size: 7.5,
          color: '#475569',
          align: 'center',
        });
      });

      // Legend
      const legY = curY + panelH - 18;
      pdf.rect(rx + 14, legY + 1, 7, 7, '#94A3B8');
      pdf.text('Jobs Started', rx + 25, legY, { size: 7.5, color: '#475569' });
      pdf.rect(rx + 98, legY + 1, 7, 7, theme.primary);
      pdf.text('Documents Issued', rx + 109, legY, { size: 7.5, color: '#475569' });
    }

    curY += panelH + 14;
  }

  // 3. Priority Ranking Panel (Respects showPriorityRanking)
  if (prefs.showPriorityRanking && scoped.length > 0) {
    const ranked = [...scoped]
      .sort((a, b) =>
        prefs.priorityMetric === 'profit'
          ? b.calc.grossProfit - a.calc.grossProfit
          : b.calc.grandTotal - a.calc.grandTotal
      )
      .slice(0, 6);

    const neededH = 34 + ranked.length * 22;
    if (curY + neededH > PAGE_H - 40) {
      pdf.newPage();
      curY = drawTopBanner(`Priority Ranking & Lifecycle Tracker | Date: ${todayStr}`);
    }

    pdf.rect(MARGIN, curY, CONTENT_W, neededH, '#FFFFFF', '#CBD5E1', 0.7);
    const rankTitle =
      prefs.priorityMetric === 'profit'
        ? 'PRIORITY JOB RANKING — BY EST. PROFITABILITY'
        : 'PRIORITY JOB RANKING — BY OVERALL VALUE';
    pdf.text(rankTitle, MARGIN + 12, curY + 10, { size: 8.5, bold: true, color: '#0F172A' });

    const maxMetric = Math.max(
      1,
      ...ranked.map((r) => (prefs.priorityMetric === 'profit' ? Math.max(0, r.calc.grossProfit) : Math.max(0, r.calc.grandTotal)))
    );

    ranked.forEach((item, idx) => {
      const ry = curY + 28 + idx * 22;
      const metricVal = prefs.priorityMetric === 'profit' ? item.calc.grossProfit : item.calc.grandTotal;
      const barPct = Math.max(0, Math.min(100, (metricVal / maxMetric) * 100));

      pdf.text(`#${idx + 1}`, MARGIN + 12, ry, { size: 8, bold: true, color: '#64748B' });
      pdf.text(
        `${item.job.jobName}  |  ${item.client}  |  ${item.job.inquiryType === 'TENDER' ? 'TENDER' : 'RFQ'} ${item.job.controls.rfqRef || ''}`,
        MARGIN + 34,
        ry,
        { size: 8, bold: true, color: '#0F172A', maxWidth: 460 }
      );
      if (prefs.showFinancials) {
        pdf.text(
          `${formatMoney(metricVal, item.job.controls.currency)} (${item.marginPct}% margin | Val ${formatMoney(item.calc.grandTotal, item.job.controls.currency)})`,
          PAGE_W - MARGIN - 12,
          ry,
          { size: 7.5, bold: true, color: '#047857', align: 'right' }
        );
      } else {
        pdf.text(`${item.metrics.percent}% Complete`, PAGE_W - MARGIN - 12, ry, {
          size: 7.5,
          bold: true,
          color: '#475569',
          align: 'right',
        });
      }
      pdf.rect(MARGIN + 34, ry + 11, 360, 4, '#E2E8F0');
      if (barPct > 0) {
        pdf.rect(
          MARGIN + 34,
          ry + 11,
          Math.max(4, (360 * barPct) / 100),
          4,
          prefs.priorityMetric === 'profit' ? '#059669' : theme.primary
        );
      }
    });

    curY += neededH + 14;
  }

  // 4. Job Lifecycle Tracker Matrix Table (Multi-page with auto-repeated headers)
  if (prefs.showJobMatrix) {
    if (curY + 80 > PAGE_H - 40) {
      pdf.newPage();
      curY = drawTopBanner(`Job Lifecycle Tracker | Date: ${todayStr}`);
    }

    const cols: { label: string; w: number; align: 'left' | 'center' | 'right' }[] = [
      { label: 'Seq', w: 44, align: 'center' },
      { label: 'Job & Client', w: 184, align: 'left' },
      { label: 'Type / RFQ-Tender Ref', w: 112, align: 'left' },
      { label: 'Stage', w: 88, align: 'left' },
      { label: 'Done', w: 42, align: 'right' },
      { label: 'IntQ', w: 38, align: 'center' },
      { label: 'QT', w: 38, align: 'center' },
      { label: 'PI', w: 38, align: 'center' },
      { label: 'PO', w: 38, align: 'center' },
      { label: 'CI', w: 38, align: 'center' },
      { label: 'PL', w: 38, align: 'center' },
      { label: 'DN', w: 38, align: 'center' },
      { label: prefs.showFinancials ? 'Value' : 'Curr', w: 42, align: 'right' },
    ];
    // Adjust exact width to match CONTENT_W (778)
    const sumW = cols.reduce((s, c) => s + c.w, 0);
    cols[1].w += CONTENT_W - sumW;

    const drawMatrixHeader = (yPos: number): number => {
      pdf.text('JOB LIFECYCLE TRACKER MATRIX', MARGIN, yPos, {
        size: 8.5,
        bold: true,
        color: '#0F172A',
      });
      const hy = yPos + 14;
      pdf.rect(MARGIN, hy, CONTENT_W, 20, theme.primary);
      let x = MARGIN;
      cols.forEach((col) => {
        const tx = col.align === 'center' ? x + col.w / 2 : col.align === 'right' ? x + col.w - 4 : x + 5;
        pdf.text(col.label, tx, hy + 6, {
          size: 7.5,
          bold: true,
          color: theme.primaryText || '#FFFFFF',
          align: col.align,
        });
        x += col.w;
      });
      return hy + 20;
    };

    curY = drawMatrixHeader(curY);

    if (scoped.length === 0) {
      pdf.rect(MARGIN, curY, CONTENT_W, 24, '#FFFFFF', '#E2E8F0', 0.6);
      pdf.text('No jobs match the active Dashboard scope.', MARGIN + CONTENT_W / 2, curY + 8, {
        size: 8,
        color: '#64748B',
        align: 'center',
      });
    } else {
      scoped.forEach((ej, idx) => {
        const rowH = 24;
        if (curY + rowH > PAGE_H - 36) {
          pdf.newPage();
          curY = drawTopBanner(`Job Lifecycle Tracker (Continued) | Date: ${todayStr}`);
          curY = drawMatrixHeader(curY);
        }

        const bg = idx % 2 === 1 ? '#F8FAFC' : '#FFFFFF';
        pdf.rect(MARGIN, curY, CONTENT_W, rowH, bg, '#E2E8F0', 0.5);

        const cells: { text: string; sub?: string; color?: string; bold?: boolean }[] = [
          {
            text: ej.job.jobSequence ? `#${String(ej.job.jobSequence).padStart(3, '0')}` : 'Draft',
            bold: true,
            color: '#0F172A',
          },
          { text: ej.job.jobName, sub: ej.client, bold: true, color: '#0F172A' },
          {
            text: `${ej.job.inquiryType === 'TENDER' ? 'TENDER' : 'RFQ'} ${ej.job.controls.rfqRef || '-'}`,
            color: '#334155',
          },
          { text: ej.metrics.currentStageLabel, color: '#334155' },
          { text: `${ej.metrics.percent}%`, bold: true, color: ej.metrics.isComplete ? '#047857' : '#0F172A' },
          { text: ej.metrics.hasInternalQuote ? 'Ready' : '-', color: ej.metrics.hasInternalQuote ? '#047857' : '#94A3B8' },
          { text: ej.metrics.hasQT ? 'OK' : '-', bold: ej.metrics.hasQT, color: ej.metrics.hasQT ? '#047857' : '#94A3B8' },
          { text: ej.metrics.hasPI ? 'OK' : '-', bold: ej.metrics.hasPI, color: ej.metrics.hasPI ? '#047857' : '#94A3B8' },
          {
            text: ej.metrics.hasPO ? 'OK' : ej.metrics.hasEndUserPoGate ? '-' : 'Pending',
            bold: ej.metrics.hasPO,
            color: ej.metrics.hasPO ? '#047857' : ej.metrics.hasEndUserPoGate ? '#64748B' : '#B45309',
          },
          {
            text: ej.metrics.hasCI ? 'OK' : ej.metrics.hasEndUserPoGate ? '-' : 'Pending',
            bold: ej.metrics.hasCI,
            color: ej.metrics.hasCI ? '#047857' : ej.metrics.hasEndUserPoGate ? '#64748B' : '#B45309',
          },
          {
            text: ej.metrics.hasPL ? 'OK' : ej.metrics.hasSupplierShipmentGate ? '-' : 'Pending',
            bold: ej.metrics.hasPL,
            color: ej.metrics.hasPL ? '#047857' : ej.metrics.hasSupplierShipmentGate ? '#64748B' : '#B45309',
          },
          {
            text: ej.metrics.hasDN ? 'OK' : ej.metrics.hasPL ? '-' : 'Pending',
            bold: ej.metrics.hasDN,
            color: ej.metrics.hasDN ? '#047857' : ej.metrics.hasPL ? '#64748B' : '#B45309',
          },
          {
            text: prefs.showFinancials
              ? Math.round(ej.calc.grandTotal).toLocaleString('en-US')
              : ej.job.controls.currency,
            bold: prefs.showFinancials,
            color: '#0F172A',
          },
        ];

        let x = MARGIN;
        cells.forEach((cell, cIdx) => {
          const col = cols[cIdx];
          const tx =
            col.align === 'center' ? x + col.w / 2 : col.align === 'right' ? x + col.w - 4 : x + 5;
          if (cell.sub) {
            pdf.text(cell.text, tx, curY + 4, {
              size: 7.5,
              bold: true,
              color: cell.color,
              maxWidth: col.w - 8,
            });
            pdf.text(cell.sub, tx, curY + 13.5, {
              size: 6.8,
              color: '#64748B',
              maxWidth: col.w - 8,
            });
          } else {
            pdf.text(cell.text, tx, curY + 8, {
              size: 7.5,
              bold: cell.bold,
              color: cell.color,
              align: col.align,
              maxWidth: col.w - 6,
            });
          }
          x += col.w;
        });

        curY += rowH;
      });
    }
  }

  pdf.addPageFooters(company.name, todayStr);
  const bytes = pdf.buildBytes();
  const filename = `GrandGroup_Dashboard_Report_${todayStr}.pdf`;
  const blob = new Blob([bytes.buffer as ArrayBuffer], { type: 'application/pdf' });
  await shareOrDownloadBlob(blob, filename, 'application/pdf');
  return filename;
}
