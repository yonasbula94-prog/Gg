import {
  CustomColumnDef,
  DashboardPreferences,
  DocPrefix,
  DocumentFontId,
  MarkupRate,
  ThemeCustomization,
  ThemePalette,
  ThemeVariantId,
} from '../types/procurement';

export const DEFAULT_DASHBOARD_PREFS: DashboardPreferences = {
  showRfqTenderCounters: true,
  showCompletionFunnel: true,
  showPerformanceChart: true,
  showPriorityRanking: true,
  showJobMatrix: true,
  showFinancials: true,
  includeCompleted: true,
  performancePeriod: 'weekly',
  priorityMetric: 'profit',
};

export const COMPANY_INFO = {
  name: 'GRAND GROUP FZC',
  tagline: 'International Procurement & Supply Chain Solutions',
  addressLine1: 'Executive Suite Z3 #81 | P.O. Box 121482',
  addressLine2: 'SAIF Zone, Sharjah, United Arab Emirates',
  phone: '+971 6 526 0000',
  email: 'sales@grandgroupfzc.com',
  trn: 'TRN 100428910200003',
  bankDetails: 'Bank: Emirates NBD · IBAN: AE420260001015428901201 · SWIFT: EBILAEAD',
};

// Confirmed 9-name End-User list (with Pharmecore fixed from Pharmaco)
export const END_USER_OPTIONS: string[] = [
  'AIAA (Asmara International Airport)',
  'Massawa Port',
  'MOT (Ministry of Transport)',
  'MOA (Ministry of Agriculture)',
  'MOE (Ministry of Education)',
  'Pharmecore',
  'ASEL',
  'MOEM (Ministry of Energy and Mines)',
  'OTHER',
];

export const DEFAULT_END_USER_ADDRESSES: Record<string, string> = {
  'AIAA (Asmara International Airport)': 'Asmara International Airport, Asmara, Eritrea',
  'Massawa Port': 'Massawa Port Authority, Massawa, Eritrea',
  'MOT (Ministry of Transport)': 'Ministry of Transport & Communications, Asmara, Eritrea',
  'MOA (Ministry of Agriculture)': 'Ministry of Agriculture, Asmara, Eritrea',
  'MOE (Ministry of Education)': 'Ministry of Education, Asmara, Eritrea',
  'Pharmecore': 'Pharmecore Medical Supplies, Asmara, Eritrea',
  'ASEL': 'ASEL Operations Office, Asmara, Eritrea',
  'MOEM (Ministry of Energy and Mines)': 'Ministry of Energy and Mines, Asmara, Eritrea',
};

export const CURRENCY_OPTIONS = ['USD', 'AED', 'EUR', 'GBP'];
export const INCOTERM_OPTIONS = ['CIF', 'CFR', 'FOB', 'EXW', 'DDP', 'DAP', 'FCA'];
export const UOM_OPTIONS = ['PCS', 'SET', 'UNIT', 'BOX', 'KG', 'LTR', 'MTR', 'ROLL', 'PAIR', 'PACK', 'LOT'];
export const MARKUP_RATES: MarkupRate[] = [0.1, 0.15, 0.18, 0.2];

export const DOCUMENT_FONT_OPTIONS: {
  id: DocumentFontId;
  label: string;
  cssStack: string;
}[] = [
  { id: 'Calibri', label: 'Calibri (Classic Excel)', cssStack: "'Calibri', 'Inter', system-ui, sans-serif" },
  { id: 'Arial', label: 'Arial (Clean Modern)', cssStack: "'Arial', 'Helvetica Neue', Helvetica, sans-serif" },
  { id: 'Segoe UI', label: 'Segoe UI (Windows Crisp)', cssStack: "'Segoe UI', system-ui, -apple-system, sans-serif" },
  { id: 'Georgia', label: 'Georgia (Executive Serif)', cssStack: "'Georgia', 'Times New Roman', serif" },
  { id: 'Times New Roman', label: 'Times New Roman (Formal)', cssStack: "'Times New Roman', Times, serif" },
  { id: 'Verdana', label: 'Verdana (Wide Readable)', cssStack: "'Verdana', Geneva, sans-serif" },
];

export const DEFAULT_PRICING_COLUMN_ORDER: string[] = [
  'requested',
  'offered',
  'uom',
  'qty',
  'unit_val',
  'total_val',
];

export const DEFAULT_LOGISTICS_COLUMN_ORDER: string[] = [
  'requested',
  'offered',
  'uom',
  'qty',
  'carton',
  'net_wt',
  'gross_wt',
];

export const DOC_META: Record<
  DocPrefix,
  {
    prefix: DocPrefix;
    title: string;
    sheetName: string;
    shortName: string;
    stageLabel: string;
    recipientLabel: string;
    isPricingDoc: boolean;
  }
> = {
  QT: {
    prefix: 'QT',
    title: 'QUOTATION',
    sheetName: 'QUOTATION',
    shortName: 'Quotation',
    stageLabel: 'Master Entry · Customer Quote',
    recipientLabel: 'TO',
    isPricingDoc: true,
  },
  PI: {
    prefix: 'PI',
    title: 'PROFORMA INVOICE',
    sheetName: 'PROFORMA INVOICE',
    shortName: 'Proforma Invoice',
    stageLabel: 'Triggered by Quotation',
    recipientLabel: 'TO',
    isPricingDoc: true,
  },
  PO: {
    prefix: 'PO',
    title: 'PURCHASE ORDER',
    sheetName: 'SUPPLIER PO',
    shortName: 'Supplier PO',
    stageLabel: 'Our PO to Supplier',
    recipientLabel: 'TO (SUPPLIER)',
    isPricingDoc: false,
  },
  PL: {
    prefix: 'PL',
    title: 'PACKING LIST',
    sheetName: 'PACKING LIST',
    shortName: 'Packing List',
    stageLabel: 'Shipment',
    recipientLabel: 'CONSIGNEE',
    isPricingDoc: false,
  },
  DN: {
    prefix: 'DN',
    title: 'DELIVERY NOTE',
    sheetName: 'DELIVERY NOTE',
    shortName: 'Delivery Note',
    stageLabel: 'Delivery',
    recipientLabel: 'DELIVER TO',
    isPricingDoc: false,
  },
  CI: {
    prefix: 'CI',
    title: 'COMMERCIAL INVOICE',
    sheetName: 'COMMERCIAL INVOICE',
    shortName: 'Commercial Invoice',
    stageLabel: 'Payment Settlement',
    recipientLabel: 'BILL TO',
    isPricingDoc: true,
  },
};

export const EXCEL_THEMES: Record<ThemeVariantId, ThemePalette> = {
  v1: {
    id: 'v1',
    name: 'v1 · Executive Navy & Slate',
    shortLabel: 'v1 Navy',
    subtitle: 'Classic institutional navy headers, crisp slate rules, ice-blue alternating rows',
    fontName: 'Calibri',
    fontCss: "'Calibri', 'Inter', system-ui, sans-serif",
    primary: '#1B365D',
    primaryText: '#FFFFFF',
    bodyText: '#0F172A',
    secondaryBg: '#EFF4FA',
    accent: '#2B5B84',
    zebraRow: '#F8FAFC',
    totalBg: '#E6EEF8',
    border: '#CBD5E1',
    strongBorder: '#1B365D',
    mutedText: '#475569',
    excel: {
      headerFill: 'FF1B365D',
      headerFont: 'FFFFFFFF',
      subHeaderFill: 'FFEFF4FA',
      accentFill: 'FF2B5B84',
      accentFont: 'FFFFFFFF',
      zebraFill: 'FFF8FAFC',
      totalFill: 'FFE6EEF8',
      borderColor: 'FFCBD5E1',
      strongBorderColor: 'FF1B365D',
      inputFont: 'FF1D4ED8',
      formulaFont: 'FF0F172A',
    },
  },
  v3: {
    id: 'v3',
    name: 'v3 · Charcoal & Burnished Amber',
    shortLabel: 'v3 Charcoal',
    subtitle: 'Modern graphite monochrome typography with sharp amber executive accent line',
    fontName: 'Calibri',
    fontCss: "'Calibri', 'Inter', system-ui, sans-serif",
    primary: '#1E293B',
    primaryText: '#FFFFFF',
    bodyText: '#18181B',
    secondaryBg: '#F4F5F7',
    accent: '#B45309',
    zebraRow: '#FAFAFA',
    totalBg: '#FEF3C7',
    border: '#D4D4D8',
    strongBorder: '#1E293B',
    mutedText: '#52525B',
    excel: {
      headerFill: 'FF1E293B',
      headerFont: 'FFFFFFFF',
      subHeaderFill: 'FFF4F5F7',
      accentFill: 'FFB45309',
      accentFont: 'FFFFFFFF',
      zebraFill: 'FFFAFAFA',
      totalFill: 'FFFEF3C7',
      borderColor: 'FFD4D4D8',
      strongBorderColor: 'FF1E293B',
      inputFont: 'FF1D4ED8',
      formulaFont: 'FF18181B',
    },
  },
};

function isValidHexColor(hex: string | undefined): hex is string {
  return typeof hex === 'string' && /^#[0-9A-Fa-f]{6}$/.test(hex.trim());
}

function hexToArgb(hex: string): string {
  return `FF${hex.replace('#', '').toUpperCase()}`;
}

function getContrastTextHex(hex: string): string {
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance > 0.62 ? '#0F172A' : '#FFFFFF';
}

export function resolveThemePalette(
  themeId: ThemeVariantId,
  custom?: ThemeCustomization | null
): ThemePalette {
  const base = EXCEL_THEMES[themeId] || EXCEL_THEMES.v1;
  if (!custom) return base;

  const primary = isValidHexColor(custom.primary) ? custom.primary.toUpperCase() : base.primary;
  const accent = isValidHexColor(custom.accent) ? custom.accent.toUpperCase() : base.accent;
  const totalBg = isValidHexColor(custom.totalBg) ? custom.totalBg.toUpperCase() : base.totalBg;
  const secondaryBg = isValidHexColor(custom.secondaryBg)
    ? custom.secondaryBg.toUpperCase()
    : base.secondaryBg;

  const fontMeta =
    DOCUMENT_FONT_OPTIONS.find((f) => f.id === custom.fontFamily) || DOCUMENT_FONT_OPTIONS[0];

  const primaryText = isValidHexColor(custom.headerText)
    ? custom.headerText.toUpperCase()
    : getContrastTextHex(primary);
  const bodyText = isValidHexColor(custom.bodyText)
    ? custom.bodyText.toUpperCase()
    : base.bodyText;
  const accentText = isValidHexColor(custom.headerText)
    ? custom.headerText.toUpperCase()
    : getContrastTextHex(accent);

  return {
    ...base,
    fontName: fontMeta.id,
    fontCss: fontMeta.cssStack,
    primary,
    primaryText,
    bodyText,
    secondaryBg,
    accent,
    totalBg,
    strongBorder: primary,
    excel: {
      ...base.excel,
      headerFill: hexToArgb(primary),
      headerFont: hexToArgb(primaryText),
      subHeaderFill: hexToArgb(secondaryBg),
      accentFill: hexToArgb(accent),
      accentFont: hexToArgb(accentText),
      totalFill: hexToArgb(totalBg),
      strongBorderColor: hexToArgb(primary),
      formulaFont: hexToArgb(bodyText),
    },
  };
}

/**
 * Resolves the unified ordered list of movable columns (excluding fixed 'no' column)
 * for any document tab so Single-Entry Schedule, A4 Print Preview, and Excel exports stay 100% synced.
 */
export function resolveDocumentColumns(
  columnOrder: string[] | undefined,
  customColumns: CustomColumnDef[] | undefined,
  isLogistics: boolean
): string[] {
  const cols = customColumns || [];
  const allowedBase = isLogistics
    ? DEFAULT_LOGISTICS_COLUMN_ORDER
    : DEFAULT_PRICING_COLUMN_ORDER;
  const allowedCustom = cols.map((c) => `custom:${c.id}`);
  const allowedSet = new Set([...allowedBase, ...allowedCustom]);

  // Build baseline order respecting legacy customColumn positions when columnOrder is absent
  const buildBaseline = (): string[] => {
    const at = (pos: string) =>
      cols.filter((c) => (c.position || 'after_no') === pos).map((c) => `custom:${c.id}`);
    return [
      ...at('after_no'),
      'requested',
      ...at('after_requested'),
      'offered',
      ...at('after_offered'),
      'uom',
      'qty',
      ...at('after_qty'),
      ...(isLogistics ? ['carton', 'net_wt', 'gross_wt'] : ['unit_val', 'total_val']),
      ...at('end'),
    ];
  };

  if (!Array.isArray(columnOrder) || columnOrder.length === 0) {
    return buildBaseline();
  }

  // Start from saved order, keeping only keys relevant to this document mode
  const result: string[] = [];
  const seen = new Set<string>();
  for (const key of columnOrder) {
    if (allowedSet.has(key) && !seen.has(key)) {
      result.push(key);
      seen.add(key);
    }
  }

  // Splice in any missing base columns in their default relative positions
  for (const baseKey of allowedBase) {
    if (!seen.has(baseKey)) {
      result.push(baseKey);
      seen.add(baseKey);
    }
  }

  // Splice in any newly added custom columns not yet in columnOrder at their chosen anchor
  for (const cc of cols) {
    const cKey = `custom:${cc.id}`;
    if (seen.has(cKey)) continue;
    const pos = cc.position || 'after_no';
    if (pos === 'after_no') {
      result.unshift(cKey);
    } else if (pos === 'after_requested') {
      const idx = result.indexOf('requested');
      result.splice(idx >= 0 ? idx + 1 : 0, 0, cKey);
    } else if (pos === 'after_offered') {
      const idx = result.indexOf('offered');
      result.splice(idx >= 0 ? idx + 1 : result.length, 0, cKey);
    } else if (pos === 'after_qty') {
      const idx = result.indexOf('qty');
      result.splice(idx >= 0 ? idx + 1 : result.length, 0, cKey);
    } else {
      result.push(cKey);
    }
    seen.add(cKey);
  }

  return result;
}

function mixHex(a: string, b: string, t: number): string {
  const parse = (h: string) => {
    const c = h.replace('#', '');
    return [0, 2, 4].map((i) => parseInt(c.slice(i, i + 2), 16));
  };
  const [ar, ag, ab] = parse(a);
  const [br, bg, bb] = parse(b);
  const mix = (x: number, y: number) => Math.round(x + (y - x) * t);
  return `#${[mix(ar, br), mix(ag, bg), mix(ab, bb)].map((v) => v.toString(16).padStart(2, '0')).join('')}`;
}

/**
 * Top-bar colour for each page. Derived only from the active theme's own
 * primary and accent colours, so every page stays inside the theme palette.
 */
export function pageBarColor(tab: string, palette: { primary: string; accent: string }): string {
  switch (tab) {
    case 'PRICE_CALC':
      return palette.accent;
    case 'INTERNAL_QUOTE':
      return mixHex(palette.primary, palette.accent, 0.4);
    case 'REGISTRY':
    case 'FILES':
    case 'DASHBOARD':
      return mixHex(palette.primary, '#000000', 0.25);
    default:
      return palette.primary;
  }
}
