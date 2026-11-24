import { DocPrefix, MarkupRate, ThemePalette, ThemeVariantId } from '../types/procurement';

export const COMPANY_INFO = {
  name: 'GRAND GROUP FZC',
  tagline: 'International Procurement & Supply Chain Solutions',
  addressLine1: 'P.O. Box 42188, Hamriyah Free Zone',
  addressLine2: 'Sharjah, United Arab Emirates',
  phone: '+971 6 526 0000',
  email: 'procurement@grandgroupfzc.com',
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

export const CURRENCY_OPTIONS = ['USD', 'AED', 'EUR', 'GBP'];
export const INCOTERM_OPTIONS = ['CIF', 'CFR', 'FOB', 'EXW', 'DDP', 'DAP', 'FCA'];
export const UOM_OPTIONS = ['PCS', 'SET', 'UNIT', 'BOX', 'KG', 'LTR', 'MTR', 'ROLL', 'PAIR', 'PACK', 'LOT'];
export const MARKUP_RATES: MarkupRate[] = [0.1, 0.15, 0.18, 0.2];

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
    stageLabel: 'Customer Quote',
    recipientLabel: 'TO (END USER)',
    isPricingDoc: true,
  },
  PI: {
    prefix: 'PI',
    title: 'PROFORMA INVOICE',
    sheetName: 'PROFORMA INVOICE',
    shortName: 'Proforma Invoice',
    stageLabel: 'Customer PO / PI',
    recipientLabel: 'TO (END USER)',
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
    recipientLabel: 'CONSIGNEE / END USER',
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
    recipientLabel: 'BILL TO (END USER)',
    isPricingDoc: true,
  },
};

export const EXCEL_THEMES: Record<ThemeVariantId, ThemePalette> = {
  v1: {
    id: 'v1',
    name: 'v1 · Executive Navy & Slate',
    shortLabel: 'v1 Navy',
    subtitle: 'Classic institutional navy headers, crisp slate rules, ice-blue alternating rows',
    primary: '#1B365D',
    primaryText: '#FFFFFF',
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
  v2: {
    id: 'v2',
    name: 'v2 · Forest Emerald & Ivory',
    shortLabel: 'v2 Emerald',
    subtitle: 'Deep heritage emerald bands, warm stone summary blocks, refined sage dividers',
    primary: '#14532D',
    primaryText: '#FFFFFF',
    secondaryBg: '#F0F7F2',
    accent: '#166534',
    zebraRow: '#F9FBF9',
    totalBg: '#E4F1E8',
    border: '#CBD8CE',
    strongBorder: '#14532D',
    mutedText: '#47594D',
    excel: {
      headerFill: 'FF14532D',
      headerFont: 'FFFFFFFF',
      subHeaderFill: 'FFF0F7F2',
      accentFill: 'FF166534',
      accentFont: 'FFFFFFFF',
      zebraFill: 'FFF9FBF9',
      totalFill: 'FFE4F1E8',
      borderColor: 'FFCBD8CE',
      strongBorderColor: 'FF14532D',
      inputFont: 'FF1D4ED8',
      formulaFont: 'FF111827',
    },
  },
  v3: {
    id: 'v3',
    name: 'v3 · Charcoal & Burnished Amber',
    shortLabel: 'v3 Charcoal',
    subtitle: 'Modern graphite monochrome typography with sharp amber executive accent line',
    primary: '#1E293B',
    primaryText: '#FFFFFF',
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
