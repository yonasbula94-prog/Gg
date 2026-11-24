export type DocPrefix = 'QT' | 'PI' | 'PO' | 'PL' | 'DN' | 'CI';

export type MarkupRate = 0.1 | 0.15 | 0.18 | 0.2;

export type ThemeVariantId = 'v1' | 'v2' | 'v3';

export type PlatformMode = 'windows' | 'ipad';

export type WorkspaceTab =
  | 'PI'
  | 'PRICE_CALC'
  | 'QT'
  | 'PO'
  | 'PL'
  | 'DN'
  | 'CI'
  | 'REGISTRY'
  | 'STUDIO';

export interface CompanyProfile {
  name: string;
  tagline: string;
  addressLine1: string;
  addressLine2: string;
  phone: string;
  email: string;
  trn: string;
  bankDetails: string;
  logoDataUrl: string;
  endUsers: string[];
}

export interface ThemePalette {
  id: ThemeVariantId;
  name: string;
  shortLabel: string;
  subtitle: string;
  // Hex colors for UI / Print
  primary: string;
  primaryText: string;
  secondaryBg: string;
  accent: string;
  zebraRow: string;
  totalBg: string;
  border: string;
  strongBorder: string;
  mutedText: string;
  // ARGB (8-char) for ExcelJS
  excel: {
    headerFill: string;
    headerFont: string;
    subHeaderFill: string;
    accentFill: string;
    accentFont: string;
    zebraFill: string;
    totalFill: string;
    borderColor: string;
    strongBorderColor: string;
    inputFont: string;
    formulaFont: string;
  };
}

export type CustomColumnPosition =
  | 'after_no'
  | 'after_requested'
  | 'after_offered'
  | 'after_qty'
  | 'end';

export interface CustomColumnDef {
  id: string;
  label: string; // e.g. "Part No.", "CAS No.", "Brand / Origin", "HS Code"
  position?: CustomColumnPosition; // Defaults to 'after_no'
}

export interface JobItemRow {
  rowNumber: number; // Strictly 17..36 (20 fixed rows)
  requestedDescription: string;
  offeredDescription: string;
  uom: string;
  qty: number | null;
  // Optional per-job custom columns (Part No., CAS No., etc.)
  customFields?: Record<string, string>;
  // PRICE CALC inputs (internal, never sent to customer)
  supplierUnitCost: number | null;
  boxes: number | null;
  boxLengthCm: number | null;
  boxWidthCm: number | null;
  boxHeightCm: number | null;
  kgPerBox: number | null;
  // PACKING LIST & DELIVERY NOTE editable shipment fields
  cartonNo: string;
  netWtKg: number | null;
  grossWtKg: number | null;
}

export interface FreightConfig {
  ratePerCbm: number;
  ratePerKg: number;
  docChargePct: number; // Default 0.10 (10%)
  selectedMarkupPct: number; // E3 on PRICE CALC: any decimal (e.g., 0.10, 0.125, 0.25)
}

export interface JobControls {
  endUserSelect: string; // Dropdown from 9 confirmed end users
  customEndUser: string; // Used when endUserSelect === 'OTHER'
  endUserContact: string;
  endUserAddress: string;
  currency: string;
  incoterm: string;
  destination: string;
  vatRate: number; // PI!L5 default 0.05 (5%), 0 for zero-rated
  quoteValidityDays: number;
  paymentTerms: string;
  deliveryLeadTime: string;
  rfqRef: string;
  customerPoRef: string;
  // Supplier PO metadata
  supplierName: string;
  supplierAddress: string;
  supplierRef: string;
  // Shipment & Delivery metadata
  shipDate: string;
  portOfLoading: string;
  deliveryDate: string;
  transportRef: string;
  receivedBy: string;
  // Editable per-job footer remarks & signature blocks (leave empty to hide/leave blank)
  remarks: string;
  authorizedSignatory: string;
  counterpartySignatory: string;
}

export interface IssuedDocRecord {
  prefix: DocPrefix;
  docNumber: string; // Hard-coded on issue, e.g., PI-GG-043-2026
  sequence: number;
  issuedAt: string;
  issuedBy: string;
  isVoided?: boolean;
  // Frozen financial totals at issuance
  frozenSubtotal: number;
  frozenVat: number;
  frozenFreight: number;
  frozenGrandTotal: number;
}

export interface ProcurementJob {
  id: string;
  jobName: string;
  createdAt: string;
  updatedAt: string;
  preparedBy: string;
  controls: JobControls;
  freight: FreightConfig;
  customColumns?: CustomColumnDef[];
  items: JobItemRow[]; // Always length 20, rowNumber 17..36
  issuedDocs: Partial<Record<DocPrefix, IssuedDocRecord>>;
}

export interface RegistryEntry {
  sequence: number;
  docNumber: string; // PREFIX-GG-SEQ-YEAR
  prefix: DocPrefix;
  year: number;
  jobId: string;
  jobName: string;
  customer: string;
  currency: string;
  amount: number;
  issuedBy: string;
  issuedAt: string;
  status: 'ISSUED' | 'VOID';
  voidReason?: string;
}

export interface ValidationIssue {
  rowNumber?: number;
  field: string;
  message: string;
}

export interface ComputedRow {
  rowNumber: number;
  index: number; // 1..20
  isEmpty: boolean;
  requestedDescription: string;
  offeredDescription: string;
  customFields: Record<string, string>;
  uom: string;
  qty: number | null;
  supplierUnitCost: number | null;
  supplierTotalCost: number | null;
  // Side-by-side markup scenarios (cost * (1 + %))
  unitPrice10: number | null;
  unitPrice15: number | null;
  unitPrice18: number | null;
  unitPrice20: number | null;
  // Active markup unit price (PRICE CALC col R -> PI Unit Price)
  activeUnitPrice: number | null;
  activeTotalPrice: number | null;
  // Logistics
  boxes: number | null;
  cbm: number | null;
  grossKgCalc: number | null;
  // PL / DN values
  cartonNo: string;
  netWtKg: number | null;
  grossWtKg: number | null;
}

export interface JobCalculationSummary {
  rows: ComputedRow[];
  activeRowCount: number;
  totalQty: number;
  totalSupplierCost: number;
  totalBoxes: number;
  totalCbm: number;
  totalGrossKgCalc: number;
  // Freight calculation
  freightByCbm: number;
  freightByKg: number;
  freightBasis: 'CBM' | 'KG' | 'NONE';
  baseFreight: number;
  docChargeAmount: number;
  freightWithDoc: number;
  freightMarkupAmount: number;
  customerFreight: number; // Separate line after VAT
  // Customer totals (QT / PI / CI)
  goodsSubtotal: number;
  vatAmount: number; // Applied to goods subtotal ONLY
  grandTotal: number; // SUBTOTAL + VAT + FREIGHT
  // Profit & scenario comparisons
  scenarioTotals: Record<MarkupRate, {
    goodsSubtotal: number;
    vatAmount: number;
    customerFreight: number;
    grandTotal: number;
    grossProfit: number;
  }>;
  grossProfit: number;
  // Packing list / DN totals
  plTotalNetWt: number;
  plTotalGrossWt: number;
}
