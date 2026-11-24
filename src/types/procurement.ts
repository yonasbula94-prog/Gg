export type DocPrefix = 'QT' | 'PI' | 'PO' | 'PL' | 'DN' | 'CI';

export type MarkupRate = 0.1 | 0.15 | 0.18 | 0.2;

export type ThemeVariantId = 'v1' | 'v3';

export type InquiryType = 'RFQ' | 'TENDER';

export type DocumentFontId =
  | 'Calibri'
  | 'Arial'
  | 'Georgia'
  | 'Segoe UI'
  | 'Times New Roman'
  | 'Verdana';

export interface ThemeCustomization {
  primary?: string;
  accent?: string;
  totalBg?: string;
  secondaryBg?: string;
  headerText?: string;
  bodyText?: string;
  fontFamily?: DocumentFontId;
}

export type PlatformMode = 'windows' | 'ipad';

export type WorkspaceTab =
  | 'TRACKER'
  | 'INTERNAL_QUOTE'
  | 'PRICE_CALC'
  | 'QT'
  | 'PI'
  | 'PO'
  | 'CI'
  | 'PL'
  | 'DN'
  | 'REGISTRY'
  | 'FILES'
  | 'STUDIO'
  | 'HELP';

export interface DashboardPreferences {
  showRfqTenderCounters: boolean;
  showCompletionFunnel: boolean;
  showPerformanceChart: boolean;
  showPriorityRanking: boolean;
  showJobMatrix: boolean;
  showFinancials: boolean;
  includeCompleted: boolean;
  performancePeriod: 'weekly' | 'monthly';
  priorityMetric: 'profit' | 'value';
}

export interface SupplierQuoteData {
  supplierName: string;
  supplierContact: string;
  supplierAddress: string;
  supplierEmail: string;
  supplierQuoteRef: string;
  quoteDate: string;
  validUntil: string;
  currency: string;
  incoterm: string;
  paymentTerms: string;
  deliveryLeadTime: string;
  notes: string;
}

export interface JobWorkflowState {
  endUserPoConfirmed: boolean;
  endUserPoConfirmedAt?: string;
  supplierPaidConfirmed: boolean;
  supplierPaidConfirmedAt?: string;
  supplierPaymentRef: string;
  supplierPackingListReceived: boolean;
  supplierPackingListRef: string;
  supplierPackingListAt?: string;
}

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
  endUserAddresses: Record<string, string>;
}

export interface ThemePalette {
  id: ThemeVariantId;
  name: string;
  shortLabel: string;
  subtitle: string;
  // Typography for Print / Excel
  fontName: DocumentFontId;
  fontCss: string;
  // Hex colors for UI / Print
  primary: string;
  primaryText: string;
  bodyText: string;
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
  // PRICE CALC & INTERNAL QUOTE inputs (internal, never sent to customer)
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
  endUserSelect: string; // Dropdown from configured clients
  customEndUser: string; // Used when endUserSelect === 'OTHER'
  endUserContact: string;
  endUserAddress: string;
  currency: string;
  incoterm: string;
  destination: string;
  includeVat: boolean; // Internal PRICE CALC toggle: when true, VAT is folded into customer Unit Price
  vatRate: number; // Internal PRICE CALC VAT rate (default 0.05 = 5%, hidden from end-user docs)
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
  docNumber: string; // Hard-coded on issue, e.g., QT-GG-001-2026
  sequence: number; // Shared job sequence across all docs in the same job
  issuedAt: string;
  issuedBy: string;
  isVoided?: boolean;
  // Frozen financial totals at issuance
  frozenSubtotal: number;
  frozenVat: number;
  frozenFreight: number;
  frozenGrandTotal: number;
  // Revision control: Rev 1 is the original issue; later saves add Rev 2, Rev 3, ...
  revision?: number;
  // Previous revisions and their totals. Purged once the job is complete (Delivery Note issued).
  history?: IssuedRevision[];
}

export interface IssuedRevision {
  revision: number;
  docNumber: string;
  supersededAt: string;
  reason: string;
  frozenSubtotal: number;
  frozenVat: number;
  frozenFreight: number;
  frozenGrandTotal: number;
}

export interface FieldChange {
  field: string;
  from: string;
  to: string;
}

export interface AuditEntry {
  id: string;
  at: string;
  jobId: string;
  jobName: string;
  docNumber: string;
  prefix: DocPrefix;
  revision: number;
  reason: string;
  field: string;
  from: string;
  to: string;
}

export interface ProcurementJob {
  id: string;
  jobName: string;
  jobSequence?: number; // Assigned on first issued document; reused by all prefixes in this job
  jobSequenceYear?: number;
  inquiryType: InquiryType; // RFQ or TENDER
  createdAt: string;
  updatedAt: string;
  preparedBy: string;
  controls: JobControls;
  supplierQuote: SupplierQuoteData;
  workflow: JobWorkflowState;
  freight: FreightConfig;
  customColumns?: CustomColumnDef[];
  columnOrder?: string[]; // Custom visual order of standard + extra columns across all 6 docs
  items: JobItemRow[]; // Always length 20, rowNumber 17..36
  issuedDocs: Partial<Record<DocPrefix, IssuedDocRecord>>;
}

export interface RegistryEntry {
  sequence: number;
  docNumber: string; // Unique key: PREFIX-GG-SEQ-YEAR
  prefix: DocPrefix;
  year: number;
  jobId: string;
  jobName: string;
  customer: string; // Supplier name for PO; Client name for QT, PI, CI, PL, DN
  counterpartyRole?: 'CLIENT' | 'SUPPLIER';
  rfqRef?: string;
  currency: string;
  amount: number;
  issuedBy: string;
  issuedAt: string;
  status: 'ISSUED' | 'VOID';
  voidReason?: string;
  revision?: number;
}

export interface ValidationIssue {
  rowNumber?: number;
  field: string;
  message: string;
}

export interface WorkflowGateCheck {
  isReady: boolean;
  title: string;
  requirements: {
    id: string;
    label: string;
    met: boolean;
  }[];
  guidance: string;
  recommendedTab?: WorkspaceTab;
  recommendedTabLabel?: string;
}

export interface JobCompletionMetrics {
  completedSteps: number;
  totalSteps: number;
  percent: number;
  currentStageLabel: string;
  nextTab: WorkspaceTab;
  nextActionLabel: string;
  hasInternalQuote: boolean;
  hasQT: boolean;
  hasPI: boolean;
  hasEndUserPoGate: boolean;
  hasPO: boolean;
  hasCI: boolean;
  hasSupplierShipmentGate: boolean;
  hasPL: boolean;
  hasDN: boolean;
  isComplete: boolean;
}

export interface ConfirmDialogRequest {
  title: string;
  message: string;
  confirmLabel?: string;
  cancelLabel?: string;
  destructive?: boolean;
  onConfirm: () => void;
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
  // Side-by-side markup scenarios: cost * (1 + markup) * (1 + effectiveVat)
  unitPrice10: number | null;
  unitPrice15: number | null;
  unitPrice18: number | null;
  unitPrice20: number | null;
  // Active markup unit price (PRICE CALC col R -> QT / PI / CI Unit Price, includes internal VAT when enabled)
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
  customerFreight: number; // Separate line after Goods Subtotal
  // Internal VAT & Customer totals (QT / PI / CI never display a VAT line)
  effectiveVatRate: number; // 0 when includeVat is false, otherwise vatRate
  goodsSubtotalExVat: number; // Internal goods total before embedded VAT
  vatAmount: number; // Embedded internal VAT inside Unit Prices (for internal accounting only)
  goodsSubtotal: number; // Customer Goods Subtotal (sum of VAT-inclusive line totals)
  grandTotal: number; // Customer Goods Subtotal + Customer Freight
  // Profit & scenario comparisons
  scenarioTotals: Record<MarkupRate, {
    goodsSubtotalExVat: number;
    goodsSubtotal: number;
    vatAmount: number;
    customerFreight: number;
    grandTotal: number;
    grossProfit: number;
  }>;
  grossProfit: number; // Excludes embedded VAT pass-through
  // Packing list / DN totals
  plTotalNetWt: number;
  plTotalGrossWt: number;
}
