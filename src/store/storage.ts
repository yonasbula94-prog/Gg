import type { ContainerCounts } from '../types/procurement';
import defaultLogoUrl from '../assets/grand-group-logo.png?inline';
import {
  COMPANY_INFO,
  DEFAULT_DASHBOARD_PREFS,
  DEFAULT_END_USER_ADDRESSES,
  DOCUMENT_FONT_OPTIONS,
  END_USER_OPTIONS,
} from '../constants/themes';
import {
  AuditEntry,
  CompanyProfile,
  CustomColumnDef,
  CustomColumnPosition,
  DashboardPreferences,
  DocPrefix,
  DocumentFontId,
  InquiryType,
  IssuedDocRecord,
  JobItemRow,
  JobWorkflowState,
  ProcurementJob,
  RegistryEntry,
  SupplierQuoteData,
  ThemeCustomization,
  ThemeVariantId,
} from '../types/procurement';

export const STORAGE_KEYS = {
  JOBS: 'gg_fzc_procurement_jobs_v1',
  REGISTRY: 'gg_fzc_master_registry_v2',
  THEME: 'gg_fzc_excel_theme_v1',
  THEME_CUSTOM: 'gg_fzc_theme_custom_v1',
  ACTIVE_JOB: 'gg_fzc_active_job_id_v1',
  COMPANY: 'gg_fzc_company_profile_v1',
  ADMIN_PASS: 'gg_fzc_admin_pass_v1',
  DASHBOARD_PREFS: 'gg_fzc_dashboard_prefs_v1',
  AUDIT_LOG: 'gg_fzc_audit_log_v1',
};

export const DEFAULT_ADMIN_PASSWORD = 'GGAdmin2026';

// Deterministic offline-safe hash (works synchronously even under file:// where crypto.subtle may be blocked)
function hashAdminSecret(raw: string): string {
  const salted = `GGFZC::ADMIN::2026::${raw}`;
  let h1 = 0x811c9dc5;
  let h2 = 0x01000193;
  for (let i = 0; i < salted.length; i++) {
    const ch = salted.charCodeAt(i);
    h1 ^= ch;
    h1 = Math.imul(h1, 0x01000193);
    h2 ^= ch + ((i + 1) * 31);
    h2 = Math.imul(h2, 0x85ebca6b);
  }
  return `v1_${(h1 >>> 0).toString(16).padStart(8, '0')}_${(h2 >>> 0).toString(16).padStart(8, '0')}`;
}

export function hasCustomAdminPassword(): boolean {
  try {
    return Boolean(localStorage.getItem(STORAGE_KEYS.ADMIN_PASS));
  } catch {
    return false;
  }
}

export function verifyAdminPassword(candidate: string): boolean {
  try {
    const stored = localStorage.getItem(STORAGE_KEYS.ADMIN_PASS);
    const expected = stored || hashAdminSecret(DEFAULT_ADMIN_PASSWORD);
    return hashAdminSecret(candidate) === expected;
  } catch {
    return candidate === DEFAULT_ADMIN_PASSWORD;
  }
}

export function setAdminPassword(newPassword: string): boolean {
  const trimmed = newPassword.trim();
  if (trimmed.length < 4) return false;
  try {
    localStorage.setItem(STORAGE_KEYS.ADMIN_PASS, hashAdminSecret(trimmed));
    return true;
  } catch {
    return false;
  }
}

const VALID_PREFIXES: DocPrefix[] = ['QT', 'PI', 'PO', 'PL', 'DN', 'CI'];
const VALID_COL_POSITIONS: CustomColumnPosition[] = [
  'after_no',
  'after_requested',
  'after_offered',
  'after_qty',
  'end',
];

function safeStr(val: unknown, fallback = ''): string {
  return typeof val === 'string' ? val : fallback;
}

function safeNullableNonNeg(val: unknown): number | null {
  if (val === null || val === undefined || val === '') return null;
  const n = typeof val === 'number' ? val : Number(val);
  if (!Number.isFinite(n) || n < 0) return null;
  return n;
}

function safeRate(val: unknown, fallback: number, max: number): number {
  const n = typeof val === 'number' ? val : Number(val);
  if (!Number.isFinite(n)) return fallback;
  if (n < 0) return 0;
  if (n > max) return max;
  return n;
}

export function safeDeepClone<T>(value: T): T {
  try {
    if (typeof structuredClone === 'function') {
      return structuredClone(value);
    }
  } catch {
    // Fall back cleanly on older Safari/iPadOS engines
  }
  return JSON.parse(JSON.stringify(value)) as T;
}

export function createDefaultCompanyProfile(): CompanyProfile {
  return {
    ...COMPANY_INFO,
    logoDataUrl: defaultLogoUrl,
    endUsers: [...END_USER_OPTIONS],
    endUserAddresses: { ...DEFAULT_END_USER_ADDRESSES },
  };
}

export function normalizeCompanyProfile(raw: unknown): CompanyProfile {
  const defaults = createDefaultCompanyProfile();
  if (!raw || typeof raw !== 'object') return defaults;
  const parsed = raw as Partial<CompanyProfile>;
  const cleanedEndUsers = Array.isArray(parsed.endUsers)
    ? parsed.endUsers
        .filter((u): u is string => typeof u === 'string' && u.trim().length > 0)
        .map((u) => u.trim())
    : [];

  const uniqueEndUsers = Array.from(new Set(cleanedEndUsers));
  if (uniqueEndUsers.length > 0 && !uniqueEndUsers.includes('OTHER')) {
    uniqueEndUsers.push('OTHER');
  }

  // Migrate untouched legacy default strings to updated SAIF Zone & sales@grandgroupfzc.com defaults
  // while strictly preserving any customized admin edits across startup/reloads.
  const rawAddr1 = safeStr(parsed.addressLine1, defaults.addressLine1);
  const rawAddr2 = safeStr(parsed.addressLine2, defaults.addressLine2);
  const rawEmail = safeStr(parsed.email, defaults.email);

  const migratedAddr1 =
    rawAddr1.trim() === 'P.O. Box 42188, Hamriyah Free Zone'
      ? defaults.addressLine1
      : rawAddr1;
  const migratedAddr2 =
    rawAddr2.trim() === 'Sharjah, United Arab Emirates'
      ? defaults.addressLine2
      : rawAddr2;
  const migratedEmail =
    rawEmail.trim().toLowerCase() === 'procurement@grandgroupfzc.com'
      ? defaults.email
      : rawEmail;

  const mergedAddresses: Record<string, string> = { ...defaults.endUserAddresses };
  if (parsed.endUserAddresses && typeof parsed.endUserAddresses === 'object') {
    for (const [k, v] of Object.entries(parsed.endUserAddresses)) {
      if (typeof k === 'string' && k.trim() && typeof v === 'string') {
        mergedAddresses[k.trim()] = v;
      }
    }
  }

  return {
    name: safeStr(parsed.name, defaults.name) || defaults.name,
    tagline: safeStr(parsed.tagline, defaults.tagline),
    addressLine1: migratedAddr1,
    addressLine2: migratedAddr2,
    phone: safeStr(parsed.phone, defaults.phone),
    email: migratedEmail,
    trn: safeStr(parsed.trn, defaults.trn),
    bankDetails: safeStr(parsed.bankDetails, defaults.bankDetails),
    logoDataUrl:
      typeof parsed.logoDataUrl === 'string' && parsed.logoDataUrl.startsWith('data:image/')
        ? parsed.logoDataUrl
        : defaults.logoDataUrl,
    endUsers: uniqueEndUsers.length > 0 ? uniqueEndUsers : defaults.endUsers,
    endUserAddresses: mergedAddresses,
  };
}

export function loadCompanyProfile(): CompanyProfile {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.COMPANY);
    if (!raw) return createDefaultCompanyProfile();
    return normalizeCompanyProfile(JSON.parse(raw));
  } catch {
    return createDefaultCompanyProfile();
  }
}

export function saveCompanyProfile(profile: CompanyProfile): boolean {
  try {
    localStorage.setItem(STORAGE_KEYS.COMPANY, JSON.stringify(profile));
    return true;
  } catch {
    return false;
  }
}

export function normalizeThemeCustomization(raw: unknown): ThemeCustomization {
  if (!raw || typeof raw !== 'object') return {};
  const src = raw as Partial<ThemeCustomization>;
  const validHex = (val: unknown): string | undefined =>
    typeof val === 'string' && /^#[0-9A-Fa-f]{6}$/.test(val.trim())
      ? val.trim().toUpperCase()
      : undefined;
  const validFontIds = new Set<DocumentFontId>(DOCUMENT_FONT_OPTIONS.map((f) => f.id));
  const fontFamily =
    typeof src.fontFamily === 'string' && validFontIds.has(src.fontFamily as DocumentFontId)
      ? (src.fontFamily as DocumentFontId)
      : undefined;

  const out: ThemeCustomization = {};
  const p = validHex(src.primary);
  const a = validHex(src.accent);
  const t = validHex(src.totalBg);
  const s = validHex(src.secondaryBg);
  const ht = validHex(src.headerText);
  const bt = validHex(src.bodyText);
  if (p) out.primary = p;
  if (a) out.accent = a;
  if (t) out.totalBg = t;
  if (s) out.secondaryBg = s;
  if (ht) out.headerText = ht;
  if (bt) out.bodyText = bt;
  if (fontFamily) out.fontFamily = fontFamily;
  return out;
}

export function loadThemeCustomization(): ThemeCustomization {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.THEME_CUSTOM);
    if (!raw) return {};
    return normalizeThemeCustomization(JSON.parse(raw));
  } catch {
    return {};
  }
}

export function saveThemeCustomization(custom: ThemeCustomization): boolean {
  try {
    localStorage.setItem(STORAGE_KEYS.THEME_CUSTOM, JSON.stringify(custom));
    return true;
  } catch {
    return false;
  }
}

export function createEmptyItems(): JobItemRow[] {
  return Array.from({ length: 20 }, (_, i) => ({
    rowNumber: 17 + i,
    requestedDescription: '',
    offeredDescription: '',
    uom: '',
    qty: null,
    customFields: {},
    supplierUnitCost: null,
    boxes: null,
    boxLengthCm: null,
    boxWidthCm: null,
    boxHeightCm: null,
    kgPerBox: null,
    cartonNo: '',
    netWtKg: null,
    grossWtKg: null,
  }));
}

export function createDefaultSupplierQuote(): SupplierQuoteData {
  return {
    supplierName: '',
    supplierContact: '',
    supplierAddress: '',
    supplierEmail: '',
    supplierQuoteRef: '',
    quoteDate: new Date().toISOString().slice(0, 10),
    validUntil: '',
    currency: 'USD',
    incoterm: 'EXW',
    paymentTerms: '100% TT Prior to Dispatch',
    deliveryLeadTime: '2–3 Weeks from Order',
    notes: '',
  };
}

export function createDefaultWorkflow(): JobWorkflowState {
  return {
    endUserPoConfirmed: false,
    supplierPaidConfirmed: false,
    supplierPaymentRef: '',
    supplierPackingListReceived: false,
    supplierPackingListRef: '',
  };
}

export function inferInquiryType(rfqRef?: string): InquiryType {
  const s = (rfqRef || '').toLowerCase();
  if (s.includes('tender') || s.includes('tnd') || s.includes('itb') || s.includes('rfp')) {
    return 'TENDER';
  }
  return 'RFQ';
}

export function createBlankJob(name?: string, company?: CompanyProfile): ProcurementJob {
  const now = new Date().toISOString();
  const id = `job-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
  const defaultEndUser = company?.endUsers?.[0] || 'AIAA (Asmara International Airport)';
  const defaultAddress =
    defaultEndUser === 'OTHER'
      ? ''
      : company?.endUserAddresses?.[defaultEndUser] ??
        DEFAULT_END_USER_ADDRESSES[defaultEndUser] ??
        '';
  return {
    id,
    jobName: name || 'Procurement Job #1',
    inquiryType: 'RFQ',
    createdAt: now,
    updatedAt: now,
    preparedBy: 'Procurement Desk',
    controls: {
      endUserSelect: defaultEndUser,
      customEndUser: '',
      endUserContact: '',
      endUserAddress: defaultAddress,
      currency: 'USD',
      incoterm: 'CIF',
      destination: 'Massawa Seaport',
      includeVat: true,
      vatRate: 0.05,
      paymentTerms: '50% Advance with PO, 50% Prior to Shipment',
      deliveryLeadTime: '4–6 Weeks from PO Confirmation',
      rfqRef: '',
      customerPoRef: '',
      supplierName: '',
      supplierAddress: '',
      supplierRef: '',
      shipDate: '',
      portOfLoading: 'Jebel Ali Port, UAE',
      deliveryDate: '',
      transportRef: '',
      receivedBy: '',
      remarks: 'All items supplied in original manufacturer export-grade packaging.',
      authorizedSignatory: 'Authorized Signature — GRAND GROUP FZC',
      counterpartySignatory: 'Client / Supplier Acknowledgment',
    },
    supplierQuote: createDefaultSupplierQuote(),
    workflow: createDefaultWorkflow(),
    freight: {
      ratePerCbm: 185,
      ratePerKg: 2.4,
      docChargePct: 0.1,
      selectedMarkupPct: 0.2,
    },
    customColumns: [],
    items: createEmptyItems(),
    issuedDocs: {},
  };
}

// Clean empty registry starting at sequence 0 (next claimed document number is #001)
export function createSeededRegistry(): RegistryEntry[] {
  return [];
}

function normalizeContainerCounts(raw: unknown): ContainerCounts | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const r = raw as Record<string, unknown>;
  const pick = (k: keyof ContainerCounts) => {
    const n = Number(r[k]);
    return Number.isFinite(n) && n > 0 ? n : 0;
  };
  return { c20: pick('c20'), c40: pick('c40'), c40hc: pick('c40hc') };
}

export function normalizeJob(raw: unknown): ProcurementJob {
  const fallback = createBlankJob('Procurement Job #1');
  if (!raw || typeof raw !== 'object') return fallback;
  const src = raw as Partial<ProcurementJob>;
  const rawCtrl = (src.controls && typeof src.controls === 'object' ? src.controls : {}) as Partial<
    ProcurementJob['controls']
  >;
  const rawFreight = (src.freight && typeof src.freight === 'object' ? src.freight : {}) as Partial<
    ProcurementJob['freight']
  >;

  const vatRate = safeRate(rawCtrl.vatRate, 0.05, 1);
  // Backward-compatible: older saved jobs without `includeVat` infer true if vatRate > 0
  const includeVat =
    typeof rawCtrl.includeVat === 'boolean' ? rawCtrl.includeVat : vatRate > 0;

  const customColumns: CustomColumnDef[] = Array.isArray(src.customColumns)
    ? src.customColumns
        .filter(
          (c): c is CustomColumnDef =>
            Boolean(c && typeof c === 'object' && typeof c.id === 'string' && typeof c.label === 'string' && c.label.trim())
        )
        .slice(0, 4)
        .map((c) => ({
          id: c.id.trim(),
          label: c.label.trim(),
          position:
            c.position && VALID_COL_POSITIONS.includes(c.position) ? c.position : 'after_no',
        }))
    : [];

  const columnOrder: string[] | undefined = Array.isArray(src.columnOrder)
    ? Array.from(
        new Set(
          src.columnOrder.filter(
            (k): k is string => typeof k === 'string' && k.trim().length > 0
          )
        )
      )
    : undefined;

  const rawItems = Array.isArray(src.items) ? src.items : [];
  const items: JobItemRow[] = Array.from({ length: 20 }, (_, idx) => {
    const it = (rawItems[idx] && typeof rawItems[idx] === 'object'
      ? rawItems[idx]
      : {}) as Partial<JobItemRow>;
    const rawCustom =
      it.customFields && typeof it.customFields === 'object' ? it.customFields : {};
    const cleanCustom: Record<string, string> = {};
    for (const [k, v] of Object.entries(rawCustom)) {
      if (typeof v === 'string') cleanCustom[k] = v;
    }
    return {
      rowNumber: 17 + idx,
      requestedDescription: safeStr(it.requestedDescription),
      offeredDescription: safeStr(it.offeredDescription),
      uom: safeStr(it.uom),
      qty: safeNullableNonNeg(it.qty),
      customFields: cleanCustom,
      supplierUnitCost: safeNullableNonNeg(it.supplierUnitCost),
      boxes: safeNullableNonNeg(it.boxes),
      boxLengthCm: safeNullableNonNeg(it.boxLengthCm),
      boxWidthCm: safeNullableNonNeg(it.boxWidthCm),
      boxHeightCm: safeNullableNonNeg(it.boxHeightCm),
      kgPerBox: safeNullableNonNeg(it.kgPerBox),
      cartonNo: safeStr(it.cartonNo),
      netWtKg: safeNullableNonNeg(it.netWtKg),
      grossWtKg: safeNullableNonNeg(it.grossWtKg),
    };
  });

  const issuedDocs: Partial<Record<DocPrefix, IssuedDocRecord>> = {};
  if (src.issuedDocs && typeof src.issuedDocs === 'object') {
    for (const prefix of VALID_PREFIXES) {
      const rec = (src.issuedDocs as Record<string, unknown>)[prefix];
      if (rec && typeof rec === 'object') {
        const r = rec as Partial<IssuedDocRecord>;
        if (typeof r.docNumber === 'string' && typeof r.sequence === 'number' && Number.isFinite(r.sequence)) {
          issuedDocs[prefix] = {
            prefix,
            docNumber: r.docNumber,
            sequence: Math.floor(r.sequence),
            issuedAt: safeStr(r.issuedAt, fallback.createdAt),
            issuedBy: safeStr(r.issuedBy, 'Procurement Desk'),
            isVoided: Boolean(r.isVoided),
            frozenSubtotal: safeNullableNonNeg(r.frozenSubtotal) ?? 0,
            frozenVat: safeNullableNonNeg(r.frozenVat) ?? 0,
            frozenFreight: safeNullableNonNeg(r.frozenFreight) ?? 0,
            frozenGrandTotal: safeNullableNonNeg(r.frozenGrandTotal) ?? 0,
          };
        }
      }
    }
  }

  const rawSq = (src.supplierQuote && typeof src.supplierQuote === 'object'
    ? src.supplierQuote
    : {}) as Partial<SupplierQuoteData>;
  const rawWf = (src.workflow && typeof src.workflow === 'object'
    ? src.workflow
    : {}) as Partial<JobWorkflowState>;

  const supplierName = safeStr(rawCtrl.supplierName) || safeStr(rawSq.supplierName);
  const supplierAddress = safeStr(rawCtrl.supplierAddress) || safeStr(rawSq.supplierAddress);
  const supplierRef = safeStr(rawCtrl.supplierRef) || safeStr(rawSq.supplierQuoteRef);
  const rfqRef = safeStr(rawCtrl.rfqRef);

  let jobSequence: number | undefined =
    typeof src.jobSequence === 'number' && Number.isFinite(src.jobSequence) && src.jobSequence > 0
      ? Math.floor(src.jobSequence)
      : undefined;
  let jobSequenceYear: number | undefined =
    typeof src.jobSequenceYear === 'number' &&
    Number.isFinite(src.jobSequenceYear) &&
    src.jobSequenceYear > 2000
      ? Math.floor(src.jobSequenceYear)
      : undefined;

  if (jobSequence === undefined) {
    const firstIssued = Object.values(issuedDocs).find((d) => d && d.sequence > 0);
    if (firstIssued) {
      jobSequence = firstIssued.sequence;
      const yrMatch = firstIssued.docNumber.match(/(\d{4})$/);
      jobSequenceYear = yrMatch ? Number(yrMatch[1]) : new Date().getFullYear();
    }
  }

  const inquiryType: InquiryType =
    src.inquiryType === 'TENDER' || src.inquiryType === 'RFQ'
      ? src.inquiryType
      : inferInquiryType(rfqRef);

  const sqDefaults = createDefaultSupplierQuote();

  return {
    id: safeStr(src.id, fallback.id) || fallback.id,
    jobName: safeStr(src.jobName, fallback.jobName) || fallback.jobName,
    jobSequence,
    jobSequenceYear,
    inquiryType,
    createdAt: safeStr(src.createdAt, fallback.createdAt),
    updatedAt: safeStr(src.updatedAt, fallback.updatedAt),
    preparedBy: safeStr(src.preparedBy, fallback.preparedBy),
    controls: {
      endUserSelect: safeStr(rawCtrl.endUserSelect, fallback.controls.endUserSelect),
      customEndUser: safeStr(rawCtrl.customEndUser, fallback.controls.customEndUser),
      endUserContact: safeStr(rawCtrl.endUserContact, fallback.controls.endUserContact),
      endUserAddress: safeStr(rawCtrl.endUserAddress, fallback.controls.endUserAddress),
      currency: safeStr(rawCtrl.currency, fallback.controls.currency) || 'USD',
      incoterm: safeStr(rawCtrl.incoterm, fallback.controls.incoterm) || 'CIF',
      destination: safeStr(rawCtrl.destination, fallback.controls.destination),
      includeVat,
      vatRate,
      paymentTerms: safeStr(rawCtrl.paymentTerms, fallback.controls.paymentTerms),
      deliveryLeadTime: safeStr(rawCtrl.deliveryLeadTime, fallback.controls.deliveryLeadTime),
      rfqRef,
      customerPoRef: safeStr(rawCtrl.customerPoRef),
      supplierName,
      supplierAddress,
      supplierRef,
      shipDate: safeStr(rawCtrl.shipDate),
      portOfLoading: safeStr(rawCtrl.portOfLoading, fallback.controls.portOfLoading),
      deliveryDate: safeStr(rawCtrl.deliveryDate),
      transportRef: safeStr(rawCtrl.transportRef),
      receivedBy: safeStr(rawCtrl.receivedBy),
      remarks:
        typeof rawCtrl.remarks === 'string' ? rawCtrl.remarks : fallback.controls.remarks,
      authorizedSignatory:
        typeof rawCtrl.authorizedSignatory === 'string'
          ? rawCtrl.authorizedSignatory
          : fallback.controls.authorizedSignatory,
      counterpartySignatory:
        typeof rawCtrl.counterpartySignatory === 'string'
          ? rawCtrl.counterpartySignatory
          : fallback.controls.counterpartySignatory,
    },
    supplierQuote: {
      supplierName,
      supplierContact: safeStr(rawSq.supplierContact),
      supplierAddress,
      supplierEmail: safeStr(rawSq.supplierEmail),
      supplierQuoteRef: supplierRef,
      quoteDate: safeStr(rawSq.quoteDate, sqDefaults.quoteDate),
      validUntil: safeStr(rawSq.validUntil),
      currency: safeStr(rawSq.currency, safeStr(rawCtrl.currency, 'USD') || 'USD'),
      incoterm: safeStr(rawSq.incoterm, sqDefaults.incoterm),
      paymentTerms: safeStr(rawSq.paymentTerms, sqDefaults.paymentTerms),
      deliveryLeadTime: safeStr(rawSq.deliveryLeadTime, sqDefaults.deliveryLeadTime),
      notes: safeStr(rawSq.notes),
    },
    workflow: {
      endUserPoConfirmed:
        typeof rawWf.endUserPoConfirmed === 'boolean'
          ? rawWf.endUserPoConfirmed
          : Boolean(issuedDocs.PO || issuedDocs.CI),
      endUserPoConfirmedAt:
        typeof rawWf.endUserPoConfirmedAt === 'string' ? rawWf.endUserPoConfirmedAt : undefined,
      supplierPaidConfirmed:
        typeof rawWf.supplierPaidConfirmed === 'boolean'
          ? rawWf.supplierPaidConfirmed
          : Boolean(issuedDocs.PL || issuedDocs.DN),
      supplierPaidConfirmedAt:
        typeof rawWf.supplierPaidConfirmedAt === 'string'
          ? rawWf.supplierPaidConfirmedAt
          : undefined,
      supplierPaymentRef: safeStr(rawWf.supplierPaymentRef),
      supplierPackingListReceived:
        typeof rawWf.supplierPackingListReceived === 'boolean'
          ? rawWf.supplierPackingListReceived
          : Boolean(issuedDocs.PL || issuedDocs.DN),
      supplierPackingListRef: safeStr(rawWf.supplierPackingListRef),
      supplierPackingListAt:
        typeof rawWf.supplierPackingListAt === 'string' ? rawWf.supplierPackingListAt : undefined,
    },
    freight: {
      ratePerCbm: safeNullableNonNeg(rawFreight.ratePerCbm) ?? fallback.freight.ratePerCbm,
      ratePerKg: safeNullableNonNeg(rawFreight.ratePerKg) ?? fallback.freight.ratePerKg,
      docChargePct: safeRate(rawFreight.docChargePct, fallback.freight.docChargePct, 1),
      selectedMarkupPct: safeRate(
        rawFreight.selectedMarkupPct,
        fallback.freight.selectedMarkupPct,
        5
      ),
      // Sea freight and freight markup. Missing values stay undefined: freight then follows the item markup (existing jobs)
      mode: rawFreight.mode === 'sea' ? 'sea' : 'air',
      containers: normalizeContainerCounts(rawFreight.containers),
      containerRates: normalizeContainerCounts(rawFreight.containerRates),
      loadingChargePct:
        rawFreight.loadingChargePct === undefined
          ? undefined
          : safeRate(rawFreight.loadingChargePct, 0.05, 1),
      freightMarkupPct:
        rawFreight.freightMarkupPct === undefined
          ? undefined
          : safeRate(rawFreight.freightMarkupPct, fallback.freight.selectedMarkupPct, 5),
    },
    customColumns,
    columnOrder,
    items,
    issuedDocs,
  };
}

export function normalizeRegistryEntries(raw: unknown): RegistryEntry[] {
  if (!Array.isArray(raw)) return [];
  const valid: RegistryEntry[] = [];
  for (const item of raw) {
    if (!item || typeof item !== 'object') continue;
    const e = item as Partial<RegistryEntry>;
    const prefix = typeof e.prefix === 'string' ? (e.prefix.toUpperCase() as DocPrefix) : null;
    if (
      typeof e.sequence !== 'number' ||
      !Number.isFinite(e.sequence) ||
      !prefix ||
      !VALID_PREFIXES.includes(prefix) ||
      typeof e.docNumber !== 'string' ||
      !e.docNumber.trim()
    ) {
      continue;
    }
    if (e.jobId === 'legacy-seed-042') continue;
    valid.push({
      sequence: Math.floor(e.sequence),
      docNumber: e.docNumber.trim(),
      prefix,
      year:
        typeof e.year === 'number' && Number.isFinite(e.year)
          ? Math.floor(e.year)
          : new Date().getFullYear(),
      jobId: safeStr(e.jobId, `reg-${e.sequence}`),
      jobName: safeStr(e.jobName, 'Untitled Job'),
      customer: safeStr(e.customer, '—'),
      counterpartyRole: prefix === 'PO' ? 'SUPPLIER' : 'CLIENT',
      rfqRef: safeStr(e.rfqRef, ''),
      currency: safeStr(e.currency, 'USD') || 'USD',
      amount: safeNullableNonNeg(e.amount) ?? 0,
      issuedBy: safeStr(e.issuedBy, 'Procurement Desk'),
      issuedAt: safeStr(e.issuedAt, new Date().toISOString()),
      status: e.status === 'VOID' ? 'VOID' : 'ISSUED',
      voidReason: typeof e.voidReason === 'string' ? e.voidReason : undefined,
    });
  }
  return valid;
}

export function loadJobsWithHealth(): { jobs: ProcurementJob[]; recovered: boolean } {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.JOBS);
    if (!raw) return { jobs: [createBlankJob('Procurement Job #1')], recovered: false };
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed) || parsed.length === 0) {
      return { jobs: [createBlankJob('Procurement Job #1')], recovered: true };
    }
    const validRaw = parsed.filter((entry) => entry && typeof entry === 'object');
    if (validRaw.length === 0) {
      return { jobs: [createBlankJob('Procurement Job #1')], recovered: true };
    }
    return {
      jobs: validRaw.map((j) => normalizeJob(j)),
      recovered: validRaw.length !== parsed.length,
    };
  } catch {
    return { jobs: [createBlankJob('Procurement Job #1')], recovered: true };
  }
}

export function saveJobs(jobs: ProcurementJob[]): boolean {
  try {
    localStorage.setItem(STORAGE_KEYS.JOBS, JSON.stringify(jobs));
    return true;
  } catch {
    return false;
  }
}

export function loadRegistryWithHealth(): { entries: RegistryEntry[]; recovered: boolean } {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.REGISTRY);
    if (!raw) return { entries: [], recovered: false };
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return { entries: [], recovered: true };
    }
    const cleaned = normalizeRegistryEntries(parsed);
    return {
      entries: cleaned,
      recovered: cleaned.length !== parsed.length,
    };
  } catch {
    return { entries: [], recovered: true };
  }
}

export function saveRegistry(entries: RegistryEntry[]): boolean {
  try {
    localStorage.setItem(STORAGE_KEYS.REGISTRY, JSON.stringify(entries));
    return true;
  } catch {
    return false;
  }
}

export function loadThemeVariant(): ThemeVariantId {
  try {
    const val = localStorage.getItem(STORAGE_KEYS.THEME);
    if (val === 'v1' || val === 'v3') return val; // v2 was retired; fall back to v1
    return 'v1';
  } catch {
    return 'v1';
  }
}

export function saveThemeVariant(theme: ThemeVariantId): void {
  try {
    localStorage.setItem(STORAGE_KEYS.THEME, theme);
  } catch {
    // Ignore theme preference errors
  }
}

export function loadActiveJobId(jobs: ProcurementJob[]): string {
  try {
    const saved = localStorage.getItem(STORAGE_KEYS.ACTIVE_JOB);
    if (saved && jobs.some((j) => j.id === saved)) return saved;
    return jobs[0].id;
  } catch {
    return jobs[0].id;
  }
}

export function saveActiveJobId(id: string): void {
  try {
    localStorage.setItem(STORAGE_KEYS.ACTIVE_JOB, id);
  } catch {
    // Ignore active-job pointer errors
  }
}

export function loadDashboardPrefs(): DashboardPreferences {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.DASHBOARD_PREFS);
    if (!raw) return { ...DEFAULT_DASHBOARD_PREFS };
    const parsed = JSON.parse(raw) as Partial<DashboardPreferences>;
    if (!parsed || typeof parsed !== 'object') return { ...DEFAULT_DASHBOARD_PREFS };
    return {
      showRfqTenderCounters:
        typeof parsed.showRfqTenderCounters === 'boolean'
          ? parsed.showRfqTenderCounters
          : DEFAULT_DASHBOARD_PREFS.showRfqTenderCounters,
      showCompletionFunnel:
        typeof parsed.showCompletionFunnel === 'boolean'
          ? parsed.showCompletionFunnel
          : DEFAULT_DASHBOARD_PREFS.showCompletionFunnel,
      showPerformanceChart:
        typeof parsed.showPerformanceChart === 'boolean'
          ? parsed.showPerformanceChart
          : DEFAULT_DASHBOARD_PREFS.showPerformanceChart,
      showPriorityRanking:
        typeof parsed.showPriorityRanking === 'boolean'
          ? parsed.showPriorityRanking
          : DEFAULT_DASHBOARD_PREFS.showPriorityRanking,
      showJobMatrix:
        typeof parsed.showJobMatrix === 'boolean'
          ? parsed.showJobMatrix
          : DEFAULT_DASHBOARD_PREFS.showJobMatrix,
      showFinancials:
        typeof parsed.showFinancials === 'boolean'
          ? parsed.showFinancials
          : DEFAULT_DASHBOARD_PREFS.showFinancials,
      includeCompleted:
        typeof parsed.includeCompleted === 'boolean'
          ? parsed.includeCompleted
          : DEFAULT_DASHBOARD_PREFS.includeCompleted,
      performancePeriod: parsed.performancePeriod === 'monthly' ? 'monthly' : 'weekly',
      priorityMetric: parsed.priorityMetric === 'value' ? 'value' : 'profit',
    };
  } catch {
    return { ...DEFAULT_DASHBOARD_PREFS };
  }
}

export function saveDashboardPrefs(prefs: DashboardPreferences): void {
  try {
    localStorage.setItem(STORAGE_KEYS.DASHBOARD_PREFS, JSON.stringify(prefs));
  } catch {
    // Ignore prefs save error
  }
}

export function getNextSequenceNumber(
  entries: RegistryEntry[],
  jobs: ProcurementJob[] = []
): number {
  const maxRegSeq = entries.reduce((max, e) => (e.sequence > max ? e.sequence : max), 0);
  const maxJobSeq = jobs.reduce(
    (max, j) => (j.jobSequence && j.jobSequence > max ? j.jobSequence : max),
    0
  );
  return Math.max(maxRegSeq, maxJobSeq) + 1;
}

export function formatDocNumber(prefix: DocPrefix, sequence: number, year?: number): string {
  const yr = year ?? new Date().getFullYear();
  const padded = String(sequence).padStart(3, '0');
  return `${prefix}-GG-${padded}-${yr}`;
}

export function exportEmergencyRawStorageJson(): string {
  const dump: Record<string, unknown> = {
    exportedAt: new Date().toISOString(),
    emergencyRecovery: true,
  };
  for (const [label, key] of Object.entries(STORAGE_KEYS)) {
    try {
      const val = localStorage.getItem(key);
      if (val !== null) {
        try {
          dump[label] = JSON.parse(val);
        } catch {
          dump[label] = val;
        }
      }
    } catch {
      dump[label] = null;
    }
  }
  return JSON.stringify(dump, null, 2);
}

export function clearAllAppStorage(): void {
  // The Master Registry is kept on purpose: clearing it would let issued numbers be claimed again.
  for (const key of Object.values(STORAGE_KEYS).filter((k) => k !== STORAGE_KEYS.REGISTRY && k !== STORAGE_KEYS.AUDIT_LOG)) {
    try {
      localStorage.removeItem(key);
    } catch {
      // Ignore
    }
  }
}

/**
 * Merges incoming registry rows into existing ones by document number.
 * Numbers are never removed, and a VOID entry is never reverted to ISSUED.
 */
export function mergeRegistryEntries(base: RegistryEntry[], incoming: RegistryEntry[]): RegistryEntry[] {
  const byDoc = new Map<string, RegistryEntry>();
  base.forEach((e) => byDoc.set(e.docNumber, e));
  incoming.forEach((e) => {
    const existing = byDoc.get(e.docNumber);
    if (existing && existing.status === 'VOID' && e.status !== 'VOID') return;
    byDoc.set(e.docNumber, existing && existing.status === 'VOID' ? existing : e);
  });
  return Array.from(byDoc.values()).sort(
    (a, b) => a.sequence - b.sequence || a.docNumber.localeCompare(b.docNumber)
  );
}

/** Change log for issued documents. Kept separate from job data and survives a storage reset. */
export function loadAuditLog(): AuditEntry[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.AUDIT_LOG);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? (parsed as AuditEntry[]) : [];
  } catch {
    return [];
  }
}

export function appendAuditLog(entries: AuditEntry[]): boolean {
  if (entries.length === 0) return true;
  try {
    const next = [...loadAuditLog(), ...entries];
    localStorage.setItem(STORAGE_KEYS.AUDIT_LOG, JSON.stringify(next));
    return true;
  } catch {
    return false;
  }
}
