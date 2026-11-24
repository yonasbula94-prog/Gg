import defaultLogoUrl from '../assets/grand-group-logo.png?inline';
import { COMPANY_INFO, END_USER_OPTIONS } from '../constants/themes';
import {
  CompanyProfile,
  CustomColumnDef,
  CustomColumnPosition,
  DocPrefix,
  IssuedDocRecord,
  JobItemRow,
  ProcurementJob,
  RegistryEntry,
  ThemeVariantId,
} from '../types/procurement';

export const STORAGE_KEYS = {
  JOBS: 'gg_fzc_procurement_jobs_v1',
  REGISTRY: 'gg_fzc_master_registry_v1',
  THEME: 'gg_fzc_excel_theme_v1',
  ACTIVE_JOB: 'gg_fzc_active_job_id_v1',
  COMPANY: 'gg_fzc_company_profile_v1',
};

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

  return {
    name: safeStr(parsed.name, defaults.name) || defaults.name,
    tagline: safeStr(parsed.tagline, defaults.tagline),
    addressLine1: safeStr(parsed.addressLine1, defaults.addressLine1),
    addressLine2: safeStr(parsed.addressLine2, defaults.addressLine2),
    phone: safeStr(parsed.phone, defaults.phone),
    email: safeStr(parsed.email, defaults.email),
    trn: safeStr(parsed.trn, defaults.trn),
    bankDetails: safeStr(parsed.bankDetails, defaults.bankDetails),
    logoDataUrl:
      typeof parsed.logoDataUrl === 'string' && parsed.logoDataUrl.startsWith('data:image/')
        ? parsed.logoDataUrl
        : defaults.logoDataUrl,
    endUsers: uniqueEndUsers.length > 0 ? uniqueEndUsers : defaults.endUsers,
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

export function createBlankJob(name?: string): ProcurementJob {
  const now = new Date().toISOString();
  const id = `job-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
  return {
    id,
    jobName: name || 'Procurement Job #1',
    createdAt: now,
    updatedAt: now,
    preparedBy: 'Procurement Desk',
    controls: {
      endUserSelect: 'AIAA (Asmara International Airport)',
      customEndUser: '',
      endUserContact: 'Procurement Department',
      endUserAddress: 'Asmara, Eritrea',
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
      counterpartySignatory: 'Customer / Supplier Acknowledgment',
    },
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

// Seeded with PI-GG-042-2026 so the next claimed document number is strictly 43
export function createSeededRegistry(): RegistryEntry[] {
  return [
    {
      sequence: 42,
      docNumber: 'PI-GG-042-2026',
      prefix: 'PI',
      year: 2026,
      jobId: 'legacy-seed-042',
      jobName: 'Massawa Port Hydraulics Maintenance Kit',
      customer: 'Massawa Port',
      currency: 'USD',
      amount: 14820.0,
      issuedBy: 'Admin Seed',
      issuedAt: '2026-11-18T09:15:00.000Z',
      status: 'ISSUED',
    },
  ];
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

  return {
    id: safeStr(src.id, fallback.id) || fallback.id,
    jobName: safeStr(src.jobName, fallback.jobName) || fallback.jobName,
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
      rfqRef: safeStr(rawCtrl.rfqRef),
      customerPoRef: safeStr(rawCtrl.customerPoRef),
      supplierName: safeStr(rawCtrl.supplierName),
      supplierAddress: safeStr(rawCtrl.supplierAddress),
      supplierRef: safeStr(rawCtrl.supplierRef),
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
    freight: {
      ratePerCbm: safeNullableNonNeg(rawFreight.ratePerCbm) ?? fallback.freight.ratePerCbm,
      ratePerKg: safeNullableNonNeg(rawFreight.ratePerKg) ?? fallback.freight.ratePerKg,
      docChargePct: safeRate(rawFreight.docChargePct, fallback.freight.docChargePct, 1),
      selectedMarkupPct: safeRate(
        rawFreight.selectedMarkupPct,
        fallback.freight.selectedMarkupPct,
        5
      ),
    },
    customColumns,
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
    if (!raw) return { entries: createSeededRegistry(), recovered: false };
    const parsed = JSON.parse(raw);
    const cleaned = normalizeRegistryEntries(parsed);
    if (cleaned.length === 0) {
      return { entries: createSeededRegistry(), recovered: true };
    }
    return {
      entries: cleaned,
      recovered: Array.isArray(parsed) && cleaned.length !== parsed.length,
    };
  } catch {
    return { entries: createSeededRegistry(), recovered: true };
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
    if (val === 'v1' || val === 'v2' || val === 'v3') return val;
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

export function getNextSequenceNumber(entries: RegistryEntry[]): number {
  const maxSeq = entries.reduce((max, e) => (e.sequence > max ? e.sequence : max), 42);
  return maxSeq + 1;
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
  for (const key of Object.values(STORAGE_KEYS)) {
    try {
      localStorage.removeItem(key);
    } catch {
      // Ignore
    }
  }
}
