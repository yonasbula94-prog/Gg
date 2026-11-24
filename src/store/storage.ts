import defaultLogoUrl from '../assets/grand-group-logo.png?inline';
import { COMPANY_INFO, END_USER_OPTIONS } from '../constants/themes';
import {
  CompanyProfile,
  DocPrefix,
  JobItemRow,
  ProcurementJob,
  RegistryEntry,
  ThemeVariantId,
} from '../types/procurement';

const STORAGE_KEYS = {
  JOBS: 'gg_fzc_procurement_jobs_v1',
  REGISTRY: 'gg_fzc_master_registry_v1',
  THEME: 'gg_fzc_excel_theme_v1',
  ACTIVE_JOB: 'gg_fzc_active_job_id_v1',
  COMPANY: 'gg_fzc_company_profile_v1',
};

export function createDefaultCompanyProfile(): CompanyProfile {
  return {
    ...COMPANY_INFO,
    logoDataUrl: defaultLogoUrl,
    endUsers: [...END_USER_OPTIONS],
  };
}

export function loadCompanyProfile(): CompanyProfile {
  const defaults = createDefaultCompanyProfile();
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.COMPANY);
    if (!raw) return defaults;
    const parsed = JSON.parse(raw) as Partial<CompanyProfile>;
    return {
      ...defaults,
      ...parsed,
      logoDataUrl: parsed.logoDataUrl || defaults.logoDataUrl,
      endUsers:
        Array.isArray(parsed.endUsers) && parsed.endUsers.length > 0
          ? parsed.endUsers
          : defaults.endUsers,
    };
  } catch {
    return defaults;
  }
}

export function saveCompanyProfile(profile: CompanyProfile): void {
  try {
    localStorage.setItem(STORAGE_KEYS.COMPANY, JSON.stringify(profile));
  } catch {
    // Ignore storage quota errors
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
    jobName: name || `New Procurement Job`,
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
      vatRate: 0.05,
      quoteValidityDays: 30,
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
      selectedMarkupPct: 0.15,
    },
    customColumns: [],
    items: createEmptyItems(),
    issuedDocs: {},
  };
}

export function createSampleJob(): ProcurementJob {
  const job = createBlankJob('AIAA Runway Lighting & Spares');
  job.id = 'job-sample-aiaa-2026';
  job.controls = {
    ...job.controls,
    endUserSelect: 'AIAA (Asmara International Airport)',
    endUserContact: 'Eng. Tesfay Ghebrehiwet — Technical Procurement',
    endUserAddress: 'Asmara International Airport Authority, Asmara',
    currency: 'USD',
    incoterm: 'CIF',
    destination: 'Massawa Seaport',
    vatRate: 0.05,
    rfqRef: 'RFQ-AIAA-2026-019',
    customerPoRef: 'PO-AIAA-8841',
    supplierName: 'Al-Futtaim Industrial & Aviation Supplies LLC',
    supplierAddress: 'Jebel Ali Free Zone South, Dubai, UAE',
    supplierRef: 'SQ-AFI-2026-412',
    shipDate: '2026-12-08',
    portOfLoading: 'Jebel Ali Port, UAE',
    deliveryDate: '2026-12-19',
    transportRef: 'BL-MSC-774920 / Truck ER-402',
    receivedBy: 'Storekeeper — AIAA Central Warehouse',
    counterpartySignatory: 'Received in Good Order — Storekeeper — AIAA Central Warehouse',
  };

  const seededRows: Partial<JobItemRow>[] = [
    {
      requestedDescription: 'LED Elevated Runway Edge Light, Bi-Directional 6.6A',
      offeredDescription: 'ADB Safegate EL-E-LED Elevated Edge Light Assembly, FAA L-862, 6.6A',
      uom: 'PCS',
      qty: 24,
      supplierUnitCost: 285,
      boxes: 4,
      boxLengthCm: 60,
      boxWidthCm: 45,
      boxHeightCm: 40,
      kgPerBox: 18.5,
      cartonNo: 'CTN 01–04',
      netWtKg: 66.0,
      grossWtKg: 74.0,
    },
    {
      requestedDescription: 'Isolation Transformer 100W 6.6A/6.6A Series Circuit',
      offeredDescription: 'Eaton Crouse-Hinds 100W 6.6A/6.6A Encapsulated Isolation Transformer',
      uom: 'PCS',
      qty: 24,
      supplierUnitCost: 72,
      boxes: 3,
      boxLengthCm: 50,
      boxWidthCm: 40,
      boxHeightCm: 32,
      kgPerBox: 22.0,
      cartonNo: 'CTN 05–07',
      netWtKg: 61.2,
      grossWtKg: 66.0,
    },
    {
      requestedDescription: 'Primary Airport Lighting Cable 5kV 1x6mm2 Copper',
      offeredDescription: 'Prysmian FAA L-824 Type C 5kV 1x6mm2 XLPE Shielded Cable (500m Drum)',
      uom: 'ROLL',
      qty: 4,
      supplierUnitCost: 940,
      boxes: 4,
      boxLengthCm: 85,
      boxWidthCm: 85,
      boxHeightCm: 62,
      kgPerBox: 88.0,
      cartonNo: 'DRUM 08–11',
      netWtKg: 328.0,
      grossWtKg: 352.0,
    },
    {
      requestedDescription: 'Waterproof Primary Connector Kit L-823 Style 2/9',
      offeredDescription: 'Amerace L-823 Molded Primary Plug & Receptacle Kit (54-D4-D4)',
      uom: 'SET',
      qty: 40,
      supplierUnitCost: 28.5,
      boxes: 2,
      boxLengthCm: 42,
      boxWidthCm: 34,
      boxHeightCm: 28,
      kgPerBox: 11.5,
      cartonNo: 'CTN 12–13',
      netWtKg: 20.4,
      grossWtKg: 23.0,
    },
  ];

  seededRows.forEach((data, idx) => {
    job.items[idx] = { ...job.items[idx], ...data };
  });

  return job;
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

export function normalizeJob(job: ProcurementJob): ProcurementJob {
  return {
    ...job,
    customColumns: Array.isArray(job.customColumns) ? job.customColumns : [],
    controls: {
      ...job.controls,
      authorizedSignatory:
        typeof job.controls.authorizedSignatory === 'string'
          ? job.controls.authorizedSignatory
          : 'Authorized Signature — GRAND GROUP FZC',
      counterpartySignatory:
        typeof job.controls.counterpartySignatory === 'string'
          ? job.controls.counterpartySignatory
          : 'Customer / Supplier Acknowledgment',
    },
    items: (job.items || createEmptyItems()).map((it, idx) => ({
      ...it,
      rowNumber: it.rowNumber || 17 + idx,
      customFields: it.customFields || {},
    })),
  };
}

export function loadJobs(): ProcurementJob[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.JOBS);
    if (!raw) return [createSampleJob()];
    const parsed = JSON.parse(raw) as ProcurementJob[];
    if (!Array.isArray(parsed) || parsed.length === 0) return [createSampleJob()];
    return parsed.map(normalizeJob);
  } catch {
    return [createSampleJob()];
  }
}

export function saveJobs(jobs: ProcurementJob[]): void {
  try {
    localStorage.setItem(STORAGE_KEYS.JOBS, JSON.stringify(jobs));
  } catch {
    // Ignore quota errors gracefully
  }
}

export function loadRegistry(): RegistryEntry[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.REGISTRY);
    if (!raw) return createSeededRegistry();
    const parsed = JSON.parse(raw) as RegistryEntry[];
    if (!Array.isArray(parsed) || parsed.length === 0) return createSeededRegistry();
    return parsed;
  } catch {
    return createSeededRegistry();
  }
}

export function saveRegistry(entries: RegistryEntry[]): void {
  try {
    localStorage.setItem(STORAGE_KEYS.REGISTRY, JSON.stringify(entries));
  } catch {
    // Ignore quota errors gracefully
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
    // Ignore
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
    // Ignore
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
