import type { JobControls, JobItemRow, ProcurementJob, WorkspaceTab } from '../types/procurement';

/**
 * Per-page undo. Each page keeps its own history, and undo restores only the fields that page edits.
 * Pages that are not listed here (dashboard, registry, settings) fall back to restoring the whole job.
 */
interface PageDomain {
  items: (keyof JobItemRow)[];
  jobKeys: (keyof ProcurementJob)[];
  controlKeys: (keyof JobControls)[];
}

const DOC_CONTROL_KEYS: (keyof JobControls)[] = [
  'endUserSelect', 'customEndUser', 'endUserContact', 'endUserAddress', 'rfqRef', 'customerPoRef',
  'paymentTerms', 'deliveryLeadTime', 'shipDate', 'portOfLoading', 'deliveryDate', 'transportRef',
  'receivedBy', 'remarks', 'authorizedSignatory', 'counterpartySignatory', 'currency', 'incoterm',
  'destination', 'vatRate', 'includeVat',
];

const DOC_ITEM_KEYS: (keyof JobItemRow)[] = [
  'requestedDescription', 'offeredDescription', 'uom', 'qty', 'customFields', 'cartonNo', 'netWtKg', 'grossWtKg',
];

const DOMAINS: Record<string, PageDomain> = {
  PRICE_CALC: {
    items: ['supplierUnitCost', 'boxes', 'boxLengthCm', 'boxWidthCm', 'boxHeightCm', 'kgPerBox'],
    jobKeys: ['freight'],
    controlKeys: ['vatRate', 'includeVat'],
  },
  INTERNAL_QUOTE: {
    items: ['requestedDescription', 'offeredDescription', 'uom', 'qty', 'customFields', 'supplierUnitCost'],
    jobKeys: ['supplierQuote'],
    controlKeys: ['supplierName', 'supplierAddress', 'supplierRef', 'paymentTerms', 'deliveryLeadTime', 'currency', 'incoterm', 'destination'],
  },
};

/** Document pages: QT, PI, PO, CI, PL, DN each have their own history */
const docDomain: PageDomain = { items: DOC_ITEM_KEYS, jobKeys: [], controlKeys: DOC_CONTROL_KEYS };

export function undoPageKey(tab: WorkspaceTab): string {
  return ['QT', 'PI', 'PO', 'CI', 'PL', 'DN'].includes(tab) ? `DOC-${tab}` : tab;
}

function domainFor(tab: WorkspaceTab): PageDomain | null {
  if (['QT', 'PI', 'PO', 'CI', 'PL', 'DN'].includes(tab)) return docDomain;
  return DOMAINS[tab] ?? null;
}

/** Returns the job with only this page's fields replaced by the snapshot's values. Issued records are kept. */
export function applyPageSnapshot(current: ProcurementJob, snapshot: ProcurementJob, tab: WorkspaceTab): ProcurementJob {
  const domain = domainFor(tab);
  if (!domain) return { ...snapshot, issuedDocs: current.issuedDocs };

  const items = current.items.map((it) => {
    const before = snapshot.items.find((s) => s.rowNumber === it.rowNumber);
    if (!before) return it;
    const patch: Record<string, unknown> = {};
    domain.items.forEach((k) => {
      patch[k] = before[k];
    });
    return { ...it, ...patch };
  });

  const jobPatch: Record<string, unknown> = {};
  domain.jobKeys.forEach((k) => {
    jobPatch[k] = snapshot[k];
  });

  const controls: Record<string, unknown> = { ...current.controls };
  domain.controlKeys.forEach((k) => {
    controls[k] = snapshot.controls[k];
  });

  return { ...current, ...jobPatch, items, controls: controls as unknown as JobControls };
}
